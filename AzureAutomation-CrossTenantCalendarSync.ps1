#Requires -Version 7.4
<#
.SYNOPSIS
    Azure Automation runbook for two-way calendar mirroring between
    joey@dzidrums.com and jdzidrums@ultrapro.com.

.DESCRIPTION
    Production-oriented cross-tenant Exchange Online calendar mirroring using
    Microsoft Graph, Exchange Online RBAC for Applications, Azure Key Vault,
    Azure Storage, and the Azure Automation account system-assigned managed
    identity.

    The runbook intentionally uses only PowerShell core cmdlets at runtime. It
    obtains managed-identity tokens directly from Azure Automation's identity
    endpoint, reads the two Graph client secrets from Key Vault, and stores its
    delta/mapping state in Azure Blob Storage.

    Mirrored events contain no attendees, no meeting body, no Teams join data,
    and no attachments. This prevents duplicate meeting invitations and limits
    cross-tenant information exposure.

    A finite Azure Blob lease provides a distributed lock so staggered Azure
    Automation schedules cannot modify the state concurrently.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$KeyVaultName,

    [string]$DzidrumsSecretName = 'dzidrums-calendar-sync-client-secret',
    [string]$UltraProSecretName = 'ultrapro-calendar-sync-client-secret',

    [Parameter(Mandatory)]
    [string]$StateStorageAccount,

    [string]$StateContainer = 'calendar-sync',
    [string]$StateBlob = 'calendar-sync-state.json',
    [string]$LockBlob = 'calendar-sync.lock',

    [Parameter(Mandatory)]
    [string]$DzidrumsTenantId,

    [Parameter(Mandatory)]
    [string]$DzidrumsClientId,

    [string]$DzidrumsMailbox = 'joey@dzidrums.com',

    [Parameter(Mandatory)]
    [string]$UltraProTenantId,

    [Parameter(Mandatory)]
    [string]$UltraProClientId,

    [string]$UltraProMailbox = 'jdzidrums@ultrapro.com',

    [ValidateSet('SubjectLocation','BusyOnly')]
    [string]$DetailMode = 'SubjectLocation',

    [bool]$RespectPrivate = $true,
    [bool]$CopyReminders = $false,

    [ValidateRange(0,3650)]
    [int]$PastDays = 30,

    [ValidateRange(1,3650)]
    [int]$FutureDays = 365,

    [ValidateRange(1,30)]
    [int]$RebaselineDays = 1,

    [ValidateRange(1,5000)]
    [int]$MaxDeltaPages = 1000,

    [ValidateSet('Sync','Test','Rebaseline')]
    [string]$Mode = 'Sync'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:SyncTagPrefix = 'CTCSYNC-v2-'
$script:StorageApiVersion = '2023-11-03'
$script:LeaseId = $null
$script:LeaseLastRenewalUtc = [datetime]::MinValue
$script:StorageToken = $null
$script:StorageTokenExpiresUtc = [datetime]::MinValue
$script:KeyVaultToken = $null
$script:KeyVaultTokenExpiresUtc = [datetime]::MinValue
$script:StateWasNew = $false

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO'
    )

    $stamp = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
    $line = "[$stamp] [$Level] $Message"
    if ($Level -eq 'WARN') {
        Write-Warning $line
    }
    else {
        Write-Output $line
    }
}

function Get-ExceptionStatusCode {
    param([Parameter(Mandatory)]$ErrorRecord)

    try {
        if ($ErrorRecord.Exception.Response.StatusCode) {
            return [int]$ErrorRecord.Exception.Response.StatusCode
        }
    }
    catch {}

    return $null
}

function Get-ManagedIdentityToken {
    param([Parameter(Mandatory)][string]$Resource)

    if ([string]::IsNullOrWhiteSpace($env:IDENTITY_ENDPOINT) -or
        [string]::IsNullOrWhiteSpace($env:IDENTITY_HEADER)) {
        throw 'Azure Automation managed-identity endpoint variables are unavailable. Enable a system-assigned managed identity on the Automation account.'
    }

    $headers = @{
        'X-IDENTITY-HEADER' = $env:IDENTITY_HEADER
        'Metadata'          = 'True'
    }

    $response = Invoke-RestMethod `
        -Method POST `
        -Uri $env:IDENTITY_ENDPOINT `
        -Headers $headers `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body @{ resource = $Resource }

    if (-not $response.access_token) {
        throw "Managed identity did not return an access token for $Resource."
    }

    [pscustomobject]@{
        Token      = [string]$response.access_token
        ExpiresUtc = if ($response.expires_on) {
            try { [DateTimeOffset]::FromUnixTimeSeconds([int64]$response.expires_on).UtcDateTime }
            catch { [datetime]::UtcNow.AddMinutes(45) }
        }
        else { [datetime]::UtcNow.AddMinutes(45) }
    }
}

function Get-CachedManagedIdentityToken {
    param([Parameter(Mandatory)][ValidateSet('Storage','KeyVault')][string]$For)

    $now = [datetime]::UtcNow

    if ($For -eq 'Storage') {
        if (-not $script:StorageToken -or $script:StorageTokenExpiresUtc -lt $now.AddMinutes(5)) {
            $result = Get-ManagedIdentityToken -Resource 'https://storage.azure.com/'
            $script:StorageToken = $result.Token
            $script:StorageTokenExpiresUtc = $result.ExpiresUtc
        }
        return $script:StorageToken
    }

    if (-not $script:KeyVaultToken -or $script:KeyVaultTokenExpiresUtc -lt $now.AddMinutes(5)) {
        $result = Get-ManagedIdentityToken -Resource 'https://vault.azure.net'
        $script:KeyVaultToken = $result.Token
        $script:KeyVaultTokenExpiresUtc = $result.ExpiresUtc
    }
    return $script:KeyVaultToken
}

function Get-KeyVaultSecretText {
    param(
        [Parameter(Mandatory)][string]$VaultName,
        [Parameter(Mandatory)][string]$SecretName
    )

    $token = Get-CachedManagedIdentityToken -For KeyVault
    $safeName = [Uri]::EscapeDataString($SecretName)
    $uri = "https://$VaultName.vault.azure.net/secrets/$($safeName)?api-version=7.4"

    $response = Invoke-RestMethod -Method GET -Uri $uri -Headers @{ Authorization = "Bearer $token" }
    if ([string]::IsNullOrWhiteSpace([string]$response.value)) {
        throw "Key Vault secret '$SecretName' is missing or empty."
    }
    return [string]$response.value
}

function New-StorageHeaders {
    param([hashtable]$Additional)

    $token = Get-CachedManagedIdentityToken -For Storage
    $headers = @{
        Authorization  = "Bearer $token"
        'x-ms-version' = $script:StorageApiVersion
        'x-ms-date'    = [datetime]::UtcNow.ToString('R')
    }

    if ($Additional) {
        foreach ($key in $Additional.Keys) {
            $headers[$key] = $Additional[$key]
        }
    }

    return $headers
}

function Get-StorageContainerUri {
    "https://$StateStorageAccount.blob.core.windows.net/$StateContainer"
}

function Get-StorageBlobUri {
    param([Parameter(Mandatory)][string]$BlobName)

    $segments = $BlobName -split '/' | ForEach-Object { [Uri]::EscapeDataString($_) }
    "$(Get-StorageContainerUri)/$($segments -join '/')"
}

function Ensure-StateContainer {
    $uri = "$(Get-StorageContainerUri)?restype=container"
    try {
        Invoke-WebRequest -Method PUT -Uri $uri -Headers (New-StorageHeaders) -UseBasicParsing | Out-Null
        Write-Log "Created state container '$StateContainer'."
    }
    catch {
        $status = Get-ExceptionStatusCode $_
        if ($status -ne 409) { throw }
    }
}

function Ensure-LockBlob {
    $uri = Get-StorageBlobUri -BlobName $LockBlob
    $headers = New-StorageHeaders -Additional @{
        'x-ms-blob-type' = 'BlockBlob'
        'If-None-Match'  = '*'
    }

    try {
        Invoke-WebRequest -Method PUT -Uri $uri -Headers $headers -Body ([byte[]]@()) -UseBasicParsing | Out-Null
    }
    catch {
        $status = Get-ExceptionStatusCode $_
        if ($status -notin @(409,412)) { throw }
    }
}

function Acquire-StorageLease {
    Ensure-LockBlob

    $leaseId = [guid]::NewGuid().ToString()
    $uri = "$(Get-StorageBlobUri -BlobName $LockBlob)?comp=lease"
    $headers = New-StorageHeaders -Additional @{
        'x-ms-lease-action'      = 'acquire'
        'x-ms-lease-duration'    = '60'
        'x-ms-proposed-lease-id' = $leaseId
    }

    try {
        Invoke-WebRequest -Method PUT -Uri $uri -Headers $headers -UseBasicParsing | Out-Null
        $script:LeaseId = $leaseId
        $script:LeaseLastRenewalUtc = [datetime]::UtcNow
        return $true
    }
    catch {
        $status = Get-ExceptionStatusCode $_
        if ($status -in @(409,412)) {
            return $false
        }
        throw
    }
}

function Renew-StorageLeaseIfNeeded {
    if (-not $script:LeaseId) { return }
    if ($script:LeaseLastRenewalUtc -gt [datetime]::UtcNow.AddSeconds(-25)) { return }

    $uri = "$(Get-StorageBlobUri -BlobName $LockBlob)?comp=lease"
    $headers = New-StorageHeaders -Additional @{
        'x-ms-lease-action' = 'renew'
        'x-ms-lease-id'     = $script:LeaseId
    }

    try {
        Invoke-WebRequest -Method PUT -Uri $uri -Headers $headers -UseBasicParsing | Out-Null
        $script:LeaseLastRenewalUtc = [datetime]::UtcNow
    }
    catch {
        throw 'The distributed synchronization lock could not be renewed. Stopping to prevent concurrent state writes.'
    }
}

function Release-StorageLease {
    if (-not $script:LeaseId) { return }

    $uri = "$(Get-StorageBlobUri -BlobName $LockBlob)?comp=lease"
    $headers = New-StorageHeaders -Additional @{
        'x-ms-lease-action' = 'release'
        'x-ms-lease-id'     = $script:LeaseId
    }

    try {
        Invoke-WebRequest -Method PUT -Uri $uri -Headers $headers -UseBasicParsing | Out-Null
        Write-Log 'Released distributed synchronization lock.'
    }
    catch {
        Write-Log 'Could not explicitly release the lock; its finite lease will expire automatically.' 'WARN'
    }
    finally {
        $script:LeaseId = $null
    }
}

function Get-NewState {
    @{
        Version = 2
        AtoB = @{
            DeltaLink       = $null
            WindowStartUtc  = $null
            WindowEndUtc    = $null
            LastBaselineUtc = $null
            Mappings        = @{}
        }
        BtoA = @{
            DeltaLink       = $null
            WindowStartUtc  = $null
            WindowEndUtc    = $null
            LastBaselineUtc = $null
            Mappings        = @{}
        }
    }
}

function Get-State {
    $uri = Get-StorageBlobUri -BlobName $StateBlob
    try {
        $response = Invoke-WebRequest -Method GET -Uri $uri -Headers (New-StorageHeaders) -UseBasicParsing
        $state = ([string]$response.Content) | ConvertFrom-Json -AsHashtable -Depth 100
        foreach ($direction in @('AtoB','BtoA')) {
            if (-not $state.ContainsKey($direction)) { $state[$direction] = @{} }
            if (-not $state[$direction].ContainsKey('Mappings') -or $null -eq $state[$direction].Mappings) {
                $state[$direction].Mappings = @{}
            }
        }
        return $state
    }
    catch {
        $status = Get-ExceptionStatusCode $_
        if ($status -eq 404) {
            $script:StateWasNew = $true
            return (Get-NewState)
        }
        throw
    }
}

function Set-State {
    param([Parameter(Mandatory)]$State)

    Renew-StorageLeaseIfNeeded
    $uri = Get-StorageBlobUri -BlobName $StateBlob
    $json = $State | ConvertTo-Json -Depth 100 -Compress
    $headers = New-StorageHeaders -Additional @{ 'x-ms-blob-type' = 'BlockBlob' }
    Invoke-WebRequest -Method PUT -Uri $uri -Headers $headers -ContentType 'application/json; charset=utf-8' -Body $json -UseBasicParsing | Out-Null
}

function Get-GraphAccessToken {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret
    )

    $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $response = Invoke-RestMethod `
        -Method POST `
        -Uri $tokenUri `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body @{
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = 'https://graph.microsoft.com/.default'
            grant_type    = 'client_credentials'
        }

    if (-not $response.access_token) {
        throw "Graph token acquisition failed for tenant $TenantId."
    }
    return [string]$response.access_token
}

function Invoke-Graph {
    param(
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][ValidateSet('GET','POST','PATCH','DELETE')][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [object]$Body,
        [int]$MaxAttempts = 6
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        Renew-StorageLeaseIfNeeded

        $headers = @{
            Authorization = "Bearer $Token"
            Accept        = 'application/json'
            Prefer        = 'IdType="ImmutableId"'
        }

        try {
            $params = @{
                Method      = $Method
                Uri         = $Uri
                Headers     = $headers
                ErrorAction = 'Stop'
            }
            if ($null -ne $Body) {
                $params.ContentType = 'application/json'
                $params.Body = $Body | ConvertTo-Json -Depth 30 -Compress
            }
            return Invoke-RestMethod @params
        }
        catch {
            $status = Get-ExceptionStatusCode $_
            $retryable = $status -in @(429,500,502,503,504)
            if (-not $retryable -or $attempt -eq $MaxAttempts) {
                $detail = $_.ErrorDetails.Message
                if ([string]::IsNullOrWhiteSpace($detail)) { $detail = $_.Exception.Message }
                throw "Graph $Method failed (HTTP $status): $detail"
            }

            $delay = [Math]::Min([Math]::Pow(2, $attempt), 30)
            $warningStamp = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
            Write-Warning "[$warningStamp] [WARN] Graph returned HTTP $status; retrying with backoff."
            Renew-StorageLeaseIfNeeded
            Start-Sleep -Seconds $delay
        }
    }
}

function New-DeltaUri {
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][datetime]$StartUtc,
        [Parameter(Mandatory)][datetime]$EndUtc
    )

    $m = [Uri]::EscapeDataString($Mailbox)
    $s = [Uri]::EscapeDataString($StartUtc.ToString('o'))
    $e = [Uri]::EscapeDataString($EndUtc.ToString('o'))
    # calendarView/delta does not support $select; Graph returns the event
    # properties for the view and carries the original query in delta links.
    "https://graph.microsoft.com/v1.0/users/$m/calendarView/delta?startDateTime=$s&endDateTime=$e"
}

function New-CalendarViewUri {
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][datetime]$StartUtc,
        [Parameter(Mandatory)][datetime]$EndUtc
    )

    $m = [Uri]::EscapeDataString($Mailbox)
    $s = [Uri]::EscapeDataString($StartUtc.ToString('o'))
    $e = [Uri]::EscapeDataString($EndUtc.ToString('o'))
    $select = 'id,transactionId,start,end'
    "https://graph.microsoft.com/v1.0/users/$m/calendarView?startDateTime=$s&endDateTime=$e&`$select=$select&`$top=1000"
}

function Get-EventUri {
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][string]$EventId
    )

    $m = [Uri]::EscapeDataString($Mailbox)
    $e = [Uri]::EscapeDataString($EventId)
    "https://graph.microsoft.com/v1.0/users/$m/events/$e"
}

function Get-CreateEventUri {
    param([Parameter(Mandatory)][string]$Mailbox)
    $m = [Uri]::EscapeDataString($Mailbox)
    "https://graph.microsoft.com/v1.0/users/$m/events"
}

function Invoke-DeltaCollection {
    param(
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][string]$InitialUri
    )

    $items = [System.Collections.Generic.List[object]]::new()
    $uri = $InitialUri
    $deltaLink = $null
    $page = 0
    $seenLinks = [System.Collections.Generic.HashSet[string]]::new()

    while ($uri) {
        Renew-StorageLeaseIfNeeded
        $page++
        if ($page -gt $MaxDeltaPages) {
            throw "Delta query exceeded MaxDeltaPages ($MaxDeltaPages)."
        }
        if (-not $seenLinks.Add($uri)) {
            throw 'Graph returned a repeated delta continuation URL.'
        }

        $response = Invoke-Graph -Token $Token -Method GET -Uri $uri
        foreach ($item in @($response.value)) { $items.Add($item) }

        $next = $response.PSObject.Properties['@odata.nextLink']
        $delta = $response.PSObject.Properties['@odata.deltaLink']

        if ($next -and $next.Value) {
            $uri = [string]$next.Value
        }
        else {
            $uri = $null
            if ($delta -and $delta.Value) { $deltaLink = [string]$delta.Value }
        }
    }

    if (-not $deltaLink) {
        throw 'Delta query completed without an @odata.deltaLink.'
    }

    [pscustomobject]@{ Items = $items; DeltaLink = $deltaLink }
}

function Get-MirrorIndex {
    param(
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][datetime]$StartUtc,
        [Parameter(Mandatory)][datetime]$EndUtc
    )

    $index = @{}
    $uri = New-CalendarViewUri -Mailbox $Mailbox -StartUtc $StartUtc -EndUtc $EndUtc
    $page = 0

    while ($uri) {
        Renew-StorageLeaseIfNeeded
        $page++
        if ($page -gt $MaxDeltaPages) { throw 'Destination mirror-index query exceeded MaxDeltaPages.' }

        $response = Invoke-Graph -Token $Token -Method GET -Uri $uri
        foreach ($event in @($response.value)) {
            $tx = [string]$event.transactionId
            if (-not [string]::IsNullOrWhiteSpace($tx) -and $tx.StartsWith($script:SyncTagPrefix)) {
                if (-not $index.ContainsKey($tx)) {
                    $index[$tx] = [string]$event.id
                }
            }
        }

        $next = $response.PSObject.Properties['@odata.nextLink']
        $uri = if ($next -and $next.Value) { [string]$next.Value } else { $null }
    }

    return $index
}

function Get-DeterministicTransactionId {
    param(
        [Parameter(Mandatory)][string]$SourceMailbox,
        [Parameter(Mandatory)][string]$SourceEventId
    )

    $bytes = [Text.Encoding]::UTF8.GetBytes("$SourceMailbox|$SourceEventId")
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $hash = $sha.ComputeHash($bytes) }
    finally { $sha.Dispose() }
    $hex = [Convert]::ToHexString($hash).ToLowerInvariant()
    "$script:SyncTagPrefix$($hex.Substring(0,32))"
}

function Test-IsMirrorEvent {
    param([Parameter(Mandatory)]$Event)

    $tx = [string]$Event.transactionId
    return (-not [string]::IsNullOrWhiteSpace($tx) -and $tx.StartsWith($script:SyncTagPrefix))
}

function Convert-SourceToMirrorPayload {
    param(
        [Parameter(Mandatory)]$Event,
        [Parameter(Mandatory)][string]$SourceLabel,
        [Parameter(Mandatory)][string]$MirrorPrefix
    )

    $private = ([string]$Event.sensitivity -eq 'private')

    if ($RespectPrivate -and $private) {
        $subject = 'Private - Busy'
        $location = ''
    }
    elseif ($DetailMode -eq 'BusyOnly') {
        $subject = "$SourceLabel - Busy"
        $location = ''
    }
    else {
        $subject = "$MirrorPrefix$([string]$Event.subject)"
        if ([string]::IsNullOrWhiteSpace($subject)) { $subject = "$SourceLabel - Busy" }
        $location = if ($Event.location -and $Event.location.displayName) { [string]$Event.location.displayName } else { '' }
    }

    $payload = [ordered]@{
        subject     = $subject
        start       = [ordered]@{ dateTime = [string]$Event.start.dateTime; timeZone = [string]$Event.start.timeZone }
        end         = [ordered]@{ dateTime = [string]$Event.end.dateTime; timeZone = [string]$Event.end.timeZone }
        isAllDay    = [bool]$Event.isAllDay
        showAs      = if ($Event.showAs) { [string]$Event.showAs } else { 'busy' }
        sensitivity = if ($Event.sensitivity) { [string]$Event.sensitivity } else { 'normal' }
        location    = [ordered]@{ displayName = $location }
        isReminderOn = $false
    }

    if ($CopyReminders) {
        $payload.isReminderOn = [bool]$Event.isReminderOn
        if ($null -ne $Event.reminderMinutesBeforeStart) {
            $payload.reminderMinutesBeforeStart = [int]$Event.reminderMinutesBeforeStart
        }
    }

    return $payload
}

function Convert-ToUtcDate {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    try { return ([datetimeoffset]::Parse($Value)).UtcDateTime }
    catch { return $null }
}

function New-MirrorEvent {
    param(
        [Parameter(Mandatory)][string]$DestinationToken,
        [Parameter(Mandatory)][string]$DestinationMailbox,
        [Parameter(Mandatory)]$SourceEvent,
        [Parameter(Mandatory)][string]$SourceLabel,
        [Parameter(Mandatory)][string]$MirrorPrefix,
        [Parameter(Mandatory)][string]$TransactionId
    )

    $payload = Convert-SourceToMirrorPayload -Event $SourceEvent -SourceLabel $SourceLabel -MirrorPrefix $MirrorPrefix
    $payload.transactionId = $TransactionId
    Invoke-Graph -Token $DestinationToken -Method POST -Uri (Get-CreateEventUri $DestinationMailbox) -Body $payload
}

function Update-MirrorEvent {
    param(
        [Parameter(Mandatory)][string]$DestinationToken,
        [Parameter(Mandatory)][string]$DestinationMailbox,
        [Parameter(Mandatory)][string]$DestinationEventId,
        [Parameter(Mandatory)]$SourceEvent,
        [Parameter(Mandatory)][string]$SourceLabel,
        [Parameter(Mandatory)][string]$MirrorPrefix
    )

    $payload = Convert-SourceToMirrorPayload -Event $SourceEvent -SourceLabel $SourceLabel -MirrorPrefix $MirrorPrefix
    Invoke-Graph -Token $DestinationToken -Method PATCH -Uri (Get-EventUri $DestinationMailbox $DestinationEventId) -Body $payload
}

function Remove-MirrorEvent {
    param(
        [Parameter(Mandatory)][string]$DestinationToken,
        [Parameter(Mandatory)][string]$DestinationMailbox,
        [Parameter(Mandatory)][string]$DestinationEventId
    )

    try {
        Invoke-Graph -Token $DestinationToken -Method DELETE -Uri (Get-EventUri $DestinationMailbox $DestinationEventId) | Out-Null
    }
    catch {
        if ($_.Exception.Message -match 'HTTP 404') { return }
        throw
    }
}

function Invoke-OneDirection {
    param(
        [Parameter(Mandatory)][string]$DirectionName,
        [Parameter(Mandatory)]$DirectionState,
        [Parameter(Mandatory)][string]$SourceMailbox,
        [Parameter(Mandatory)][string]$SourceLabel,
        [Parameter(Mandatory)][string]$MirrorPrefix,
        [Parameter(Mandatory)][string]$DestinationMailbox,
        [Parameter(Mandatory)][string]$SourceToken,
        [Parameter(Mandatory)][string]$DestinationToken,
        [bool]$ForceBaseline = $false
    )

    $now = [datetime]::UtcNow
    $needsBaseline = $ForceBaseline -or [string]::IsNullOrWhiteSpace([string]$DirectionState.DeltaLink)

    if (-not $needsBaseline -and $DirectionState.LastBaselineUtc) {
        try {
            $lastBaseline = [datetimeoffset]::Parse([string]$DirectionState.LastBaselineUtc).UtcDateTime
            if ($lastBaseline -lt $now.AddDays(-$RebaselineDays)) { $needsBaseline = $true }
        }
        catch { $needsBaseline = $true }
    }

    if ($needsBaseline) {
        $windowStart = $now.AddDays(-$PastDays)
        $windowEnd = $now.AddDays($FutureDays)
        $deltaUri = New-DeltaUri -Mailbox $SourceMailbox -StartUtc $windowStart -EndUtc $windowEnd
        $mirrorIndex = Get-MirrorIndex -Token $DestinationToken -Mailbox $DestinationMailbox -StartUtc $windowStart -EndUtc $windowEnd
        Write-Log "$DirectionName baseline started. Existing mirror count in window: $($mirrorIndex.Count)."
    }
    else {
        $deltaUri = [string]$DirectionState.DeltaLink
        $windowStart = [datetimeoffset]::Parse([string]$DirectionState.WindowStartUtc).UtcDateTime
        $windowEnd = [datetimeoffset]::Parse([string]$DirectionState.WindowEndUtc).UtcDateTime
        $mirrorIndex = @{}
        Write-Log "$DirectionName incremental delta started."
    }

    $delta = Invoke-DeltaCollection -Token $SourceToken -InitialUri $deltaUri
    $seenSourceIds = [System.Collections.Generic.HashSet[string]]::new()
    $createdCount = 0
    $updatedCount = 0
    $deletedCount = 0
    $ignoredMirrorCount = 0

    foreach ($event in $delta.Items) {
        Renew-StorageLeaseIfNeeded
        $sourceEventId = [string]$event.id
        if ([string]::IsNullOrWhiteSpace($sourceEventId)) { continue }

        $isRemoved = $null -ne $event.PSObject.Properties['@removed']
        if ($isRemoved) {
            if ($DirectionState.Mappings.ContainsKey($sourceEventId)) {
                $destinationId = [string]$DirectionState.Mappings[$sourceEventId].DestinationEventId
                Remove-MirrorEvent -DestinationToken $DestinationToken -DestinationMailbox $DestinationMailbox -DestinationEventId $destinationId
                [void]$DirectionState.Mappings.Remove($sourceEventId)
                $deletedCount++
            }
            continue
        }

        if (Test-IsMirrorEvent $event) {
            $ignoredMirrorCount++
            continue
        }

        [void]$seenSourceIds.Add($sourceEventId)

        if ([bool]$event.isCancelled) {
            if ($DirectionState.Mappings.ContainsKey($sourceEventId)) {
                $destinationId = [string]$DirectionState.Mappings[$sourceEventId].DestinationEventId
                Remove-MirrorEvent -DestinationToken $DestinationToken -DestinationMailbox $DestinationMailbox -DestinationEventId $destinationId
                [void]$DirectionState.Mappings.Remove($sourceEventId)
                $deletedCount++
            }
            continue
        }

        $transactionId = Get-DeterministicTransactionId -SourceMailbox $SourceMailbox -SourceEventId $sourceEventId
        $startUtc = $null
        $endUtc = $null
        try { $startUtc = ([datetimeoffset]::Parse([string]$event.start.dateTime)).UtcDateTime } catch {}
        try { $endUtc = ([datetimeoffset]::Parse([string]$event.end.dateTime)).UtcDateTime } catch {}

        if (-not $DirectionState.Mappings.ContainsKey($sourceEventId) -and $mirrorIndex.ContainsKey($transactionId)) {
            $DirectionState.Mappings[$sourceEventId] = @{
                DestinationEventId = [string]$mirrorIndex[$transactionId]
                TransactionId      = $transactionId
                SourceStartUtc     = if ($startUtc) { $startUtc.ToString('o') } else { $null }
                SourceEndUtc       = if ($endUtc) { $endUtc.ToString('o') } else { $null }
                LastSeenUtc        = $now.ToString('o')
            }
        }

        if ($DirectionState.Mappings.ContainsKey($sourceEventId)) {
            $mapping = $DirectionState.Mappings[$sourceEventId]
            $destinationId = [string]$mapping.DestinationEventId
            try {
                Update-MirrorEvent `
                    -DestinationToken $DestinationToken `
                    -DestinationMailbox $DestinationMailbox `
                    -DestinationEventId $destinationId `
                    -SourceEvent $event `
                    -SourceLabel $SourceLabel `
                    -MirrorPrefix $MirrorPrefix | Out-Null
                $updatedCount++
            }
            catch {
                if ($_.Exception.Message -match 'HTTP 404') {
                    $created = New-MirrorEvent `
                        -DestinationToken $DestinationToken `
                        -DestinationMailbox $DestinationMailbox `
                        -SourceEvent $event `
                        -SourceLabel $SourceLabel `
                        -MirrorPrefix $MirrorPrefix `
                        -TransactionId $transactionId
                    $mapping.DestinationEventId = [string]$created.id
                    $createdCount++
                }
                else { throw }
            }

            $mapping.TransactionId = $transactionId
            $mapping.SourceStartUtc = if ($startUtc) { $startUtc.ToString('o') } else { $null }
            $mapping.SourceEndUtc = if ($endUtc) { $endUtc.ToString('o') } else { $null }
            $mapping.LastSeenUtc = $now.ToString('o')
        }
        else {
            $created = New-MirrorEvent `
                -DestinationToken $DestinationToken `
                -DestinationMailbox $DestinationMailbox `
                -SourceEvent $event `
                -SourceLabel $SourceLabel `
                -MirrorPrefix $MirrorPrefix `
                -TransactionId $transactionId

            $DirectionState.Mappings[$sourceEventId] = @{
                DestinationEventId = [string]$created.id
                TransactionId      = $transactionId
                SourceStartUtc     = if ($startUtc) { $startUtc.ToString('o') } else { $null }
                SourceEndUtc       = if ($endUtc) { $endUtc.ToString('o') } else { $null }
                LastSeenUtc        = $now.ToString('o')
            }
            $createdCount++
        }
    }

    if ($needsBaseline) {
        foreach ($sourceId in @($DirectionState.Mappings.Keys)) {
            $mapping = $DirectionState.Mappings[$sourceId]
            $startUtc = Convert-ToUtcDate $mapping.SourceStartUtc
            $endUtc = Convert-ToUtcDate $mapping.SourceEndUtc
            $overlaps = $false
            if ($startUtc -and $endUtc) {
                $overlaps = ($endUtc -ge $windowStart -and $startUtc -le $windowEnd)
            }

            if ($overlaps -and -not $seenSourceIds.Contains($sourceId)) {
                Remove-MirrorEvent `
                    -DestinationToken $DestinationToken `
                    -DestinationMailbox $DestinationMailbox `
                    -DestinationEventId ([string]$mapping.DestinationEventId)
                [void]$DirectionState.Mappings.Remove($sourceId)
                $deletedCount++
            }
            elseif ($endUtc -and $endUtc -lt $windowStart.AddDays(-7)) {
                [void]$DirectionState.Mappings.Remove($sourceId)
            }
        }
    }

    $DirectionState.DeltaLink = $delta.DeltaLink
    if ($needsBaseline) {
        $DirectionState.WindowStartUtc = $windowStart.ToString('o')
        $DirectionState.WindowEndUtc = $windowEnd.ToString('o')
        $DirectionState.LastBaselineUtc = $now.ToString('o')
    }

    Write-Log "$DirectionName completed: created=$createdCount updated=$updatedCount deleted=$deletedCount ignoredMirrors=$ignoredMirrorCount."
}

# -------------------- Main --------------------

Write-Log "Runbook started in mode '$Mode'."
Ensure-StateContainer

$leaseAcquired = Acquire-StorageLease
if (-not $leaseAcquired) {
    Write-Log 'Another synchronization job currently owns the lock. This run will exit without changes.' 'WARN'
    return
}
Write-Log 'Acquired distributed synchronization lock.'

try {
    $dzSecret = Get-KeyVaultSecretText -VaultName $KeyVaultName -SecretName $DzidrumsSecretName
    $upSecret = Get-KeyVaultSecretText -VaultName $KeyVaultName -SecretName $UltraProSecretName

    Write-Log 'Acquiring Microsoft Graph application tokens.'
    $dzToken = Get-GraphAccessToken -TenantId $DzidrumsTenantId -ClientId $DzidrumsClientId -ClientSecret $dzSecret
    $upToken = Get-GraphAccessToken -TenantId $UltraProTenantId -ClientId $UltraProClientId -ClientSecret $upSecret

    # Avoid keeping plaintext client secrets around for the remainder of the run.
    $dzSecret = $null
    $upSecret = $null

    if ($Mode -eq 'Test') {
        foreach ($test in @(
            @{ Name = 'Dzidrums'; Mailbox = $DzidrumsMailbox; Token = $dzToken },
            @{ Name = 'Ultra PRO'; Mailbox = $UltraProMailbox; Token = $upToken }
        )) {
            $mailbox = [Uri]::EscapeDataString([string]$test.Mailbox)
            $calendar = Invoke-Graph -Token $test.Token -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$mailbox/calendar?`$select=id,name"
            Write-Log "$($test.Name) calendar authorization is valid for $($test.Mailbox)."
        }
        Write-Log 'TEST PASSED.'
        return
    }

    $state = Get-State
    if ($script:StateWasNew) {
        Write-Log 'No state blob exists yet; a baseline synchronization will be performed.'
    }
    $forceBaseline = ($Mode -eq 'Rebaseline')
    if ($forceBaseline) {
        $state.AtoB.DeltaLink = $null
        $state.BtoA.DeltaLink = $null
        Write-Log 'Rebaseline requested; existing event mappings are being retained.'
    }

    Invoke-OneDirection `
        -DirectionName 'Dzidrums -> Ultra PRO' `
        -DirectionState $state.AtoB `
        -SourceMailbox $DzidrumsMailbox `
        -SourceLabel 'Personal' `
        -MirrorPrefix '[Personal] ' `
        -DestinationMailbox $UltraProMailbox `
        -SourceToken $dzToken `
        -DestinationToken $upToken `
        -ForceBaseline $forceBaseline

    Invoke-OneDirection `
        -DirectionName 'Ultra PRO -> Dzidrums' `
        -DirectionState $state.BtoA `
        -SourceMailbox $UltraProMailbox `
        -SourceLabel 'Work' `
        -MirrorPrefix '[Work] ' `
        -DestinationMailbox $DzidrumsMailbox `
        -SourceToken $upToken `
        -DestinationToken $dzToken `
        -ForceBaseline $forceBaseline

    Set-State -State $state
    Write-Log 'Synchronization completed successfully and state was persisted to Azure Blob Storage.'
}
finally {
    Release-StorageLease
}
