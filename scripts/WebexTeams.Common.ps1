# Shared functions for the three CSV pipeline scripts. Dot-source from a script in this folder.

$script:WebexUrlPattern = '(?i)https?://(?:[a-z0-9-]+\.)*webex\.(?:com|eu)/[^\s<>"'']+'

function Initialize-WebexEwsApi {
    param([Parameter(Mandatory = $true)][string] $ManagedApiDll)

    if (-not (Test-Path -LiteralPath $ManagedApiDll -PathType Leaf)) {
        throw "EWS Managed API assembly not found: $ManagedApiDll"
    }
    if (-not ('Microsoft.Exchange.WebServices.Data.ExchangeService' -as [type])) {
        Add-Type -Path $ManagedApiDll
    }
}

function ConvertTo-UtcIsoString {
    param([Parameter(Mandatory = $true)][datetime] $Value)

    if ($Value.Kind -eq [DateTimeKind]::Unspecified) {
        $utc = [DateTime]::SpecifyKind($Value, [DateTimeKind]::Utc)
    }
    elseif ($Value.Kind -eq [DateTimeKind]::Local) {
        $utc = $Value.ToUniversalTime()
    }
    else {
        $utc = $Value
    }
    return $utc.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'")
}

function New-EwsService {
    param(
        [Parameter(Mandatory = $true)][string] $MailboxSmtp,
        [Parameter(Mandatory = $true)][uri] $Endpoint,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential] $Credential,
        [Parameter(Mandatory = $true)][ValidateSet('System', 'Direct', 'SystemDefaultCredentials')][string] $ProxyMode
    )

    $networkCredential = $Credential.GetNetworkCredential()
    $service = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new(
        [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1,
        [TimeZoneInfo]::Utc
    )
    if ($ProxyMode -eq 'Direct') {
        $service.WebProxy = [System.Net.WebProxy]::new()
    }
    elseif ($ProxyMode -eq 'SystemDefaultCredentials') {
        $systemProxy = [System.Net.WebRequest]::DefaultWebProxy
        if ($null -ne $systemProxy) {
            $systemProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
            $service.WebProxy = $systemProxy
        }
    }
    $service.Url = $Endpoint
    $service.Credentials = [Microsoft.Exchange.WebServices.Data.WebCredentials]::new(
        $networkCredential.UserName,
        $networkCredential.Password,
        $networkCredential.Domain
    )
    $service.ImpersonatedUserId = [Microsoft.Exchange.WebServices.Data.ImpersonatedUserId]::new(
        [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,
        $MailboxSmtp
    )
    $service.HttpHeaders.Add('X-AnchorMailbox', $MailboxSmtp)
    return $service
}

function Resolve-EwsCredential {
    param(
        [System.Management.Automation.PSCredential] $Credential,
        [Parameter(Mandatory = $true)][string] $CredentialPath
    )

    if ($env:OS -ne 'Windows_NT') {
        throw 'The encrypted EWS credential cache requires Windows DPAPI. Run this script on Windows PowerShell on Windows.'
    }
    if ($Credential) { return $Credential }
    if ([string]::IsNullOrWhiteSpace($CredentialPath)) { throw 'EwsCredentialPath cannot be empty.' }

    $fullPath = [System.IO.Path]::GetFullPath($CredentialPath)
    if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
        try { $saved = Import-Clixml -LiteralPath $fullPath -ErrorAction Stop }
        catch {
            throw "Could not decrypt the saved EWS credential at '$fullPath'. Use the same Windows user and computer, or remove the file to enter it again. Details: $($_.Exception.Message)"
        }
        if ($saved -isnot [System.Management.Automation.PSCredential]) {
            throw "The saved file '$fullPath' is not a PowerShell credential. Remove it and rerun to create a new cache."
        }
        return $saved
    }

    $newCredential = Get-Credential -Message 'Enter the EWS service account credential. It will be saved encrypted for this Windows account and computer.'
    if (-not $newCredential) { throw 'No EWS credential was supplied.' }
    $directory = Split-Path -Path $fullPath -Parent
    if ($directory -and -not (Test-Path -LiteralPath $directory -PathType Container)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    try { $newCredential | Export-Clixml -Path $fullPath -Force -ErrorAction Stop }
    catch { throw "Could not save the encrypted EWS credential to '$fullPath'. Details: $($_.Exception.Message)" }
    Write-Host "Encrypted EWS credential saved to: $fullPath"
    return $newCredential
}

function New-EwsFindPropertySet {
    $props = [Microsoft.Exchange.WebServices.Data.PropertySet]::new(
        [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly
    )
    @(
        [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::IsMeeting,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::IsCancelled,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::AppointmentType,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::MyResponseType,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::ICalUid
    ) | ForEach-Object { $props.Add($_) }
    return ,$props
}

function New-EwsDetailPropertySet {
    $props = New-EwsFindPropertySet
    @(
        [Microsoft.Exchange.WebServices.Data.ItemSchema]::Body,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Location,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::IsRecurring,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::ModifiedOccurrences
    ) | ForEach-Object { $props.Add($_) }
    return ,$props
}

function Get-EwsCalendarItemsInRange {
    param(
        [Parameter(Mandatory = $true)] $Service,
        [Parameter(Mandatory = $true)][datetime] $StartUtc,
        [Parameter(Mandatory = $true)][datetime] $EndUtc,
        [Parameter(Mandatory = $true)] $PropertySet,
        [timespan] $MinimumSlice = ([timespan]::FromHours(1))
    )

    $view = [Microsoft.Exchange.WebServices.Data.CalendarView]::new($StartUtc, $EndUtc, 1000)
    $view.PropertySet = $PropertySet
    $result = $Service.FindAppointments([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, $view)
    if (-not $result.MoreAvailable) {
        foreach ($item in $result.Items) { Write-Output $item }
        return
    }

    $span = $EndUtc - $StartUtc
    if ($span -le $MinimumSlice) {
        throw "EWS returned more than 1,000 calendar items in a $($span.TotalMinutes)-minute range. The mailbox report is incomplete for this range."
    }
    $midpoint = $StartUtc.AddTicks([long]($span.Ticks / 2))
    foreach ($item in @(Get-EwsCalendarItemsInRange -Service $Service -StartUtc $StartUtc -EndUtc $midpoint -PropertySet $PropertySet -MinimumSlice $MinimumSlice)) {
        Write-Output $item
    }
    foreach ($item in @(Get-EwsCalendarItemsInRange -Service $Service -StartUtc $midpoint -EndUtc $EndUtc -PropertySet $PropertySet -MinimumSlice $MinimumSlice)) {
        Write-Output $item
    }
}

function Get-WebexMtidValue {
    param([Parameter(Mandatory = $true)][string] $Url)
    $decoded = [System.Net.WebUtility]::HtmlDecode($Url)
    $match = [regex]::Match($decoded, '(?i)(?:[?&])MTID=([^&;#\s]+)')
    if (-not $match.Success) {
        $match = [regex]::Match($Url, '(?i)(?:[?&](?:amp;)?|%3f|%26)MTID(?:=|%3d)([^&;#\s]+)')
    }
    if (-not $match.Success) { return '' }
    $value = [System.Net.WebUtility]::HtmlDecode([string]$match.Groups[1].Value)
    try { $value = [System.Uri]::UnescapeDataString($value) } catch { }
    return [string]$value.Trim().TrimEnd('.', ',', ';', ':', '!', '?', ')', ']')
}

function Select-WebexJoinUrls {
    param(
        [Parameter(Mandatory = $true)][string] $BodyText,
        [Parameter(Mandatory = $true)][string] $LocationText
    )

    $candidates = New-Object System.Collections.Generic.List[object]
    foreach ($field in @(
        [pscustomobject]@{ Source = 'Body'; Text = $BodyText },
        [pscustomobject]@{ Source = 'Location'; Text = $LocationText }
    )) {
        foreach ($match in [regex]::Matches([string]$field.Text, $script:WebexUrlPattern)) {
            $url = [string]$match.Value.TrimEnd('.', ',', ';', ':', '!', '?', ')', ']')
            if (-not [string]::IsNullOrWhiteSpace($url)) {
                $candidates.Add([pscustomobject]@{ Url = $url; Source = $field.Source })
            }
        }
    }

    $diagnostics = New-Object System.Collections.Generic.List[string]
    $number = 0
    foreach ($candidate in $candidates) {
        $number++
        $mtid = Get-WebexMtidValue -Url $candidate.Url
        if ([string]::IsNullOrWhiteSpace($mtid)) { $mtid = '<none>' }
        $kind = if ($candidate.Url -match '(?i)/j\.php(?:[?#]|$)') { 'JoinUrl' } else { 'Other' }
        $diagnostics.Add(('#{0} Kind={1}; Source={2}; MTID={3}; URL={4}' -f $number, $kind, $candidate.Source, $mtid, $candidate.Url))
    }

    $map = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::Ordinal)
    foreach ($candidate in $candidates) {
        if ($map.ContainsKey($candidate.Url)) {
            if (-not $map[$candidate.Url].Sources.Contains($candidate.Source)) { $map[$candidate.Url].Sources.Add($candidate.Source) }
        }
        else {
            $sources = New-Object System.Collections.Generic.List[string]
            $sources.Add($candidate.Source)
            $map[$candidate.Url] = [pscustomobject]@{
                Url = $candidate.Url
                Sources = $sources
                MtidValue = Get-WebexMtidValue -Url $candidate.Url
                IsJoinUrl = $candidate.Url -match '(?i)/j\.php(?:[?#]|$)'
            }
        }
    }

    $unique = @($map.Values)
    if ($unique.Count -eq 0) {
        return [pscustomobject]@{ Status = 'SkippedNoWebex'; Note = 'No Webex URL found in body or location.'; OriginalWebexUrl = ''; WebexUrlSource = ''; WebexUrlDiagnostics = ($diagnostics -join ' || '); SelectedWebexUrlsJson = '[]' }
    }
    $joinPool = @($unique | Where-Object { $_.IsJoinUrl -and -not [string]::IsNullOrWhiteSpace([string]$_.MtidValue) })
    if ($joinPool.Count -gt 0) { $pool = $joinPool; $poolDescription = 'j.php join URLs' }
    else { $pool = @($unique | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.MtidValue) }); $poolDescription = 'Webex URLs' }

    $mtids = New-Object System.Collections.Generic.List[string]
    foreach ($url in $pool) {
        if (-not [string]::IsNullOrWhiteSpace([string]$url.MtidValue) -and -not $mtids.Contains([string]$url.MtidValue)) {
            $mtids.Add([string]$url.MtidValue)
        }
    }
    if ($mtids.Count -gt 1) {
        return [pscustomobject]@{
            Status = 'SkippedMultipleMtidWebex'
            Note = "Found $($mtids.Count) different MTID values in $poolDescription; manual review required. See WebexUrlDiagnostics."
            OriginalWebexUrl = ''
            WebexUrlSource = ''
            WebexUrlDiagnostics = ($diagnostics -join ' || ')
            SelectedWebexUrlsJson = '[]'
        }
    }
    if ($mtids.Count -eq 1) {
        $selected = @($pool | Where-Object { $_.MtidValue -ceq $mtids[0] })
    }
    elseif ($unique.Count -eq 1) { $selected = @($unique[0]) }
    else {
        return [pscustomobject]@{
            Status = 'SkippedMultipleWebex'
            Note = "Found $($unique.Count) distinct Webex URLs and none has an MTID; manual review required. See WebexUrlDiagnostics."
            OriginalWebexUrl = ''
            WebexUrlSource = ''
            WebexUrlDiagnostics = ($diagnostics -join ' || ')
            SelectedWebexUrlsJson = '[]'
        }
    }

    $selectedUrls = @($selected | ForEach-Object { [string]$_.Url } | Sort-Object Length -Descending)
    $selectedSources = New-Object System.Collections.Generic.List[string]
    foreach ($selectedUrl in $selected) {
        foreach ($source in $selectedUrl.Sources) {
            if (-not $selectedSources.Contains($source)) { $selectedSources.Add($source) }
        }
    }
    $note = ''
    if ($joinPool.Count -gt 0) { $note += ' Selected the scheduled j.php join-link group.' }
    if ($selected.Count -gt 1) { $note += " Found $($selected.Count) URL variants sharing the same MTID; all selected variants will be replaced." }
    $otherCount = $unique.Count - $selected.Count
    if ($otherCount -gt 0) { $note += " Left $otherCount other Webex URL(s) outside the selected group untouched." }
    $json = ConvertTo-Json -InputObject ([object[]]$selectedUrls) -Compress

    return [pscustomobject]@{
        Status = 'ReadyForTeamsLink'
        Note = $note
        OriginalWebexUrl = [string]$selected[0].Url
        WebexUrlSource = ($selectedSources.ToArray() -join '+')
        WebexUrlDiagnostics = ($diagnostics -join ' || ')
        SelectedWebexUrlsJson = $json
    }
}

function Get-SafeUserFileStem {
    param([Parameter(Mandatory = $true)][string] $MailboxSmtp)
    $stem = [regex]::Replace($MailboxSmtp.Trim(), '[^A-Za-z0-9._@+-]', '_').TrimEnd(' ', '.')
    if ([string]::IsNullOrWhiteSpace($stem)) { throw "Cannot create a filename for mailbox '$MailboxSmtp'." }
    return $stem
}

function Get-TeamsExternalId {
    param(
        [Parameter(Mandatory = $true)][string] $MailboxSmtp,
        [Parameter(Mandatory = $true)][string] $ICalUid,
        [Parameter(Mandatory = $true)][string] $IdentityStartUtc
    )
    $text = '{0}|{1}|{2}' -f $MailboxSmtp.ToLowerInvariant(), $ICalUid, $IdentityStartUtc
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = [BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($text))).Replace('-', '').ToLowerInvariant()
        return $hash.Substring(0, 32)
    }
    finally { $sha.Dispose() }
}

function Get-TeamsJoinUrl {
    param(
        [Parameter(Mandatory = $true)][string] $EntraObjectId,
        [Parameter(Mandatory = $true)][string] $ExternalId,
        [Parameter(Mandatory = $true)][string] $Subject,
        [Parameter(Mandatory = $true)][string] $StartUtc,
        [Parameter(Mandatory = $true)][string] $EndUtc
    )
    $escapedId = [Uri]::EscapeDataString($EntraObjectId)
    $uri = "https://graph.microsoft.com/v1.0/users/$escapedId/onlineMeetings/createOrGet"
    $payload = @{
        externalId = $ExternalId
        subject = $Subject
        startDateTime = $StartUtc
        endDateTime = $EndUtc
    } | ConvertTo-Json -Depth 5
    $response = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $payload -ContentType 'application/json'
    if ([string]::IsNullOrWhiteSpace([string]$response.joinWebUrl)) {
        throw "Graph returned no joinWebUrl for organizer object ID $EntraObjectId."
    }
    return [string]$response.joinWebUrl
}

function Resolve-GraphClientSecretCredential {
    param([Parameter(Mandatory = $true)][string] $AppId, [Parameter(Mandatory = $true)][string] $CredentialPath)
    if ($env:OS -ne 'Windows_NT') {
        throw 'The encrypted Graph client secret cache requires Windows DPAPI. Run this script on Windows PowerShell on Windows.'
    }
    if ([string]::IsNullOrWhiteSpace($CredentialPath)) { throw 'GraphClientSecretPath cannot be empty.' }
    $fullPath = [System.IO.Path]::GetFullPath($CredentialPath)
    if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
        try { $saved = Import-Clixml -LiteralPath $fullPath -ErrorAction Stop }
        catch { throw "Could not decrypt the saved Graph secret at '$fullPath'. Use the same Windows user/computer, or remove it to re-enter the secret. Details: $($_.Exception.Message)" }
        if ($saved -isnot [System.Management.Automation.PSCredential]) { throw "The saved Graph cache '$fullPath' is not a PowerShell credential. Remove it and rerun." }
        if ($saved.UserName -ne $AppId) { throw "The saved Graph cache '$fullPath' belongs to a different app ID. Remove it or choose another path." }
        return [pscustomobject]@{ Credential = $saved; NeedsSave = $false; Path = $fullPath }
    }
    $secureSecret = Read-Host -Prompt 'Enter the Graph app client secret. It will be encrypted after successful Graph authentication.' -AsSecureString
    if ($null -eq $secureSecret -or $secureSecret.Length -eq 0) { throw 'A Graph client secret is required.' }
    return [pscustomobject]@{
        Credential = [System.Management.Automation.PSCredential]::new($AppId, $secureSecret)
        NeedsSave = $true
        Path = $fullPath
    }
}

function Save-GraphClientSecretCredential {
    param([Parameter(Mandatory = $true)][System.Management.Automation.PSCredential] $Credential, [Parameter(Mandatory = $true)][string] $CredentialPath)
    $directory = Split-Path -Path $CredentialPath -Parent
    if ($directory -and -not (Test-Path -LiteralPath $directory -PathType Container)) { New-Item -ItemType Directory -Path $directory -Force | Out-Null }
    try { $Credential | Export-Clixml -Path $CredentialPath -Force -ErrorAction Stop }
    catch { throw "Could not save the encrypted Graph secret to '$CredentialPath'. Details: $($_.Exception.Message)" }
    Write-Host "Encrypted Graph client secret saved to: $CredentialPath"
}

function Connect-WebexTeamsGraph {
    param(
        [Parameter(Mandatory = $true)][string] $TenantId,
        [Parameter(Mandatory = $true)][string] $AppId,
        [Parameter(Mandatory = $true)][ValidateSet('Certificate', 'ClientSecret')][string] $AuthMethod,
        [string] $CertificateThumbprint,
        [Parameter(Mandatory = $true)][string] $ClientSecretPath
    )
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    if ($AuthMethod -eq 'Certificate') {
        if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) { throw 'Certificate auth requires a Graph certificate thumbprint.' }
        Connect-MgGraph -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $CertificateThumbprint -ContextScope Process -NoWelcome
        return
    }
    $resolved = Resolve-GraphClientSecretCredential -AppId $AppId -CredentialPath $ClientSecretPath
    try {
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $resolved.Credential -ContextScope Process -NoWelcome
    }
    catch {
        throw "Could not connect to Graph with the client secret. If it was rotated, remove the encrypted cache at '$($resolved.Path)' and rerun. Details: $($_.Exception.Message)"
    }
    if ($resolved.NeedsSave) { Save-GraphClientSecretCredential -Credential $resolved.Credential -CredentialPath $resolved.Path }
}

function Disconnect-WebexTeamsGraph {
    if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
        if (Get-MgContext) { Disconnect-MgGraph | Out-Null }
    }
}

function Get-NewBodyText {
    param(
        [Parameter(Mandatory = $true)][string] $OriginalBody,
        [Parameter(Mandatory = $true)][string[]] $OriginalUrls,
        [Parameter(Mandatory = $true)][string] $TeamsUrl,
        [Parameter(Mandatory = $true)][Microsoft.Exchange.WebServices.Data.BodyType] $BodyType
    )
    if ($BodyType -eq [Microsoft.Exchange.WebServices.Data.BodyType]::HTML) {
        $safe = [System.Security.SecurityElement]::Escape($TeamsUrl)
        $updated = $OriginalBody
        foreach ($url in $OriginalUrls) { $updated = $updated.Replace($url, $safe) }
        $banner = '<div><p><strong>Join Microsoft Teams meeting:</strong> <a href="{0}">Join the meeting now</a></p></div><hr />' -f $safe
        return $banner + $updated
    }
    $updated = $OriginalBody
    foreach ($url in $OriginalUrls) { $updated = $updated.Replace($url, $TeamsUrl) }
    return "Join Microsoft Teams meeting: $TeamsUrl`r`n`r`n$updated"
}

function New-BodyValue {
    param(
        [Parameter(Mandatory = $true)][Microsoft.Exchange.WebServices.Data.BodyType] $BodyType,
        [Parameter(Mandatory = $true)][string] $Text
    )
    return [Microsoft.Exchange.WebServices.Data.MessageBody]::new($BodyType, $Text)
}

function Write-JsonLine {
    param([Parameter(Mandatory = $true)] $Object, [Parameter(Mandatory = $true)][string] $Path)
    Add-Content -LiteralPath $Path -Value ($Object | ConvertTo-Json -Depth 8 -Compress) -Encoding UTF8
}

function Initialize-OutputDirectory {
    param([Parameter(Mandatory = $true)][string] $Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}

function Assert-OutputFileAvailable {
    param([Parameter(Mandatory = $true)][string] $Path, [switch] $Force)
    if ((Test-Path -LiteralPath $Path -PathType Leaf) -and -not $Force) {
        throw "Output file already exists: $Path. Choose another output directory or rerun with -Force to replace it."
    }
}
