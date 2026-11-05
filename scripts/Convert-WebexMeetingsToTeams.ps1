[CmdletBinding()]
param(
    [ValidateSet('Preview', 'Apply', 'Rollback')]
    [string] $Mode = 'Preview',

    [Parameter(Mandatory = $true)]
    [string] $EwsUrl,

    [ValidateSet('System', 'Direct', 'SystemDefaultCredentials')]
    [string] $EwsProxyMode = 'Direct',

    [Parameter(Mandatory = $true)]
    [string] $EwsManagedApiDll,

    [System.Management.Automation.PSCredential] $EwsCredential,

    [string] $EwsCredentialPath = (Join-Path $env:LOCALAPPDATA 'WebexTeamsPOC\ews-credential.xml'),

    [string] $TargetsCsv,
    [datetimeoffset] $WindowStartUtc,
    [datetimeoffset] $WindowEndUtc,

    [string] $GraphTenantId,
    [string] $GraphAppId,
    [string] $GraphCertificateThumbprint,

    [string] $BackupPath = ('.\teams-poc-backup-{0}.jsonl' -f (Get-Date -Format 'yyyyMMdd-HHmmss')),
    [string] $ReportPath = ('.\teams-poc-report-{0}.csv' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
)

$ErrorActionPreference = 'Stop'
$webexPattern = '(?i)https?://(?:[a-z0-9-]+\.)*webex\.(?:com|eu)/[^\s<>"'']+'

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
        # An empty WebProxy has no address and bypasses the system proxy.
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
    if ($Credential) {
        return $Credential
    }
    if ([string]::IsNullOrWhiteSpace($CredentialPath)) {
        throw 'EwsCredentialPath cannot be empty.'
    }

    $credentialFullPath = [System.IO.Path]::GetFullPath($CredentialPath)
    if (Test-Path -LiteralPath $credentialFullPath -PathType Leaf) {
        try {
            $savedCredential = Import-Clixml -LiteralPath $credentialFullPath -ErrorAction Stop
        }
        catch {
            throw "Could not decrypt the saved EWS credential at '$credentialFullPath'. Run as the same Windows user on the same computer that created it, or remove the file and rerun to save a replacement. Details: $($_.Exception.Message)"
        }
        if ($savedCredential -isnot [System.Management.Automation.PSCredential]) {
            throw "The saved file '$credentialFullPath' does not contain a PowerShell credential. Remove it and rerun the script to create a new encrypted credential file."
        }
        return $savedCredential
    }

    $newCredential = Get-Credential -Message 'Enter the EWS service account credential. It will be saved encrypted for this Windows account and computer.'
    if (-not $newCredential) {
        throw 'No EWS credential was supplied.'
    }

    $credentialDirectory = Split-Path -Path $credentialFullPath -Parent
    if ($credentialDirectory -and -not (Test-Path -LiteralPath $credentialDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $credentialDirectory -Force | Out-Null
    }
    try {
        $newCredential | Export-Clixml -Path $credentialFullPath -Force -ErrorAction Stop
    }
    catch {
        throw "Could not save the encrypted EWS credential to '$credentialFullPath'. Details: $($_.Exception.Message)"
    }

    Write-Host "Encrypted EWS credential saved to: $credentialFullPath"
    return $newCredential
}

function New-TeamsExternalId {
    param(
        [Parameter(Mandatory = $true)][string] $MailboxSmtp,
        [Parameter(Mandatory = $true)][string] $ICalUid,
        [Parameter(Mandatory = $true)][datetime] $Start
    )

    $inputText = '{0}|{1}|{2}' -f $MailboxSmtp.ToLowerInvariant(), $ICalUid, (ConvertTo-UtcIsoString -Value $Start)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($inputText)
        $hash = [BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '').ToLowerInvariant()
        return $hash.Substring(0, 32)
    }
    finally {
        $sha.Dispose()
    }
}

function Get-TeamsJoinUrl {
    param(
        [Parameter(Mandatory = $true)][string] $EntraObjectId,
        [Parameter(Mandatory = $true)][string] $ExternalId,
        [Parameter(Mandatory = $true)][string] $Subject,
        [Parameter(Mandatory = $true)][datetime] $Start,
        [Parameter(Mandatory = $true)][datetime] $End
    )

    $escapedUserId = [Uri]::EscapeDataString($EntraObjectId)
    $uri = "https://graph.microsoft.com/v1.0/users/$escapedUserId/onlineMeetings/createOrGet"
    $payload = @{
        externalId    = $ExternalId
        subject       = $Subject
        startDateTime = ConvertTo-UtcIsoString -Value $Start
        endDateTime   = ConvertTo-UtcIsoString -Value $End
    } | ConvertTo-Json -Depth 6

    $response = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $payload -ContentType 'application/json'
    if ([string]::IsNullOrWhiteSpace([string]$response.joinWebUrl)) {
        throw "Graph returned no joinWebUrl for organizer object ID $EntraObjectId."
    }
    return [string]$response.joinWebUrl
}

function Get-NewBodyText {
    param(
        [Parameter(Mandatory = $true)][string] $OriginalBody,
        [Parameter(Mandatory = $true)][string[]] $OriginalUrls,
        [Parameter(Mandatory = $true)][string] $TeamsUrl,
        [Parameter(Mandatory = $true)][Microsoft.Exchange.WebServices.Data.BodyType] $BodyType
    )

    if ($BodyType -eq [Microsoft.Exchange.WebServices.Data.BodyType]::HTML) {
        $safeTeamsUrl = [System.Security.SecurityElement]::Escape($TeamsUrl)
        $updated = $OriginalBody
        foreach ($originalUrl in $OriginalUrls) {
            $updated = $updated.Replace($originalUrl, $safeTeamsUrl)
        }
        $banner = '<div><p><strong>Join Microsoft Teams meeting:</strong> <a href="{0}">Join the meeting now</a></p></div><hr />' -f $safeTeamsUrl
        return $banner + $updated
    }

    $updated = $OriginalBody
    foreach ($originalUrl in $OriginalUrls) {
        $updated = $updated.Replace($originalUrl, $TeamsUrl)
    }
    return "Join Microsoft Teams meeting: $TeamsUrl`r`n`r`n$updated"
}

function Get-WebexMtidValue {
    param([Parameter(Mandatory = $true)][string] $Url)

    $decodedUrl = [System.Net.WebUtility]::HtmlDecode($Url)
    $mtidMatch = [regex]::Match($decodedUrl, '(?i)(?:[?&])MTID=([^&;#\s]+)')
    if (-not $mtidMatch.Success) {
        $mtidMatch = [regex]::Match($Url, '(?i)(?:[?&](?:amp;)?|%3f|%26)MTID(?:=|%3d)([^&;#\s]+)')
    }
    if ($mtidMatch.Success) {
        $mtidValue = [System.Net.WebUtility]::HtmlDecode([string]$mtidMatch.Groups[1].Value)
        try {
            $mtidValue = [System.Uri]::UnescapeDataString($mtidValue)
        }
        catch {
            # Keep the extracted value as-is if it contains malformed percent encoding.
        }
        return [string]$mtidValue.Trim().TrimEnd('.', ',', ';', ':', '!', '?', ')', ']')
    }
    return ''
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
    $line = $Object | ConvertTo-Json -Depth 8 -Compress
    Add-Content -LiteralPath $Path -Value $line -Encoding UTF8
}

function New-DetailPropertySet {
    $props = [Microsoft.Exchange.WebServices.Data.PropertySet]::new(
        [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly
    )
    @(
        [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject,
        [Microsoft.Exchange.WebServices.Data.ItemSchema]::Body,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Location,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::IsMeeting,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::IsCancelled,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::IsRecurring,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::AppointmentType,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::MyResponseType,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::ICalUid,
        [Microsoft.Exchange.WebServices.Data.AppointmentSchema]::ModifiedOccurrences
    ) | ForEach-Object { $props.Add($_) }
    # PropertySet implements IEnumerable; the unary comma prevents PowerShell from unrolling it.
    return ,$props
}

if (-not (Test-Path -LiteralPath $EwsManagedApiDll -PathType Leaf)) {
    throw "EWS Managed API assembly not found: $EwsManagedApiDll"
}
Add-Type -Path $EwsManagedApiDll

$endpointUri = [uri]$EwsUrl
if ($endpointUri.Scheme -ne 'https') {
    throw 'EWS URL must use HTTPS.'
}
$EwsCredential = Resolve-EwsCredential -Credential $EwsCredential -CredentialPath $EwsCredentialPath

$detailProps = New-DetailPropertySet
$report = New-Object System.Collections.Generic.List[object]

if ($Mode -eq 'Rollback') {
    if (-not $BackupPath -or -not (Test-Path -LiteralPath $BackupPath -PathType Leaf)) {
        throw "Rollback backup file not found: $BackupPath"
    }

    $records = Get-Content -LiteralPath $BackupPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_ | ConvertFrom-Json }
    foreach ($record in $records) {
        try {
            $service = New-EwsService -MailboxSmtp $record.MailboxSmtp -Endpoint $endpointUri `
                -Credential $EwsCredential -ProxyMode $EwsProxyMode
            $itemId = [Microsoft.Exchange.WebServices.Data.ItemId]::new([string]$record.ItemId)
            $appointment = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service, $itemId)
            $appointment.Load($detailProps)
            $locationWasBackedUp = $null -ne $record.PSObject.Properties['NewLocation']
            $bodyMatchesBackup = $appointment.Body.Text -ceq [string]$record.NewBodyText
            $locationMatchesBackup = (-not $locationWasBackedUp) -or ($appointment.Location -ceq [string]$record.NewLocation)
            if (-not $bodyMatchesBackup -or -not $locationMatchesBackup) {
                $report.Add([pscustomobject]@{
                    MailboxSmtp = $record.MailboxSmtp; Subject = $record.Subject; StartUtc = $record.StartUtc
                    Status = 'RollbackSkippedChanged'; Note = 'Current body or location no longer matches this script backup.'
                })
                continue
            }

            $originalType = [Microsoft.Exchange.WebServices.Data.BodyType][Enum]::Parse(
                [Microsoft.Exchange.WebServices.Data.BodyType], [string]$record.OriginalBodyType, $true
            )
            $appointment.Body = New-BodyValue -BodyType $originalType -Text ([string]$record.OriginalBodyText)
            if ($null -ne $record.PSObject.Properties['OriginalLocation']) {
                $appointment.Location = [string]$record.OriginalLocation
            }
            $appointment.Update(
                [Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite,
                [Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode]::SendToAllAndSaveCopy
            )
            $report.Add([pscustomobject]@{
                MailboxSmtp = $record.MailboxSmtp; Subject = $record.Subject; StartUtc = $record.StartUtc
                Status = 'RolledBack'; Note = 'Original body and location restored; attendee update sent.'
            })
        }
        catch {
            $report.Add([pscustomobject]@{
                MailboxSmtp = $record.MailboxSmtp; Subject = $record.Subject; StartUtc = $record.StartUtc
                Status = 'RollbackError'; Note = $_.Exception.Message
            })
        }
    }
}
else {
    if (-not $TargetsCsv -or -not (Test-Path -LiteralPath $TargetsCsv -PathType Leaf)) {
        throw "Targets CSV not found: $TargetsCsv"
    }
    if ($WindowStartUtc -eq [datetimeoffset]::MinValue -or
        $WindowEndUtc -eq [datetimeoffset]::MinValue -or
        $WindowEndUtc -le $WindowStartUtc) {
        throw 'Provide a valid -WindowStartUtc and -WindowEndUtc. Use UTC timestamps (for example, 2026-11-05T00:00:00Z).'
    }

    $targets = @(Import-Csv -LiteralPath $TargetsCsv)
    if ($targets.Count -eq 0) { throw 'Targets CSV is empty.' }
    foreach ($target in $targets) {
        if (-not $target.MailboxSmtp -or -not $target.EntraObjectId) {
            throw 'Each targets.csv row must have MailboxSmtp and EntraObjectId values.'
        }
    }

    if ($Mode -eq 'Apply') {
        if (-not $GraphTenantId -or -not $GraphAppId -or -not $GraphCertificateThumbprint) {
            throw 'Apply requires GraphTenantId, GraphAppId, and GraphCertificateThumbprint.'
        }
        if (Test-Path -LiteralPath $BackupPath) {
            throw "Backup path already exists. Choose a new path so the rollback source is unambiguous: $BackupPath"
        }
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        Connect-MgGraph -TenantId $GraphTenantId -ClientId $GraphAppId `
            -CertificateThumbprint $GraphCertificateThumbprint -ContextScope Process -NoWelcome
    }

    $findProps = [Microsoft.Exchange.WebServices.Data.PropertySet]::new(
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
    ) | ForEach-Object { $findProps.Add($_) }

    $startUtc = $WindowStartUtc.UtcDateTime
    $endUtc = $WindowEndUtc.UtcDateTime
    $view = [Microsoft.Exchange.WebServices.Data.CalendarView]::new($startUtc, $endUtc, 1000)
    $view.PropertySet = $findProps

    foreach ($target in $targets) {
        $mailbox = ([string]$target.MailboxSmtp).Trim()
        $entraObjectId = ([string]$target.EntraObjectId).Trim()
        try {
            $service = New-EwsService -MailboxSmtp $mailbox -Endpoint $endpointUri `
                -Credential $EwsCredential -ProxyMode $EwsProxyMode
            $items = $service.FindAppointments(
                [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,
                $view
            )
            $processedRecurringSeries = @{}

            foreach ($candidate in $items.Items) {
                $subject = [string]$candidate.Subject
                $startLabel = ConvertTo-UtcIsoString -Value $candidate.Start
                $base = @{
                    MailboxSmtp = $mailbox
                    Subject = $subject
                    StartUtc = $startLabel
                    MeetingScope = 'Single'
                    HasModifiedOccurrences = $false
                    OriginalWebexUrl = ''
                    WebexUrlSource = ''
                    WebexUrlDiagnostics = ''
                    TeamsJoinUrl = ''
                    Status = ''
                    Note = ''
                }

                if (-not $candidate.IsMeeting) {
                    $base.Status = 'SkippedNotMeeting'; $base.Note = 'Calendar item is not a meeting.'
                    $report.Add([pscustomobject]$base); continue
                }
                if ($candidate.IsCancelled) {
                    $base.Status = 'SkippedCancelled'; $base.Note = 'Meeting is cancelled.'
                    $report.Add([pscustomobject]$base); continue
                }
                if ([string]$candidate.MyResponseType -ne 'Organizer') {
                    $base.Status = 'SkippedNotOrganizer'; $base.Note = 'Not organizer.'
                    $report.Add([pscustomobject]$base); continue
                }

                $appointmentType = [string]$candidate.AppointmentType
                $seriesLookupKey = ''
                if ($appointmentType -in @('Occurrence', 'Exception', 'RecurringMaster') -and $candidate.ICalUid) {
                    $seriesLookupKey = '{0}|{1}' -f $mailbox.ToLowerInvariant(), [string]$candidate.ICalUid
                    if ($processedRecurringSeries.ContainsKey($seriesLookupKey)) {
                        $base.MeetingScope = 'RecurringSeries'
                        $base.Status = 'SkippedDuplicateSeriesOccurrence'
                        $base.Note = 'This recurring series was already processed from another occurrence in the window.'
                        $report.Add([pscustomobject]$base); continue
                    }
                }
                $meetingStart = $candidate.Start
                $meetingEnd = $candidate.End
                $identityStart = $candidate.Start

                if ($appointmentType -eq 'Single') {
                    $appointment = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service, $candidate.Id)
                    $appointment.Load($detailProps)
                }
                elseif ($appointmentType -in @('Occurrence', 'Exception')) {
                    try {
                        # CalendarView expands recurring series into occurrences/exceptions.
                        # Bind to the series master so one body update covers the whole series.
                        $appointment = [Microsoft.Exchange.WebServices.Data.Appointment]::BindToRecurringMaster($service, $candidate.Id)
                        $appointment.Load($detailProps)
                    }
                    catch {
                        $base.MeetingScope = 'RecurringSeries'
                        $base.Status = 'RecurringMasterBindError'
                        $base.Note = $_.Exception.Message
                        $report.Add([pscustomobject]$base); continue
                    }

                    $base.MeetingScope = 'RecurringSeries'
                    $identityStart = $appointment.Start
                    if (-not $seriesLookupKey) {
                        $seriesLookupKey = '{0}|{1}' -f $mailbox.ToLowerInvariant(), [string]$appointment.ICalUid
                    }
                    if ($processedRecurringSeries.ContainsKey($seriesLookupKey)) {
                        $base.Status = 'SkippedDuplicateSeriesOccurrence'
                        $base.Note = 'This recurring series was already processed from another occurrence in the window.'
                        $report.Add([pscustomobject]$base); continue
                    }
                    $processedRecurringSeries[$seriesLookupKey] = $true

                    if ($appointment.ModifiedOccurrences -and $appointment.ModifiedOccurrences.Count -gt 0) {
                        $base.HasModifiedOccurrences = $true
                    }
                }
                elseif ($appointmentType -eq 'RecurringMaster') {
                    $appointment = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service, $candidate.Id)
                    $appointment.Load($detailProps)
                    $base.MeetingScope = 'RecurringSeries'
                    $identityStart = $appointment.Start
                    if (-not $seriesLookupKey) {
                        $seriesLookupKey = '{0}|{1}' -f $mailbox.ToLowerInvariant(), [string]$appointment.ICalUid
                    }
                    if ($processedRecurringSeries.ContainsKey($seriesLookupKey)) {
                        $base.Status = 'SkippedDuplicateSeriesOccurrence'
                        $base.Note = 'This recurring series was already processed from another occurrence in the window.'
                        $report.Add([pscustomobject]$base); continue
                    }
                    $processedRecurringSeries[$seriesLookupKey] = $true
                    if ($appointment.ModifiedOccurrences -and $appointment.ModifiedOccurrences.Count -gt 0) {
                        $base.HasModifiedOccurrences = $true
                    }
                }
                else {
                    $base.Status = 'SkippedUnsupportedCalendarItemType'
                    $base.Note = "Unsupported EWS appointment type: $appointmentType"
                    $report.Add([pscustomobject]$base); continue
                }

                if (-not $appointment.IsMeeting -or $appointment.IsCancelled) {
                    $base.Status = 'SkippedSeriesMasterNotActiveMeeting'
                    $base.Note = 'The bound item is not an active organizer meeting.'
                    $report.Add([pscustomobject]$base); continue
                }
                if ([string]$appointment.MyResponseType -ne 'Organizer') {
                    $base.Status = 'SkippedSeriesMasterNotOrganizer'
                    $base.Note = 'The mailbox is not the organizer of the bound recurring master.'
                    $report.Add([pscustomobject]$base); continue
                }

                $bodyText = [string]$appointment.Body.Text
                $locationText = [string]$appointment.Location
                if (($bodyText + "`n" + $locationText) -match '(?i)teams\.microsoft\.com/l/meetup-join') {
                    $base.Status = 'SkippedAlreadyTeams'; $base.Note = 'A Teams join URL is already present.'
                    $report.Add([pscustomobject]$base); continue
                }

                $webexCandidates = New-Object System.Collections.Generic.List[object]
                foreach ($field in @(
                    [pscustomobject]@{ Source = 'Body'; Text = $bodyText },
                    [pscustomobject]@{ Source = 'Location'; Text = $locationText }
                )) {
                    foreach ($urlMatch in [regex]::Matches([string]$field.Text, $webexPattern)) {
                        $candidateUrl = [string]$urlMatch.Value.TrimEnd('.', ',', ';', ':', '!', '?', ')', ']')
                        if (-not [string]::IsNullOrWhiteSpace($candidateUrl)) {
                            $webexCandidates.Add([pscustomobject]@{ Url = $candidateUrl; Source = $field.Source })
                        }
                    }
                }

                $diagnosticEntries = New-Object System.Collections.Generic.List[string]
                $candidateNumber = 0
                foreach ($webexCandidate in $webexCandidates) {
                    $candidateNumber++
                    $parsedMtid = Get-WebexMtidValue -Url $webexCandidate.Url
                    if ([string]::IsNullOrWhiteSpace($parsedMtid)) { $parsedMtid = '<none>' }
                    $candidateKind = if ($webexCandidate.Url -match '(?i)/j\.php(?:[?#]|$)') { 'JoinUrl' } else { 'Other' }
                    $diagnosticEntries.Add(('#{0} Kind={1}; Source={2}; MTID={3}; URL={4}' -f $candidateNumber, $candidateKind, $webexCandidate.Source, $parsedMtid, $webexCandidate.Url))
                }
                $base.WebexUrlDiagnostics = $diagnosticEntries.ToArray() -join ' || '

                # Deduplicate links repeated in the body and Location while retaining their source fields.
                # Treat full URLs as case-sensitive so distinct meeting IDs are never collapsed.
                $webexUrlsByValue = [System.Collections.Generic.Dictionary[string, object]]::new(
                    [System.StringComparer]::Ordinal
                )
                foreach ($candidateUrl in $webexCandidates) {
                    if ($webexUrlsByValue.ContainsKey($candidateUrl.Url)) {
                        if (-not $webexUrlsByValue[$candidateUrl.Url].Sources.Contains($candidateUrl.Source)) {
                            $webexUrlsByValue[$candidateUrl.Url].Sources.Add($candidateUrl.Source)
                        }
                    }
                    else {
                        $sources = New-Object System.Collections.Generic.List[string]
                        $sources.Add($candidateUrl.Source)
                        $webexUrlsByValue[$candidateUrl.Url] = [pscustomobject]@{
                            Url = $candidateUrl.Url
                            Sources = $sources
                            MtidValue = Get-WebexMtidValue -Url $candidateUrl.Url
                            IsJoinUrl = $candidateUrl.Url -match '(?i)/j\.php(?:[?#]|$)'
                        }
                    }
                }
                $uniqueWebexUrls = @($webexUrlsByValue.Values)
                if ($uniqueWebexUrls.Count -eq 0) {
                    $base.Status = 'SkippedNoWebex'; $base.Note = 'No Webex URL found in body or location.'
                    $report.Add([pscustomobject]$base); continue
                }

                # The scheduled j.php link is the attendee join URL. Cisco call-in links can
                # also have an MTID, so prefer j.php candidates when any are present.
                $joinMtidUrls = @($uniqueWebexUrls | Where-Object { $_.IsJoinUrl -and -not [string]::IsNullOrWhiteSpace([string]$_.MtidValue) })
                if ($joinMtidUrls.Count -gt 0) {
                    $mtidSelectionPool = $joinMtidUrls
                    $mtidSelectionDescription = 'j.php join URLs'
                }
                else {
                    $mtidSelectionPool = @($uniqueWebexUrls | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.MtidValue) })
                    $mtidSelectionDescription = 'Webex URLs'
                }

                $distinctMtidValues = New-Object System.Collections.Generic.List[string]
                foreach ($webexUrl in $mtidSelectionPool) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$webexUrl.MtidValue) -and -not $distinctMtidValues.Contains([string]$webexUrl.MtidValue)) {
                        $distinctMtidValues.Add([string]$webexUrl.MtidValue)
                    }
                }
                if ($distinctMtidValues.Count -gt 1) {
                    $base.Status = 'SkippedMultipleMtidWebex'
                    $base.Note = "Found $($distinctMtidValues.Count) different MTID values in $mtidSelectionDescription; manual review required. See WebexUrlDiagnostics for each extracted URL and MTID."
                    $report.Add([pscustomobject]$base); continue
                }

                if ($distinctMtidValues.Count -eq 1) {
                    $selectedMtidValue = $distinctMtidValues[0]
                    $selectedWebexUrls = @($mtidSelectionPool | Where-Object { $_.MtidValue -ceq $selectedMtidValue })
                }
                elseif ($uniqueWebexUrls.Count -eq 1) {
                    $selectedWebexUrls = @($uniqueWebexUrls[0])
                }
                else {
                    $base.Status = 'SkippedMultipleWebex'
                    $base.Note = "Found $($uniqueWebexUrls.Count) distinct Webex URLs and none has an MTID; manual review required. See WebexUrlDiagnostics for each extracted URL."
                    $report.Add([pscustomobject]$base); continue
                }

                $selectedWebexUrl = $selectedWebexUrls[0]
                $oldUrl = [string]$selectedWebexUrl.Url
                $selectedWebexUrlValues = @($selectedWebexUrls | ForEach-Object { [string]$_.Url } | Sort-Object Length -Descending)
                $selectedSources = New-Object System.Collections.Generic.List[string]
                foreach ($selectedUrl in $selectedWebexUrls) {
                    foreach ($source in $selectedUrl.Sources) {
                        if (-not $selectedSources.Contains($source)) { $selectedSources.Add($source) }
                    }
                }
                $webexUrlSource = $selectedSources.ToArray() -join '+'
                $selectionNote = ''
                if ($joinMtidUrls.Count -gt 0) {
                    $selectionNote += ' Selected the scheduled j.php join-link group.'
                }
                if ($selectedWebexUrls.Count -gt 1) {
                    $selectionNote += " Found $($selectedWebexUrls.Count) URL variants sharing the same MTID; all selected variants will be replaced."
                }
                $otherUrlCount = $uniqueWebexUrls.Count - $selectedWebexUrls.Count
                if ($otherUrlCount -gt 0) {
                    $selectionNote += " Left $otherUrlCount other Webex URL(s) outside the selected group untouched."
                }
                $base.OriginalWebexUrl = $oldUrl
                $base.WebexUrlSource = $webexUrlSource
                if ([string]::IsNullOrWhiteSpace([string]$appointment.ICalUid)) {
                    $base.Status = 'SkippedNoICalUid'; $base.Note = 'No iCalendar UID; stable meeting identity cannot be generated.'
                    $report.Add([pscustomobject]$base); continue
                }

                if ($Mode -eq 'Preview') {
                    $base.Status = if ($base.MeetingScope -eq 'RecurringSeries') { 'WouldApplyRecurringSeries' } else { 'WouldApply' }
                    $base.Note = 'Preview only; Graph was not called and Exchange was not changed.'
                    $base.Note += $selectionNote
                    if ($base.HasModifiedOccurrences) {
                        $base.Note += ' The series has modified occurrences; this script does not rewrite exception bodies independently. Verify and manually update each exception invite after applying.'
                    }
                    $report.Add([pscustomobject]$base); continue
                }

                try {
                    $externalId = New-TeamsExternalId -MailboxSmtp $mailbox -ICalUid ([string]$appointment.ICalUid) -Start $identityStart
                    $teamsUrl = Get-TeamsJoinUrl -EntraObjectId $entraObjectId -ExternalId $externalId `
                        -Subject $appointment.Subject -Start $meetingStart -End $meetingEnd
                    $base.TeamsJoinUrl = $teamsUrl

                    $bodyType = $appointment.Body.BodyType
                    $newBodyText = Get-NewBodyText -OriginalBody $bodyText -OriginalUrls $selectedWebexUrlValues `
                        -TeamsUrl $teamsUrl -BodyType $bodyType
                    $newLocationText = $locationText
                    foreach ($selectedUrl in $selectedWebexUrlValues) {
                        $newLocationText = $newLocationText.Replace($selectedUrl, $teamsUrl)
                    }
                    $base.Status = 'ReadyToApply'
                    $base.Note = if ($base.MeetingScope -eq 'RecurringSeries') {
                        'Graph meeting created or reused. The recurring master is updated so the series uses one shared Teams join URL.'
                    } else {
                        'Graph meeting created or reused. The Teams join URL is added to the body and the selected Webex URL is replaced.'
                    }
                    $base.Note += $selectionNote
                    if ($base.HasModifiedOccurrences) {
                        $base.Note += ' The series has modified occurrences; this script does not rewrite exception bodies independently. Verify and manually update each exception invite after applying.'
                    }

                    $backup = [pscustomobject]@{
                        MailboxSmtp = $mailbox
                        ItemId = [string]$appointment.Id.UniqueId
                        ICalUid = [string]$appointment.ICalUid
                        Subject = [string]$appointment.Subject
                        StartUtc = ConvertTo-UtcIsoString -Value $meetingStart
                        MeetingScope = $base.MeetingScope
                        OriginalBodyType = [string]$bodyType
                        OriginalBodyText = $bodyText
                        NewBodyText = $newBodyText
                        OriginalLocation = $locationText
                        NewLocation = $newLocationText
                        OriginalWebexUrl = $oldUrl
                        TeamsJoinUrl = $teamsUrl
                    }
                    if (-not (Test-Path -LiteralPath $BackupPath)) {
                        New-Item -ItemType File -Path $BackupPath -Force | Out-Null
                    }
                    Write-JsonLine -Object $backup -Path $BackupPath

                    $appointment.Body = New-BodyValue -BodyType $bodyType -Text $newBodyText
                    if ($newLocationText -cne $locationText) {
                        $appointment.Location = $newLocationText
                    }
                    $appointment.Update(
                        [Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite,
                        [Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode]::SendToAllAndSaveCopy
                    )
                    $base.Status = 'Applied'
                    if ($base.MeetingScope -eq 'RecurringSeries') {
                        $base.Note = 'Updated recurring master; EWS sent a meeting update for the series. This script does not rewrite exception bodies independently; verify and manually update each modified occurrence invite.'
                    } else {
                        $base.Note = 'Updated in Exchange; meeting update sent to all attendees. Review link label and invite content.'
                    }
                    $base.Note += $selectionNote
                }
                catch {
                    $base.Status = 'ApplyError'
                    $base.Note = $_.Exception.Message
                }
                $report.Add([pscustomobject]$base)
            }
        }
        catch {
            $report.Add([pscustomobject]@{
                MailboxSmtp = $mailbox; Subject = ''; StartUtc = ''; OriginalWebexUrl = ''
                WebexUrlSource = ''; WebexUrlDiagnostics = ''; TeamsJoinUrl = ''
                Status = 'MailboxError'; Note = $_.Exception.Message
            })
        }
    }
}

if ($report.Count -gt 0) {
    $report | Export-Csv -LiteralPath $ReportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Report written to: $ReportPath"
}
if ($Mode -eq 'Apply' -and (Test-Path -LiteralPath $BackupPath)) {
    Write-Host "Rollback backup written to: $BackupPath"
}
if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
    if (Get-MgContext) {
        Disconnect-MgGraph | Out-Null
    }
}
