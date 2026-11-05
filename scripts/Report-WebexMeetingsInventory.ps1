[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string] $EwsUrl,
    [ValidateSet('System', 'Direct', 'SystemDefaultCredentials')][string] $EwsProxyMode = 'Direct',
    [Parameter(Mandatory = $true)][string] $EwsManagedApiDll,
    [Parameter(Mandatory = $true)][string] $TargetsCsv,
    [datetimeoffset] $WindowStartUtc,
    [datetimeoffset] $WindowEndUtc,
    [System.Management.Automation.PSCredential] $EwsCredential,
    [string] $EwsCredentialPath = (Join-Path $env:LOCALAPPDATA 'WebexTeamsPOC\ews-credential.xml'),
    [string] $ReportPath = ('.\webex-meetings-detail-{0}.csv' -f (Get-Date -Format 'yyyyMMdd-HHmmss')),
    [string] $SummaryPath = ('.\webex-meetings-summary-{0}.csv' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
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
    if (-not $newCredential) { throw 'No EWS credential was supplied.' }
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
    return ,$props
}

function Get-CalendarAppointmentsInRange {
    param(
        [Parameter(Mandatory = $true)] $Service,
        [Parameter(Mandatory = $true)][datetime] $StartUtc,
        [Parameter(Mandatory = $true)][datetime] $EndUtc,
        [Parameter(Mandatory = $true)] $PropertySet,
        [int] $Depth = 0
    )

    $calendarView = [Microsoft.Exchange.WebServices.Data.CalendarView]::new($StartUtc, $EndUtc, 1000)
    $calendarView.PropertySet = $PropertySet
    $page = $Service.FindAppointments(
        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,
        $calendarView
    )
    if ($page.MoreAvailable) {
        $duration = $EndUtc - $StartUtc
        if ($duration.TotalMinutes -le 60 -or $Depth -ge 24) {
            throw 'Calendar query still exceeds 1,000 results after splitting the window to one hour or less; mailbox report is incomplete.'
        }
        $midpoint = $StartUtc.AddTicks([long]($duration.Ticks / 2))
        $firstHalf = @(Get-CalendarAppointmentsInRange -Service $Service -StartUtc $StartUtc `
            -EndUtc $midpoint -PropertySet $PropertySet -Depth ($Depth + 1))
        $secondHalf = @(Get-CalendarAppointmentsInRange -Service $Service -StartUtc $midpoint `
            -EndUtc $EndUtc -PropertySet $PropertySet -Depth ($Depth + 1))
        return @($firstHalf + $secondHalf)
    }
    return @($page.Items)
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
        try { $mtidValue = [System.Uri]::UnescapeDataString($mtidValue) } catch { }
        return [string]$mtidValue.Trim().TrimEnd('.', ',', ';', ':', '!', '?', ')', ']')
    }
    return ''
}

if (-not (Test-Path -LiteralPath $EwsManagedApiDll -PathType Leaf)) {
    throw "EWS Managed API assembly not found: $EwsManagedApiDll"
}
if (-not (Test-Path -LiteralPath $TargetsCsv -PathType Leaf)) {
    throw "Mailbox inventory CSV not found: $TargetsCsv"
}
if ($WindowStartUtc -eq [datetimeoffset]::MinValue -or $WindowEndUtc -eq [datetimeoffset]::MinValue -or $WindowEndUtc -le $WindowStartUtc) {
    throw 'Provide a valid -WindowStartUtc and -WindowEndUtc. Use UTC timestamps, for example 2026-11-05T00:00:00Z.'
}

$endpointUri = [uri]$EwsUrl
if ($endpointUri.Scheme -ne 'https') { throw 'EWS URL must use HTTPS.' }
Add-Type -Path $EwsManagedApiDll
$EwsCredential = Resolve-EwsCredential -Credential $EwsCredential -CredentialPath $EwsCredentialPath
$targets = @(Import-Csv -LiteralPath $TargetsCsv)
if ($targets.Count -eq 0) { throw 'Mailbox inventory CSV is empty.' }
foreach ($target in $targets) {
    if ([string]::IsNullOrWhiteSpace([string]$target.MailboxSmtp)) {
        throw 'Each mailbox inventory row must contain MailboxSmtp.'
    }
}

$detailProps = New-DetailPropertySet
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

$report = New-Object System.Collections.Generic.List[object]
$startUtc = $WindowStartUtc.UtcDateTime
$endUtc = $WindowEndUtc.UtcDateTime

foreach ($target in $targets) {
    $mailbox = ([string]$target.MailboxSmtp).Trim()
    $mailboxType = if ($target.PSObject.Properties['MailboxType']) { ([string]$target.MailboxType).Trim() } else { '' }
    try {
        $service = New-EwsService -MailboxSmtp $mailbox -Endpoint $endpointUri `
            -Credential $EwsCredential -ProxyMode $EwsProxyMode
        $items = @(Get-CalendarAppointmentsInRange -Service $service -StartUtc $startUtc `
            -EndUtc $endUtc -PropertySet $findProps)
        $processedItems = @{}
        $processedRecurringSeries = @{}

        if ($items.Count -eq 0) {
            $report.Add([pscustomobject]@{
                MailboxSmtp = $mailbox; MailboxType = $mailboxType; Subject = ''; StartUtc = ''
                MeetingScope = ''; HasModifiedOccurrences = $false; OriginalWebexUrl = ''
                WebexUrlSource = ''; WebexUrlDiagnostics = ''; Status = 'NoCalendarItemsInWindow'
                Note = 'No calendar items were returned for this mailbox and date window.'
            })
            continue
        }

        foreach ($candidate in $items) {
            $itemKey = [string]$candidate.Id.UniqueId
            if ($itemKey -and $processedItems.ContainsKey($itemKey)) { continue }
            if ($itemKey) { $processedItems[$itemKey] = $true }

            $subject = [string]$candidate.Subject
            $base = @{
                MailboxSmtp = $mailbox; MailboxType = $mailboxType; Subject = $subject
                StartUtc = ConvertTo-UtcIsoString -Value $candidate.Start
                MeetingScope = 'Single'; HasModifiedOccurrences = $false
                OriginalWebexUrl = ''; WebexUrlSource = ''; WebexUrlDiagnostics = ''
                Status = ''; Note = ''
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
                    $base.Note = 'This recurring series was already represented by another occurrence in the selected window.'
                    $report.Add([pscustomobject]$base); continue
                }
            }
            $appointment = $null
            if ($appointmentType -eq 'Single') {
                $appointment = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service, $candidate.Id)
                $appointment.Load($detailProps)
            }
            elseif ($appointmentType -in @('Occurrence', 'Exception')) {
                try {
                    $appointment = [Microsoft.Exchange.WebServices.Data.Appointment]::BindToRecurringMaster($service, $candidate.Id)
                    $appointment.Load($detailProps)
                }
                catch {
                    $base.MeetingScope = 'RecurringSeries'; $base.Status = 'RecurringMasterBindError'
                    $base.Note = $_.Exception.Message
                    $report.Add([pscustomobject]$base); continue
                }
                $base.MeetingScope = 'RecurringSeries'
                if (-not $seriesLookupKey) {
                    $seriesLookupKey = '{0}|{1}' -f $mailbox.ToLowerInvariant(), [string]$appointment.ICalUid
                }
                if ($processedRecurringSeries.ContainsKey($seriesLookupKey)) {
                    $base.Status = 'SkippedDuplicateSeriesOccurrence'
                    $base.Note = 'This recurring series was already represented by another occurrence in the selected window.'
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
                if (-not $seriesLookupKey) {
                    $seriesLookupKey = '{0}|{1}' -f $mailbox.ToLowerInvariant(), [string]$appointment.ICalUid
                }
                if ($processedRecurringSeries.ContainsKey($seriesLookupKey)) {
                    $base.Status = 'SkippedDuplicateSeriesOccurrence'
                    $base.Note = 'This recurring series was already represented by another occurrence in the selected window.'
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
                    if ($candidateUrl) { $webexCandidates.Add([pscustomobject]@{ Url = $candidateUrl; Source = $field.Source }) }
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

            $webexUrlsByValue = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::Ordinal)
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
                        Url = $candidateUrl.Url; Sources = $sources
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
                $base.Note = "Found $($distinctMtidValues.Count) different MTID values in $mtidSelectionDescription; manual review required. See WebexUrlDiagnostics."
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
                $base.Note = "Found $($uniqueWebexUrls.Count) distinct Webex URLs and none has an MTID; manual review required. See WebexUrlDiagnostics."
                $report.Add([pscustomobject]$base); continue
            }

            $selectedSources = New-Object System.Collections.Generic.List[string]
            foreach ($selectedUrl in $selectedWebexUrls) {
                foreach ($source in $selectedUrl.Sources) {
                    if (-not $selectedSources.Contains($source)) { $selectedSources.Add($source) }
                }
            }
            $base.OriginalWebexUrl = [string]$selectedWebexUrls[0].Url
            $base.WebexUrlSource = $selectedSources.ToArray() -join '+'
            $selectionNote = ''
            if ($joinMtidUrls.Count -gt 0) { $selectionNote += ' Selected the scheduled j.php join-link group.' }
            if ($selectedWebexUrls.Count -gt 1) {
                $selectionNote += " Found $($selectedWebexUrls.Count) URL variants sharing the same MTID; all are in the selected group."
            }
            $otherUrlCount = $uniqueWebexUrls.Count - $selectedWebexUrls.Count
            if ($otherUrlCount -gt 0) { $selectionNote += " Left $otherUrlCount other Webex URL(s) outside the selected group untouched." }

            if ([string]::IsNullOrWhiteSpace([string]$appointment.ICalUid)) {
                $base.Status = 'SkippedNoICalUid'; $base.Note = 'No iCalendar UID; stable meeting identity cannot be generated.'
                $report.Add([pscustomobject]$base); continue
            }
            $base.Status = if ($base.MeetingScope -eq 'RecurringSeries') { 'WouldApplyRecurringSeries' } else { 'WouldApply' }
            $base.Note = 'Inventory only; Graph was not called and Exchange was not changed.' + $selectionNote
            if ($base.HasModifiedOccurrences) {
                $base.Note += ' The recurring series has modified occurrences; review those exception invites manually.'
            }
            $report.Add([pscustomobject]$base)
        }
    }
    catch {
        $report.Add([pscustomobject]@{
            MailboxSmtp = $mailbox; MailboxType = $mailboxType; Subject = ''; StartUtc = ''
            MeetingScope = ''; HasModifiedOccurrences = $false; OriginalWebexUrl = ''
            WebexUrlSource = ''; WebexUrlDiagnostics = ''; Status = 'MailboxError'
            Note = $_.Exception.Message
        })
    }
}

if ($report.Count -gt 0) {
    $report | Export-Csv -LiteralPath $ReportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Detailed report written to: $ReportPath"

    $summaryRows = New-Object System.Collections.Generic.List[object]
    foreach ($group in @($report | Group-Object -Property Status)) {
        foreach ($typeGroup in @($group.Group | Group-Object -Property {
            if ([string]::IsNullOrWhiteSpace([string]$_.MailboxType)) { 'Unspecified' } else { [string]$_.MailboxType }
        })) {
            $mailboxCount = @($typeGroup.Group | ForEach-Object { [string]$_.MailboxSmtp } | Sort-Object -Unique).Count
            $summaryRows.Add([pscustomobject]@{
                MailboxType = [string]$typeGroup.Name; Status = [string]$group.Name
                ItemCount = $typeGroup.Count; MailboxCount = $mailboxCount
            })
        }
        $allMailboxCount = @($group.Group | ForEach-Object { [string]$_.MailboxSmtp } | Sort-Object -Unique).Count
        $summaryRows.Add([pscustomobject]@{
            MailboxType = 'All'; Status = [string]$group.Name
            ItemCount = $group.Count; MailboxCount = $allMailboxCount
        })
    }
    $summaryRows | Sort-Object MailboxType, Status | Export-Csv -LiteralPath $SummaryPath -NoTypeInformation -Encoding UTF8
    Write-Host "Status summary written to: $SummaryPath"
    $actionableCount = @($report | Where-Object { $_.Status -in @('WouldApply', 'WouldApplyRecurringSeries') }).Count
    Write-Host "Inventory found $actionableCount meeting or recurring-series item(s) ready for review."
}
