[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string] $EwsUrl,
    [ValidateSet('System', 'Direct', 'SystemDefaultCredentials')][string] $EwsProxyMode = 'Direct',
    [Parameter(Mandatory = $true)][string] $EwsManagedApiDll,
    [Parameter(Mandatory = $true)][string] $TargetsCsv,
    [Parameter(Mandatory = $true)][datetimeoffset] $WindowStartUtc,
    [Parameter(Mandatory = $true)][datetimeoffset] $WindowEndUtc,
    [System.Management.Automation.PSCredential] $EwsCredential,
    [string] $EwsCredentialPath = (Join-Path $env:LOCALAPPDATA 'WebexTeamsPOC\ews-credential.xml'),
    [string] $OutputDirectory = '.\output',
    [switch] $Force
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'WebexTeams.Common.ps1')

if (-not (Test-Path -LiteralPath $TargetsCsv -PathType Leaf)) { throw "Targets CSV not found: $TargetsCsv" }
if ($WindowEndUtc -le $WindowStartUtc) { throw 'WindowEndUtc must be later than WindowStartUtc.' }
$endpoint = [uri]$EwsUrl
if ($endpoint.Scheme -ne 'https') { throw 'EWS URL must use HTTPS.' }
Initialize-WebexEwsApi -ManagedApiDll $EwsManagedApiDll
Initialize-OutputDirectory -Path $OutputDirectory

$targets = @(Import-Csv -LiteralPath $TargetsCsv)
if ($targets.Count -eq 0) { throw 'Targets CSV is empty.' }
foreach ($target in $targets) {
    if ([string]::IsNullOrWhiteSpace([string]$target.MailboxSmtp) -or [string]::IsNullOrWhiteSpace([string]$target.EntraObjectId)) {
        throw 'Each target row must have MailboxSmtp and EntraObjectId columns.'
    }
}

$preflightMailboxes = @{}
foreach ($target in $targets) {
    $mailbox = ([string]$target.MailboxSmtp).Trim()
    $mailboxKey = $mailbox.ToLowerInvariant()
    if ($preflightMailboxes.ContainsKey($mailboxKey)) { continue }
    $preflightMailboxes[$mailboxKey] = $true
    $outputPath = Join-Path $OutputDirectory ("{0}-webex-export.csv" -f (Get-SafeUserFileStem -MailboxSmtp $mailbox))
    Assert-OutputFileAvailable -Path $outputPath -Force:$Force
}
$EwsCredential = Resolve-EwsCredential -Credential $EwsCredential -CredentialPath $EwsCredentialPath

$findProps = New-EwsFindPropertySet
$detailProps = New-EwsDetailPropertySet
$start = $WindowStartUtc.UtcDateTime
$end = $WindowEndUtc.UtcDateTime
$seenMailboxes = @{}

foreach ($target in $targets) {
    $mailbox = ([string]$target.MailboxSmtp).Trim()
    $entraId = ([string]$target.EntraObjectId).Trim()
    $mailboxKey = $mailbox.ToLowerInvariant()
    if ($seenMailboxes.ContainsKey($mailboxKey)) { continue }
    $seenMailboxes[$mailboxKey] = $true

    $fileStem = Get-SafeUserFileStem -MailboxSmtp $mailbox
    $outputPath = Join-Path $OutputDirectory ("{0}-webex-export.csv" -f $fileStem)
    Assert-OutputFileAvailable -Path $outputPath -Force:$Force
    $rows = New-Object System.Collections.Generic.List[object]

    try {
        $service = New-EwsService -MailboxSmtp $mailbox -Endpoint $endpoint -Credential $EwsCredential -ProxyMode $EwsProxyMode
        $calendarItems = @(Get-EwsCalendarItemsInRange -Service $service -StartUtc $start -EndUtc $end -PropertySet $findProps)
        $uniqueItems = New-Object System.Collections.Generic.List[object]
        $itemKeys = @{}
        foreach ($item in $calendarItems) {
            $key = '{0}|{1}' -f [string]$item.Id.UniqueId, (ConvertTo-UtcIsoString -Value $item.Start)
            if (-not $itemKeys.ContainsKey($key)) { $itemKeys[$key] = $true; $uniqueItems.Add($item) }
        }
        $processedSeries = @{}

        foreach ($candidate in $uniqueItems) {
            $candidateStart = ConvertTo-UtcIsoString -Value $candidate.Start
            $row = [ordered]@{
                MailboxSmtp = $mailbox
                EntraObjectId = $entraId
                Subject = [string]$candidate.Subject
                StartUtc = $candidateStart
                MeetingScope = 'Single'
                HasModifiedOccurrences = $false
                AppointmentType = [string]$candidate.AppointmentType
                EwsItemId = ''
                EwsChangeKey = ''
                ICalUid = [string]$candidate.ICalUid
                IdentityStartUtc = $candidateStart
                GraphStartUtc = $candidateStart
                GraphEndUtc = ConvertTo-UtcIsoString -Value $candidate.End
                OriginalWebexUrl = ''
                WebexUrlSource = ''
                WebexUrlDiagnostics = ''
                SelectedWebexUrlsJson = '[]'
                TeamsJoinUrl = ''
                Status = ''
                Note = ''
            }

            if (-not $candidate.IsMeeting) { $row.Status = 'SkippedNotMeeting'; $row.Note = 'Calendar item is not a meeting.'; $rows.Add([pscustomobject]$row); continue }
            if ($candidate.IsCancelled) { $row.Status = 'SkippedCancelled'; $row.Note = 'Meeting is cancelled.'; $rows.Add([pscustomobject]$row); continue }
            if ([string]$candidate.MyResponseType -ne 'Organizer') { $row.Status = 'SkippedNotOrganizer'; $row.Note = 'Not organizer.'; $rows.Add([pscustomobject]$row); continue }

            $type = [string]$candidate.AppointmentType
            $seriesKey = ''
            if ($type -in @('Occurrence', 'Exception', 'RecurringMaster') -and $candidate.ICalUid) {
                $seriesKey = '{0}|{1}' -f $mailboxKey, [string]$candidate.ICalUid
                if ($processedSeries.ContainsKey($seriesKey)) {
                    $row.MeetingScope = 'RecurringSeries'
                    $row.Status = 'SkippedDuplicateSeriesOccurrence'
                    $row.Note = 'This recurring series was already exported from another occurrence in the window.'
                    $rows.Add([pscustomobject]$row); continue
                }
            }

            try {
                if ($type -eq 'Single') {
                    $appointment = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service, $candidate.Id)
                }
                elseif ($type -in @('Occurrence', 'Exception')) {
                    $appointment = [Microsoft.Exchange.WebServices.Data.Appointment]::BindToRecurringMaster($service, $candidate.Id)
                    $row.MeetingScope = 'RecurringSeries'
                }
                elseif ($type -eq 'RecurringMaster') {
                    $appointment = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service, $candidate.Id)
                    $row.MeetingScope = 'RecurringSeries'
                }
                else {
                    $row.Status = 'SkippedUnsupportedCalendarItemType'
                    $row.Note = "Unsupported EWS appointment type: $type"
                    $rows.Add([pscustomobject]$row); continue
                }
                $appointment.Load($detailProps)
            }
            catch {
                $row.MeetingScope = if ($type -eq 'Single') { 'Single' } else { 'RecurringSeries' }
                $row.Status = 'EwsBindError'
                $row.Note = $_.Exception.Message
                $rows.Add([pscustomobject]$row); continue
            }

            if ($row.MeetingScope -eq 'RecurringSeries') {
                if (-not $seriesKey) { $seriesKey = '{0}|{1}' -f $mailboxKey, [string]$appointment.ICalUid }
                if ($processedSeries.ContainsKey($seriesKey)) {
                    $row.Status = 'SkippedDuplicateSeriesOccurrence'
                    $row.Note = 'This recurring series was already exported from another occurrence in the window.'
                    $rows.Add([pscustomobject]$row); continue
                }
                $processedSeries[$seriesKey] = $true
                $row.HasModifiedOccurrences = ($appointment.ModifiedOccurrences -and $appointment.ModifiedOccurrences.Count -gt 0)
                $row.IdentityStartUtc = ConvertTo-UtcIsoString -Value $appointment.Start
            }

            if (-not $appointment.IsMeeting -or $appointment.IsCancelled) {
                $row.Status = 'SkippedSeriesMasterNotActiveMeeting'; $row.Note = 'The bound item is not an active organizer meeting.'; $rows.Add([pscustomobject]$row); continue
            }
            if ([string]$appointment.MyResponseType -ne 'Organizer') {
                $row.Status = 'SkippedSeriesMasterNotOrganizer'; $row.Note = 'The mailbox is not the organizer of the bound item.'; $rows.Add([pscustomobject]$row); continue
            }

            $row.EwsItemId = [string]$appointment.Id.UniqueId
            $row.EwsChangeKey = [string]$appointment.Id.ChangeKey
            $row.ICalUid = [string]$appointment.ICalUid
            $bodyText = [string]$appointment.Body.Text
            $locationText = [string]$appointment.Location
            if (($bodyText + "`n" + $locationText) -match '(?i)teams\.microsoft\.com/l/meetup-join') {
                $row.Status = 'SkippedAlreadyTeams'; $row.Note = 'A Teams join URL is already present.'; $rows.Add([pscustomobject]$row); continue
            }

            $selection = Select-WebexJoinUrls -BodyText $bodyText -LocationText $locationText
            $row.Status = $selection.Status
            $row.Note = $selection.Note
            $row.OriginalWebexUrl = $selection.OriginalWebexUrl
            $row.WebexUrlSource = $selection.WebexUrlSource
            $row.WebexUrlDiagnostics = $selection.WebexUrlDiagnostics
            $row.SelectedWebexUrlsJson = $selection.SelectedWebexUrlsJson
            if ($row.Status -eq 'ReadyForTeamsLink' -and $row.HasModifiedOccurrences) {
                $row.Note += ' This recurring series has modified occurrences; the update stage changes the master only. Inspect exception invites after applying.'
            }
            $rows.Add([pscustomobject]$row)
        }

        if ($rows.Count -eq 0) {
            $rows.Add([pscustomobject]@{
                MailboxSmtp = $mailbox; EntraObjectId = $entraId; Subject = ''; StartUtc = ''
                MeetingScope = ''; HasModifiedOccurrences = $false; AppointmentType = ''; EwsItemId = ''
                EwsChangeKey = ''; ICalUid = ''; IdentityStartUtc = ''; GraphStartUtc = ''; GraphEndUtc = ''
                OriginalWebexUrl = ''; WebexUrlSource = ''; WebexUrlDiagnostics = ''; SelectedWebexUrlsJson = '[]'
                TeamsJoinUrl = ''; Status = 'NoCalendarItemsInWindow'; Note = 'No calendar items were found in the requested window.'
            })
        }
    }
    catch {
        $rows.Clear()
        $rows.Add([pscustomobject]@{
            MailboxSmtp = $mailbox; EntraObjectId = $entraId; Subject = ''; StartUtc = ''
            MeetingScope = ''; HasModifiedOccurrences = $false; AppointmentType = ''; EwsItemId = ''
            EwsChangeKey = ''; ICalUid = ''; IdentityStartUtc = ''; GraphStartUtc = ''; GraphEndUtc = ''
            OriginalWebexUrl = ''; WebexUrlSource = ''; WebexUrlDiagnostics = ''; SelectedWebexUrlsJson = '[]'
            TeamsJoinUrl = ''; Status = 'MailboxError'; Note = $_.Exception.Message
        })
    }

    $rows | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding UTF8
    $eligibleCount = @($rows | Where-Object { $_.Status -eq 'ReadyForTeamsLink' }).Count
    Write-Host ("{0}: {1} rows; {2} eligible for Teams link creation. Report: {3}" -f $mailbox, $rows.Count, $eligibleCount, $outputPath)
}
