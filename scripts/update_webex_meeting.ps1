[CmdletBinding()]
param(
    [ValidateSet('Apply', 'Rollback')][string] $Mode = 'Apply',
    [Parameter(Mandatory = $true)][string] $EwsUrl,
    [ValidateSet('System', 'Direct', 'SystemDefaultCredentials')][string] $EwsProxyMode = 'Direct',
    [Parameter(Mandatory = $true)][string] $EwsManagedApiDll,
    [System.Management.Automation.PSCredential] $EwsCredential,
    [string] $EwsCredentialPath = (Join-Path $env:LOCALAPPDATA 'WebexTeamsPOC\ews-credential.xml'),
    [string] $InputDirectory = '.\prepared',
    [string] $FilePattern = '*.csv',
    [string] $OutputDirectory = '.\applied',
    [string] $BackupPath = ('.\backups\webex-teams-update-{0}.jsonl' -f (Get-Date -Format 'yyyyMMdd-HHmmss')),
    [switch] $Force
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'WebexTeams.Common.ps1')

$endpoint = [uri]$EwsUrl
if ($endpoint.Scheme -ne 'https') { throw 'EWS URL must use HTTPS.' }
Initialize-WebexEwsApi -ManagedApiDll $EwsManagedApiDll
$EwsCredential = Resolve-EwsCredential -Credential $EwsCredential -CredentialPath $EwsCredentialPath
$detailProps = New-EwsDetailPropertySet
Initialize-OutputDirectory -Path $OutputDirectory

if ($Mode -eq 'Rollback') {
    if (-not (Test-Path -LiteralPath $BackupPath -PathType Leaf)) { throw "Rollback backup file not found: $BackupPath" }
    $rollbackRows = New-Object System.Collections.Generic.List[object]
    $records = @(Get-Content -LiteralPath $BackupPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_ | ConvertFrom-Json })
    foreach ($record in $records) {
        try {
            $service = New-EwsService -MailboxSmtp ([string]$record.MailboxSmtp) -Endpoint $endpoint -Credential $EwsCredential -ProxyMode $EwsProxyMode
            $itemId = [Microsoft.Exchange.WebServices.Data.ItemId]::new([string]$record.EwsItemId)
            $appointment = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service, $itemId)
            $appointment.Load($detailProps)
            $locationMatches = $appointment.Location -ceq [string]$record.NewLocation
            $bodyMatches = $appointment.Body.Text -ceq [string]$record.NewBodyText
            if (-not $bodyMatches -or -not $locationMatches) {
                $rollbackRows.Add([pscustomobject]@{ MailboxSmtp = $record.MailboxSmtp; Subject = $record.Subject; StartUtc = $record.StartUtc; Status = 'RollbackSkippedChanged'; Note = 'Current body or location no longer matches the saved post-apply value.' })
                continue
            }
            $oldType = [Microsoft.Exchange.WebServices.Data.BodyType][Enum]::Parse([Microsoft.Exchange.WebServices.Data.BodyType], [string]$record.OriginalBodyType, $true)
            $appointment.Body = New-BodyValue -BodyType $oldType -Text ([string]$record.OriginalBodyText)
            $appointment.Location = [string]$record.OriginalLocation
            $appointment.Update(
                [Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite,
                [Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode]::SendToAllAndSaveCopy
            )
            $rollbackRows.Add([pscustomobject]@{ MailboxSmtp = $record.MailboxSmtp; Subject = $record.Subject; StartUtc = $record.StartUtc; Status = 'RolledBack'; Note = 'Original body and location restored; attendee update sent.' })
        }
        catch {
            $rollbackRows.Add([pscustomobject]@{ MailboxSmtp = $record.MailboxSmtp; Subject = $record.Subject; StartUtc = $record.StartUtc; Status = 'RollbackError'; Note = $_.Exception.Message })
        }
    }
    $rollbackPath = Join-Path $OutputDirectory 'rollback-report.csv'
    Assert-OutputFileAvailable -Path $rollbackPath -Force:$Force
    $rollbackRows | Export-Csv -LiteralPath $rollbackPath -NoTypeInformation -Encoding UTF8
    Write-Host "Rollback report written to: $rollbackPath"
    return
}

if (-not (Test-Path -LiteralPath $InputDirectory -PathType Container)) { throw "Prepared CSV directory not found: $InputDirectory" }
$files = @(Get-ChildItem -LiteralPath $InputDirectory -Filter $FilePattern -File | Sort-Object Name)
if ($files.Count -eq 0) { throw "No prepared CSV files matching '$FilePattern' were found in '$InputDirectory'." }
if (Test-Path -LiteralPath $BackupPath) { throw "Backup path already exists. Choose a new path: $BackupPath" }
$plannedOutputs = @{}
foreach ($file in $files) {
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    $outputPath = Join-Path $OutputDirectory ($stem + '-applied.csv')
    if ($plannedOutputs.ContainsKey($outputPath)) { throw "More than one prepared CSV maps to the same result report: $outputPath" }
    $plannedOutputs[$outputPath] = $true
    Assert-OutputFileAvailable -Path $outputPath -Force:$Force
}
$backupDirectory = Split-Path -Path ([System.IO.Path]::GetFullPath($BackupPath)) -Parent
if ($backupDirectory) { Initialize-OutputDirectory -Path $backupDirectory }

foreach ($file in $files) {
    $rows = @(Import-Csv -LiteralPath $file.FullName)
    $outputRows = New-Object System.Collections.Generic.List[object]
    foreach ($row in $rows) {
        if ([string]$row.Status -ne 'ReadyForUpdate') { $outputRows.Add($row); continue }
        if ([string]$row.ApprovedForUpdate -ne 'Yes') {
            $row.Status = 'AwaitingApproval'
            $row.Note = 'Not applied. Review this row in the prepared CSV and set ApprovedForUpdate to Yes to include it.'
            $outputRows.Add($row); continue
        }
        try {
            foreach ($field in @('MailboxSmtp', 'EwsItemId', 'EwsChangeKey', 'SelectedWebexUrlsJson', 'TeamsJoinUrl')) {
                if ([string]::IsNullOrWhiteSpace([string]$row.$field)) { throw "Prepared row is missing required field '$field'." }
            }
            $selectedUrls = @(([string]$row.SelectedWebexUrlsJson | ConvertFrom-Json) | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($selectedUrls.Count -eq 0) { throw 'Prepared row has no selected Webex URLs.' }
            $service = New-EwsService -MailboxSmtp ([string]$row.MailboxSmtp) -Endpoint $endpoint -Credential $EwsCredential -ProxyMode $EwsProxyMode
            $itemId = [Microsoft.Exchange.WebServices.Data.ItemId]::new([string]$row.EwsItemId)
            $appointment = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind($service, $itemId)
            $appointment.Load($detailProps)
            if ([string]$appointment.Id.ChangeKey -cne [string]$row.EwsChangeKey) {
                $row.Status = 'SkippedChangedSinceFetch'
                $row.Note = 'The Exchange item changed after the Webex export. Fetch it again and review a new prepared report.'
                $outputRows.Add($row); continue
            }
            if ([string]$appointment.MyResponseType -ne 'Organizer') {
                $row.Status = 'SkippedNotOrganizer'
                $row.Note = 'The mailbox is no longer the organizer of this item.'
                $outputRows.Add($row); continue
            }

            $bodyType = $appointment.Body.BodyType
            $bodyText = [string]$appointment.Body.Text
            $locationText = [string]$appointment.Location
            $missingUrls = New-Object System.Collections.Generic.List[string]
            foreach ($url in $selectedUrls) {
                if (-not $bodyText.Contains($url) -and -not $locationText.Contains($url)) {
                    $missingUrls.Add($url)
                }
            }
            if ($missingUrls.Count -gt 0) {
                $row.Status = 'SkippedWebexUrlChanged'
                $row.Note = 'Selected Webex URL no longer appears in the current Exchange item. Fetch and review this meeting again.'
                $outputRows.Add($row); continue
            }
            if (($bodyText + "`n" + $locationText) -match '(?i)teams\.microsoft\.com/l/meetup-join') {
                $row.Status = 'SkippedAlreadyTeams'
                $row.Note = 'A Teams join URL is already present in the Exchange item.'
                $outputRows.Add($row); continue
            }

            $newBodyText = Get-NewBodyText -OriginalBody $bodyText -OriginalUrls $selectedUrls -TeamsUrl ([string]$row.TeamsJoinUrl) -BodyType $bodyType
            $newLocationText = $locationText
            foreach ($url in $selectedUrls) { $newLocationText = $newLocationText.Replace($url, [string]$row.TeamsJoinUrl) }
            $backup = [pscustomobject]@{
                MailboxSmtp = [string]$row.MailboxSmtp
                EwsItemId = [string]$appointment.Id.UniqueId
                ICalUid = [string]$row.ICalUid
                Subject = [string]$row.Subject
                StartUtc = [string]$row.StartUtc
                MeetingScope = [string]$row.MeetingScope
                OriginalBodyType = [string]$bodyType
                OriginalBodyText = $bodyText
                NewBodyText = $newBodyText
                OriginalLocation = $locationText
                NewLocation = $newLocationText
                OriginalWebexUrl = [string]$row.OriginalWebexUrl
                TeamsJoinUrl = [string]$row.TeamsJoinUrl
            }
            Write-JsonLine -Object $backup -Path $BackupPath
            $appointment.Body = New-BodyValue -BodyType $bodyType -Text $newBodyText
            if ($newLocationText -cne $locationText) { $appointment.Location = $newLocationText }
            $appointment.Update(
                [Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite,
                [Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode]::SendToAllAndSaveCopy
            )
            $row.Status = 'Applied'
            if ([string]$row.MeetingScope -eq 'RecurringSeries') {
                $row.Note = 'Updated recurring master; EWS sent a meeting update for the series. Inspect modified occurrence invites separately.'
            }
            else { $row.Note = 'Updated in Exchange; meeting update sent to all attendees.' }
        }
        catch {
            $row.Status = 'UpdateError'
            $row.Note = $_.Exception.Message
        }
        $outputRows.Add($row)
    }

    $stem = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    $outputPath = Join-Path $OutputDirectory ($stem + '-applied.csv')
    $outputRows | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding UTF8
    $appliedCount = @($outputRows | Where-Object { $_.Status -eq 'Applied' }).Count
    Write-Host ("{0}: {1} meetings updated. Result report: {2}" -f $file.Name, $appliedCount, $outputPath)
}
if (Test-Path -LiteralPath $BackupPath) { Write-Host "Rollback backup written to: $BackupPath" }
