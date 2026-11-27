[CmdletBinding()]
param(
    [string] $InputDirectory = '.\output',
    [string] $FilePattern = '*-webex-export.csv',
    [string] $OutputDirectory = '.\prepared',
    [Parameter(Mandatory = $true)][string] $GraphTenantId,
    [Parameter(Mandatory = $true)][string] $GraphAppId,
    [ValidateSet('Certificate', 'ClientSecret')][string] $GraphAuthMethod = 'Certificate',
    [string] $GraphCertificateThumbprint,
    [string] $GraphClientSecretPath = (Join-Path $env:LOCALAPPDATA 'WebexTeamsPOC\graph-client-secret.xml'),
    [switch] $Force
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'WebexTeams.Common.ps1')

if (-not (Test-Path -LiteralPath $InputDirectory -PathType Container)) { throw "Input directory not found: $InputDirectory" }
$files = @(Get-ChildItem -LiteralPath $InputDirectory -Filter $FilePattern -File | Sort-Object Name)
if ($files.Count -eq 0) { throw "No input CSV files matching '$FilePattern' were found in '$InputDirectory'." }
Initialize-OutputDirectory -Path $OutputDirectory

$plannedOutputs = @{}
foreach ($file in $files) {
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    $stem = $stem -replace '(?i)-webex-export$', ''
    $outputPath = Join-Path $OutputDirectory ($stem + '.csv')
    if ($plannedOutputs.ContainsKey($outputPath)) { throw "More than one input CSV maps to the same prepared output: $outputPath" }
    $plannedOutputs[$outputPath] = $true
    Assert-OutputFileAvailable -Path $outputPath -Force:$Force
}

$work = New-Object System.Collections.Generic.List[object]
$needsGraph = $false
foreach ($file in $files) {
    $rows = @(Import-Csv -LiteralPath $file.FullName)
    foreach ($row in $rows) {
        $work.Add([pscustomobject]@{ SourceFile = $file; Row = $row })
        if ([string]$row.Status -eq 'ReadyForTeamsLink') { $needsGraph = $true }
    }
}

if ($needsGraph) {
    Connect-WebexTeamsGraph -TenantId $GraphTenantId -AppId $GraphAppId -AuthMethod $GraphAuthMethod `
        -CertificateThumbprint $GraphCertificateThumbprint -ClientSecretPath $GraphClientSecretPath
}

try {
    foreach ($file in $files) {
        $outputRows = New-Object System.Collections.Generic.List[object]
        $fileWork = @($work | Where-Object { $_.SourceFile.FullName -eq $file.FullName })
        foreach ($entry in $fileWork) {
            $row = $entry.Row
            if ($null -eq $row.PSObject.Properties['ApprovedForUpdate']) {
                $row | Add-Member -NotePropertyName ApprovedForUpdate -NotePropertyValue 'No'
            }
            if ([string]$row.Status -eq 'ReadyForTeamsLink') {
                try {
                    foreach ($field in @('MailboxSmtp', 'EntraObjectId', 'ICalUid', 'IdentityStartUtc', 'GraphStartUtc', 'GraphEndUtc')) {
                        if ([string]::IsNullOrWhiteSpace([string]$row.$field)) { throw "Source row is missing required field '$field'." }
                    }
                    $externalId = Get-TeamsExternalId -MailboxSmtp ([string]$row.MailboxSmtp) `
                        -ICalUid ([string]$row.ICalUid) -IdentityStartUtc ([string]$row.IdentityStartUtc)
                    $teamsUrl = Get-TeamsJoinUrl -EntraObjectId ([string]$row.EntraObjectId) `
                        -ExternalId $externalId -Subject ([string]$row.Subject) `
                        -StartUtc ([string]$row.GraphStartUtc) -EndUtc ([string]$row.GraphEndUtc)
                    $row.TeamsJoinUrl = $teamsUrl
                    $row.Status = 'ReadyForUpdate'
                    $row.Note = 'Teams meeting created or reused. Exchange was not changed.'
                    if ([string]$row.HasModifiedOccurrences -eq 'True') {
                        $row.Note += ' This recurring series has modified occurrences; the update script changes the master only. Inspect exception invites after applying.'
                    }
                }
                catch {
                    $row.Status = 'GraphError'
                    $row.Note = $_.Exception.Message
                }
            }
            $outputRows.Add($row)
        }

        $stem = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $stem = $stem -replace '(?i)-webex-export$', ''
        $outputPath = Join-Path $OutputDirectory ($stem + '.csv')
        $outputRows | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding UTF8
        $readyCount = @($outputRows | Where-Object { $_.Status -eq 'ReadyForUpdate' }).Count
        Write-Host ("{0}: {1} rows ready for update. Prepared report: {2}" -f $file.Name, $readyCount, $outputPath)
    }
}
finally {
    if ($needsGraph) { Disconnect-WebexTeamsGraph }
}
