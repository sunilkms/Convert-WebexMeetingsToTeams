[CmdletBinding()]
param(
    [string] $InputCsv = '.\pilot-mailboxes.csv',
    [string] $OutputCsv = '.\targets.csv',
    [switch] $Force
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $InputCsv -PathType Leaf)) {
    throw "Input CSV not found: $InputCsv"
}
if ((Test-Path -LiteralPath $OutputCsv) -and -not $Force) {
    throw "Output already exists: $OutputCsv. Review it first, or pass -Force to replace it."
}

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop

$inputRows = @(Import-Csv -LiteralPath $InputCsv)
if ($inputRows.Count -eq 0) {
    throw "Input CSV has no mailbox rows: $InputCsv"
}

foreach ($row in $inputRows) {
    if (-not $row.MailboxSmtp -or [string]::IsNullOrWhiteSpace([string]$row.MailboxSmtp)) {
        throw 'Every input row must contain MailboxSmtp.'
    }
    if ($row.PSObject.Properties.Name -notcontains 'EntraUserPrincipalName') {
        throw 'Input CSV must have MailboxSmtp and optional EntraUserPrincipalName columns.'
    }
}

# Delegated, read-only access is used only to map SMTP/UPN values to Entra object IDs.
# This is separate from the POC app permission used to create Teams meetings.
Connect-MgGraph -Scopes 'User.ReadBasic.All' -ContextScope Process -NoWelcome

$resolved = New-Object System.Collections.Generic.List[object]
$errors = New-Object System.Collections.Generic.List[string]
$seenSmtp = @{}
$seenIds = @{}

try {
    foreach ($row in $inputRows) {
        $smtp = ([string]$row.MailboxSmtp).Trim()
        $upn = ([string]$row.EntraUserPrincipalName).Trim()

        if ($seenSmtp.ContainsKey($smtp)) {
            $errors.Add("Duplicate mailbox SMTP address in input: $smtp")
            continue
        }
        $seenSmtp[$smtp] = $true

        # Use an explicitly supplied UPN when SMTP is an alias or differs from the sign-in name.
        $lookupProperty = if ($upn) { 'userPrincipalName' } else { 'mail' }
        $lookupValue = if ($upn) { $upn } else { $smtp }
        $escapedValue = $lookupValue.Replace("'", "''")
        $filter = "$lookupProperty eq '$escapedValue'"

        try {
            $users = @(Get-MgUser -Filter $filter -Property @('id', 'mail', 'userPrincipalName') -Top 2 -ErrorAction Stop)
        }
        catch {
            $errors.Add("Lookup failed for '$smtp' using $lookupProperty '$lookupValue': $($_.Exception.Message)")
            continue
        }

        if ($users.Count -eq 0 -and -not $upn) {
            # Common hybrid configuration: Entra UPN equals the on-prem primary SMTP address.
            $upnFilter = "userPrincipalName eq '$escapedValue'"
            try {
                $users = @(Get-MgUser -Filter $upnFilter -Property @('id', 'mail', 'userPrincipalName') -Top 2 -ErrorAction Stop)
            }
            catch {
                $errors.Add("UPN fallback lookup failed for '$smtp': $($_.Exception.Message)")
                continue
            }
        }

        if ($users.Count -ne 1) {
            $errors.Add("Expected one Entra user for '$smtp'; found $($users.Count). Check the primary SMTP address or supply EntraUserPrincipalName.")
            continue
        }

        $user = $users[0]
        if (-not $user.Id) {
            $errors.Add("Entra returned no object ID for '$smtp'.")
            continue
        }
        if ($seenIds.ContainsKey([string]$user.Id)) {
            $errors.Add("More than one input address resolved to Entra object ID $($user.Id). Check the input list for aliases or duplicates.")
            continue
        }
        $seenIds[[string]$user.Id] = $true

        $resolved.Add([pscustomobject]@{
            MailboxSmtp = $smtp
            EntraObjectId = [string]$user.Id
        })
        Write-Host ("Resolved {0} -> {1} ({2})" -f $smtp, $user.Id, $user.UserPrincipalName)
    }

    if ($errors.Count -gt 0) {
        $errors | ForEach-Object { Write-Error $_ -ErrorAction Continue }
        throw "No output written: fix $($errors.Count) lookup/input issue(s), then rerun."
    }

    $resolved | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Wrote $($resolved.Count) target(s) to: $OutputCsv"
}
finally {
    if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
        if (Get-MgContext) {
            Disconnect-MgGraph | Out-Null
        }
    }
}
