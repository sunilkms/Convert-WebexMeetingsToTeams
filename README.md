# Webex-to-Teams Meeting URL POC

This project is a proof of concept for replacing Webex join links in selected meeting invitations with Microsoft Teams join links. It targets a hybrid environment where Exchange Server Subscription Edition hosts the mailboxes on premises and users are synchronized to Microsoft Entra ID.

## What it does

The PowerShell script uses two services:

1. **Exchange Web Services (EWS)** reads the organizer's on-premises calendar and updates eligible meeting bodies and, when the join link is there, the Location field.
2. **Microsoft Graph** creates or retrieves a Teams online meeting for the organizer and returns its join URL.

The script inserts the Teams join link into the existing invite, replaces the selected Webex URL in the body and/or Location field, and sends an updated meeting request to attendees.

```mermaid
flowchart LR
    A[Target mailbox list] --> B[POC PowerShell script]
    B -->|Read and update calendar| C[On-prem Exchange SE via EWS]
    B -->|Create or retrieve join link| D[Teams via Microsoft Graph]
    D -->|Teams join URL| B
    C -->|Updated invitation| E[Attendees]
```

## Important behavior

Graph's `onlineMeeting:createOrGet` creates a standalone meeting that does not appear on the organizer's calendar. The script updates the existing on-prem Exchange event separately through EWS. This POC therefore changes the invitation's join instructions; it does not create a fully calendar-backed Teams meeting or migrate Webex meeting IDs, dial-in details, registration, recordings, attendance data, or policies. See [Microsoft Graph `createOrGet`](https://learn.microsoft.com/en-us/graph/api/onlinemeeting-createorget?view=graph-rest-1.0).

The included script processes organizer-owned, non-cancelled single meetings and recurring series. It scans both the item body and Location field. It prefers scheduled Webex join links in the `j.php?MTID=...` form; URLs with the same MTID in that join-link group are treated as variants and replaced together. Other URLs, including `globalcallin.php` links, are left untouched. If there is no `j.php` join link with an MTID, the script uses a unique MTID or proceeds only when exactly one distinct Webex URL exists; multiple different MTIDs in the selected candidate group or multiple URLs without an MTID require manual review. The report's `WebexUrlSource` column identifies where selected links appeared, and `WebexUrlDiagnostics` lists each detected URL, its type, source, and parsed MTID—including duplicates—to help troubleshoot matching. Treat reports as sensitive because they contain meeting join URLs and MTIDs. For a recurring series, it binds to and updates the series master once, so the series uses one shared Teams URL. It skips attendee copies, items without a Webex URL, ambiguous multiple-link items, and events that already contain a Teams join URL. A modified occurrence (an exception) can have its own invite body; the script flags such series but does not rewrite exception bodies independently. Review those exceptions manually after applying, along with generated invite content and any Webex-specific instructions that remain.

The recurring series uses a standalone Graph meeting object. Microsoft documents a 60-day expiry after the meeting start/end that can be extended by updating or activating the meeting before expiry. Validate long recurring gaps over 60 days separately; they may require a refreshed link. [Graph onlineMeeting lifetime](https://learn.microsoft.com/en-us/graph/api/resources/onlinemeeting?view=graph-rest-1.0)

## Project files

| File | Description |
| --- | --- |
| `README.md` | Project overview and quick-start guidance. |
| `POC-Webex-to-Teams-Meeting-URL.md` | Detailed requirements, permission setup, Exchange scoping, runbook, validation steps, and rollback instructions. |
| `scripts/Convert-WebexMeetingsToTeams.ps1` | Script with `Preview`, `Apply`, and `Rollback` modes and a Windows encrypted EWS credential cache. |
| `scripts/New-TargetsCsv.ps1` | Resolves pilot SMTP addresses to Entra object IDs and writes `targets.csv`. |
| `pilot-mailboxes.csv` | Input template for pilot mailbox SMTP addresses and optional Entra UPNs. |
| `targets.csv` | Output template consumed by the meeting conversion script. |

## Requirements and permissions

- Exchange SE hybrid deployment with EWS reachable over HTTPS from the script runner. HMA supports EWS, but the script still needs an authentication method accepted by the EWS endpoint. The included example uses Windows authentication and a dedicated service account. See [HMA protocol support](https://learn.microsoft.com/en-us/microsoft-365/enterprise/configure-exchange-server-for-hybrid-modern-authentication).
- The EWS service account must have the Exchange `ApplicationImpersonation` role scoped to pilot mailboxes. Do not grant unscoped impersonation for this POC.
- A single-tenant Entra app registration with a certificate and Microsoft Graph **application** permission `OnlineMeetings.ReadWrite.All`, with admin consent.
- For the POC, create a self-signed certificate in `Cert:\CurrentUser\My` while signed in as the Windows account that will run the script; upload only its public `.cer` to the app registration and pass its thumbprint to the script. The detailed generation, upload, and app-only connection check are in [the runbook](POC-Webex-to-Teams-Meeting-URL.md). Use your organization's PKI process for production.
- A Teams application access policy assigned only to pilot organizers. Graph requires this policy when application permission is used to create meetings on behalf of users.
- The EWS credential cache uses Windows DPAPI through `Export-Clixml`. The first run prompts once and saves it under `%LOCALAPPDATA%\WebexTeamsPOC\ews-credential.xml`; later runs must use the same Windows account on the same computer. [Microsoft Export-Clixml credential encryption](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/export-clixml?view=powershell-5.1)
- The EWS request defaults to direct access, which fits this on-premises internal EWS endpoint and avoids the system proxy that returned HTTP 407. Graph calls continue to use the normal Windows network settings. To use the proxy for EWS, set `-EwsProxyMode System` or `-EwsProxyMode SystemDefaultCredentials` as described in the runbook.
- To generate the target list, an interactive delegated Graph lookup using `User.ReadBasic.All`. This read-only permission is used only by the helper script, not by the POC meeting creation app. [Graph list users permissions](https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0)
- Pilot organizers synchronized to Entra ID and enabled/licensed for Teams.
- Windows PowerShell 5.1, the approved EWS Managed API 2.2 assembly, and the `Microsoft.Graph.Authentication` and `Microsoft.Graph.Users` modules.
- Populate `pilot-mailboxes.csv` with each organizer's SMTP address. If it differs from the Entra UPN, supply the UPN in the second column:

```csv
MailboxSmtp,EntraUserPrincipalName
pilot.organizer@contoso.com,
```

Generate `targets.csv` by running the resolver. It signs in interactively, looks up each user, and exports `MailboxSmtp,EntraObjectId`. The existing POC app permissions remain unchanged.

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Users -Scope CurrentUser
.\scripts\New-TargetsCsv.ps1 -InputCsv ".\pilot-mailboxes.csv" -OutputCsv ".\targets.csv" -Force
```

The resolver requires delegated `User.ReadBasic.All` to read basic user profiles; Microsoft documents that this profile includes `id`, `mail`, and `userPrincipalName`. [Graph user-list permissions](https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0), [basic profile properties](https://learn.microsoft.com/en-us/graph/permissions-reference)

The meeting-conversion script does not need Graph calendar or mail permissions. The separate target resolver uses only delegated read-only `User.ReadBasic.All`. For Teams' native calendar integration with on-premises mailboxes, there are additional hybrid, published EWS/Autodiscover, and OAuth prerequisites; see Microsoft's [Teams and Exchange integration requirements](https://learn.microsoft.com/en-us/microsoftteams/exchange-teams-interact).

## Quick start

First, follow the detailed [POC runbook](POC-Webex-to-Teams-Meeting-URL.md) to register the Graph app, scope Exchange impersonation, assign the Teams application access policy, prepare the certificate, and confirm EWS connectivity.

### Preview

Preview reads meetings and writes a report. It does not call Graph or update Exchange. The first run prompts for the EWS credential and saves it encrypted; later runs reuse that cache. See the credential-cache note in the runbook if the password changes or the script will run under a different Windows account.

```powershell
.\scripts\Convert-WebexMeetingsToTeams.ps1 `
  -Mode Preview `
  -EwsUrl "https://mail.contoso.com/EWS/Exchange.asmx" `
  -EwsManagedApiDll "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll" `
  -TargetsCsv ".\targets.csv" `
  -WindowStartUtc "2026-11-05T00:00:00Z" `
  -WindowEndUtc "2026-11-20T00:00:00Z" `
  -ReportPath ".\preview.csv"
```

### Apply

After reviewing the preview, use a one-organizer CSV for the first apply. Add the Graph tenant ID, app client ID, certificate thumbprint, and a new backup path. The script creates/reuses the Teams meeting, backs up the original invite body and Location, updates Exchange, and sends the attendee update.

```powershell
.\scripts\Convert-WebexMeetingsToTeams.ps1 `
  -Mode Apply `
  -EwsUrl "https://mail.contoso.com/EWS/Exchange.asmx" `
  -EwsManagedApiDll "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll" `
  -TargetsCsv ".\targets-one.csv" `
  -WindowStartUtc "2026-11-05T00:00:00Z" `
  -WindowEndUtc "2026-11-20T00:00:00Z" `
  -GraphTenantId "<TENANT-ID>" `
  -GraphAppId "<GRAPH-APP-CLIENT-ID>" `
  -GraphCertificateThumbprint "<CERTIFICATE-THUMBPRINT>" `
  -BackupPath ".\teams-poc-backup-20261105-120000.jsonl" `
  -ReportPath ".\apply.csv"
```

### Rollback

Rollback restores the original meeting body and Location from the JSONL backup and sends another meeting update. It skips restoration if either field no longer matches the script's saved post-apply value. It does not delete the standalone Teams online meeting object. Protect the backup because it contains original meeting text and Location.

```powershell
.\scripts\Convert-WebexMeetingsToTeams.ps1 `
  -Mode Rollback `
  -EwsUrl "https://mail.contoso.com/EWS/Exchange.asmx" `
  -EwsManagedApiDll "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll" `
  -BackupPath ".\teams-poc-backup-20261105-120000.jsonl" `
  -ReportPath ".\rollback.csv"
```

## Validation checklist

1. Start with one organizer and one future, single-instance meeting with a test attendee.
2. Review the `Preview` report, especially `OriginalWebexUrl`, `WebexUrlSource`, `WebexUrlDiagnostics`, and `Note`. Confirm that a selected MTID URL is the intended join link and understand which other Webex URLs will remain untouched.
3. Apply the change and confirm the Teams URL appears in the organizer's Exchange event.
4. Confirm the attendee receives the updated invite and can join the Teams meeting.
5. Repeat with a simple recurring series without modified occurrences and confirm the master update reaches multiple occurrences. If `HasModifiedOccurrences` is true, inspect and update each exception invite manually; the script only edits the recurring master.
6. Check lobby, anonymous-join, and external attendee behavior against Teams meeting policies.
7. Test rollback before widening the pilot.

For the complete procedure, troubleshooting notes, and cleanup commands, use the [detailed runbook](POC-Webex-to-Teams-Meeting-URL.md).
