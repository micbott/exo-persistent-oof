# OOF Scheduled-Task Manager

Unattended **Out-of-Office (OOF)** management for Exchange Online mailboxes - PowerShell 7 and Windows Task Scheduler.

A daily scheduled task re-applies the OOF auto-reply on each configured
mailbox for a fixed date window, ensuring the setting survives any manual or
system reset. When the window expires the script disables the auto-reply and
Task Scheduler then deletes the expired task.

Because Exchange Online only sends one auto-reply per sender for the lifetime
of a Scheduled OOF window, re-applying the configuration daily effectively
resets the "already replied" tracker. This means each sender receives the OOF
message **once per day** instead of once for the entire period. Particularly
useful for mailboxes of departing employees, where correspondents need a
repeated reminder that the mailbox is no longer monitored.

---

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Microsoft Entra ID App Registration](#microsoft-entra-id-app-registration)
3. [Certificate Setup](#certificate-setup)
4. [Installation](#installation)
5. [Configuration — config.json](#configuration--configjson)
6. [Message Files](#message-files)
7. [Usage](#usage)
8. [Testing a Deployment](#testing-a-deployment)
9. [File Overview](#file-overview)
10. [Logs](#logs)
11. [Troubleshooting](#troubleshooting)

---

## Prerequisites

| Component | Minimum version |
|---|---|
| **Windows** | Server 2016 / Windows 10+ |
| **PowerShell** | 7.x, installed to the default path (`C:\Program Files\PowerShell\7\pwsh.exe`, on the system PATH) |
| **ExchangeOnlineManagement** module | 3.x, installed for **AllUsers** |
| **Rights** | Local Administrator (to install the module and to register, run or remove scheduled tasks) |

Install the EXO module system-wide from an **elevated** shell. This is
required: the tasks run as SYSTEM, which cannot see CurrentUser installs, and
`register-ooftask.ps1` refuses to register a task otherwise.

```powershell
Install-Module ExchangeOnlineManagement -Scope AllUsers -Force
```

---

## Microsoft Entra ID App Registration

The scripts authenticate to Exchange Online using a **certificate-based
application identity** — no interactive sign-in and no stored passwords.

### 1. Register the application

1. Open the [Azure portal](https://portal.azure.com) >
   **Microsoft Entra ID** > **App registrations** > **New registration**.
2. Name: e.g. `EXO-OOF-Automation`.
3. Supported account types: **Accounts in this organizational directory only**.
4. Redirect URI: leave blank.
5. Click **Register** and note the **Application (client) ID**.

### 2. Grant API permissions

1. In the app registration, go to **API permissions** > **Add a permission**.
2. Select **APIs my organization uses** and search for **Office 365 Exchange
   Online** (resource ID `00000002-0000-0ff1-ce00-000000000000`).
3. Choose **Application permissions** >
   **Exchange.ManageAsApp** > **Add permissions**.
4. Click **Grant admin consent for \<tenant\>**.

### 3. Assign an Exchange Online admin role

The app must hold an Exchange role that allows managing mailbox auto-reply
settings. The simplest approach:

1. Go to **Microsoft Entra ID** > **Roles and administrators**.
2. Open the **Exchange Administrator** role.
3. **Add assignments** > select the app registration you just created.

> **Tip:** For least-privilege, create a custom RBAC role in Exchange Online
> scoped to `Set-MailboxAutoReplyConfiguration` and
> `Get-MailboxAutoReplyConfiguration`, then assign the app's service principal
> to that role.

### 4. Upload the certificate

See the next section for how to create the certificate. Once you have the
`.cer` (public key) file:

1. In the app registration, go to **Certificates & secrets** >
   **Certificates** > **Upload certificate**.
2. Upload the `.cer` file.

---

## Certificate Setup

Generate a self-signed certificate on the host that will run the scheduled
tasks, and install it in the **Local Machine** store (so the SYSTEM account
can access it):

```powershell
$cert = New-SelfSignedCertificate `
    -Subject        "CN=EXO-OOF-Automation" `
    -CertStoreLocation Cert:\LocalMachine\My `
    -KeyExportPolicy Exportable `
    -KeySpec         Signature `
    -KeyLength       2048 `
    -NotAfter        (Get-Date).AddYears(3)

# Export the public key to upload to Entra ID
Export-Certificate -Cert $cert -FilePath "$env:USERPROFILE\Desktop\exo-oof.cer"

# Note the thumbprint — you'll need it for config.json
$cert.Thumbprint
```

Upload `exo-oof.cer` to the App Registration as described above.

---

## Installation

1. Copy `set-outofoffice.ps1`, `register-ooftask.ps1`, `exo-persistent-oof.ps1`
   and `messages\default.txt` to `C:\Scripts\tasks\`, and create an empty
   `logs\` folder. Cloning the whole repository there also works (the clone
   already contains `logs\`).
2. The folder structure must look like this:

```
C:\Scripts\tasks\
    config.json              # <-- you create this (see below)
    exo-persistent-oof.ps1
    register-ooftask.ps1
    set-outofoffice.ps1
    messages\
        default.txt          # fallback OOF message
        <usertag>.txt        # per-user OOF messages (optional)
    logs\                    # transcripts; set-outofoffice.ps1 also creates it if missing
```

> The install path `C:\Scripts\tasks` is **hardcoded** in all three scripts;
> installing anywhere else needs code changes. Folder names are
> case-insensitive on Windows (the scripts refer to `Messages` and `Logs`).

3. Create `config.json` (next section).
4. Edit or add message files under `messages\`.
5. Run `.\test-deployment.ps1` from your repo checkout to verify the
   installation (see [Testing a Deployment](#testing-a-deployment)).

---

## Configuration — config.json

Create a file named `config.json` in the `C:\Scripts\tasks\` directory with
the three values from your Entra ID App Registration:

```json
{
    "AppId": "<Application (client) ID>",
    "Organization": "<tenant>.onmicrosoft.com",
    "CertificateThumbprint": "<certificate thumbprint>"
}
```

| Key | Where to find it |
|---|---|
| `AppId` | Azure portal > App registrations > your app > **Overview** > Application (client) ID |
| `Organization` | Your tenant's `.onmicrosoft.com` domain (e.g. `contoso.onmicrosoft.com`) |
| `CertificateThumbprint` | The SHA-1 thumbprint of the certificate installed in `Cert:\LocalMachine\My` |

> **Security note:** `config.json` contains authentication identifiers.
> It is listed in `.gitignore` and must **not** be committed to version
> control. Treat it as a secret.

---

## Message Files

OOF messages are plain-text (or HTML) files stored in `C:\Scripts\tasks\messages\`.

The scripts resolve messages in this order:

1. **Per-user file** — `messages\<usertag>.txt`
   (e.g. `messages\john.doe.txt` for `john.doe@contoso.com`).
2. **Default file** — `messages\default.txt`.
3. **Hardcoded fallback** — a generic one-liner built into the script.

The *usertag* is the local part of the mailbox address (before `@`),
lowercased, with any characters invalid for filenames replaced by `_`.

### Example message file

```
Thank you for your message. Please note that this mailbox is not being actively monitored.

For general enquiries, please refer to youremail@example.com

```

The same text is used for both internal and external recipients, and replies
go to all external senders (`ExternalAudience All`). The file is re-read on
every run, so edits apply from the next run.

---

## Usage

All management is done through the interactive controller, which must be run
in an **elevated** (Administrator) PowerShell 7 session:

```powershell
pwsh -ExecutionPolicy Bypass -File "C:\Scripts\tasks\exo-persistent-oof.ps1"
```

### Menu options

| # | Option | Description |
|---|--------|-------------|
| 1 | **Register** | Prompts for mailbox, start date (default today), duration in days (default 30), daily run time, and whether to auto-decline meetings (default Y). Computes the end date (start + duration), previews the task and message, warns if a task with that name exists, and calls `register-ooftask.ps1` on confirmation. Suggested run time is the latest existing trigger + 5 min (`00:30` if none). Reports *Registration failed* if the register script exits with an error. |
| 2 | **List** | Shows all `Set-OutOfOffice*` tasks with state, mailbox, OOF window, daily run time, DeclineMeetings, next/last run, last result (hex), and the first 3 non-blank lines of the message that would be applied. |
| 3 | **Remove** | Numbered multi-select deletion (comma-separated numbers, `A` = all, `Q` = cancel) with confirmation. |
| 4 | **Run & Log** | Triggers a task immediately, waits for completion (up to 60 s), and displays the last 30 lines of its transcript log. |
| 5 | **Show OOF** | Pick a mailbox that has a task; connects to Exchange Online with `config.json` and shows AutoReplyState, Start/EndTime, ExternalAudience, AutoDeclineFutureRequestsWhenOOF and the first 5 lines of the internal/external message. |
| 6 | **Exit** | Leaves the menu. |

### Direct registration (without the menu)

```powershell
.\register-ooftask.ps1 `
    -Mailbox   "user@contoso.com" `
    -StartDate "2026-03-01" `
    -EndDate   "2026-03-31" `
    -RunTime   "00:30"
```

- Run from an elevated session. The task always points at
  `C:\Scripts\tasks\Set-OutOfOffice.ps1`, wherever you run this from.
- `-RunTime` defaults to `00:30`.
- Stagger times across mailboxes (e.g. `00:30`, `00:35`, `00:40`) to avoid
  Exchange Online throttling.
- `-DeclineMeetings` defaults to `$true` (future meeting requests are
  auto-declined while OOF is active). Pass `-DeclineMeetings $false` for shared
  mailboxes that must keep accepting meetings.
- When a script is called through `pwsh -File` (as the scheduled task does),
  a `[bool]` must use colon syntax: `-DeclineMeetings:$false`. With a space,
  the value is passed as a string and fails. The task action is built this way
  automatically.
- `-EndDate` is exclusive: a date without a time means 00:00 on that day, so
  the OOF ends as that day begins. The run on EndDate detects the window has
  passed and disables the auto-reply. The trigger expires at 00:00 the day
  after EndDate, and Task Scheduler deletes the expired task about one day
  later.
- Fails before registering if ExchangeOnlineManagement is not installed for
  AllUsers, or if StartDate is not before EndDate.
- Add `-WhatIf` to run every check (including the AllUsers module check) and
  show the task that would be registered, without registering it.

---

## Testing a Deployment

After copying updated scripts to `C:\Scripts\tasks\`, run `test-deployment.ps1`
**from the repo checkout** (it compares the repo with the deployed folder):

```powershell
# Offline checks only -- no Exchange Online contact
.\test-deployment.ps1

# Offline checks, then run one task live and check its exit code and log
.\test-deployment.ps1 -LiveTask ctao-media

# Show what the live run would do, without starting the task
.\test-deployment.ps1 -LiveTask ctao-media -WhatIf
```

Parameters: `-DeployPath` (default `C:\Scripts\tasks`), `-LiveTask <usertag>`
(runs `Set-OutOfOffice-<usertag>`), `-AllowDeclineMeetings`,
`-TimeoutSeconds` (default 120, range 10–1800), `-WhatIf`.

Offline checks: `pwsh.exe` at `C:\Program Files\PowerShell\7\pwsh.exe`,
ExchangeOnlineManagement installed for AllUsers, each of the three deployed
scripts identical to the repo (case-sensitive, line endings ignored) and
parseable, and `config.json` / `messages\default.txt` / `logs\` present. It
also reports every `Set-OutOfOffice-*` task's state, last result and
DeclineMeetings flag; a non-zero last result is a warning, not a failure.

The live run applies the OOF settings in Exchange Online exactly like the
nightly run. It passes only if the task finishes within the timeout with
result 0 **and** this run's transcript contains a success line; those lines
are printed at the end. It is skipped if any offline check fails, and refused
for tasks that decline meetings unless `-AllowDeclineMeetings` is passed,
because such a run can decline real meeting requests. A task without a
`-DeclineMeetings` argument counts as declining (that is the parameter's
default). With `-WhatIf` the offline checks still run. The script exits `0`
when no check failed and `1` otherwise, including when the deploy folder is
missing.

---

## File Overview

| File | Purpose |
|---|---|
| `config.json` | Entra ID authentication settings (AppId, Organization, CertificateThumbprint). Not committed to git. |
| `set-outofoffice.ps1` | Core script called by the scheduled task. Connects to EXO, applies or disables the OOF auto-reply, writes a transcript log. |
| `register-ooftask.ps1` | Creates a per-user Windows Scheduled Task that runs `set-outofoffice.ps1` daily as SYSTEM. Supports `-WhatIf`. |
| `exo-persistent-oof.ps1` | Interactive 6-option controller menu for day-to-day management. |
| `test-deployment.ps1` | Run from the repo: checks the deployed scripts match and can run as SYSTEM, optionally runs one task live. Need not be deployed. |
| `messages\default.txt` | Fallback OOF message used when no per-user file exists. |
| `messages\<usertag>.txt` | Per-user OOF message (optional, not committed). |
| `logs\task_output-<usertag>.log` | Transcript log for each mailbox (not committed). |
| `logs\.gitkeep` | Keeps the empty `logs` folder in git. |
| `README.md` | This document. |
| `CLAUDE.md` | Project notes for Claude Code (architecture, conventions, testing rules). |
| `LICENSE` | MIT licence. |
| `.gitignore` | Local only (it ignores itself, so it is not committed). Excludes `logs/*` (except `.gitkeep`), `config.json`, `messages/*` (except `default.txt`), `.github/` and `CLAUDE.local.md`. |

---

## Logs

Each mailbox gets its own transcript log at:

```
C:\Scripts\tasks\Logs\task_output-<usertag>.log
```

`<usertag>` is derived as described in [Message Files](#message-files). Each
run appends a `--- Run at <time> for <mailbox> [OOF window ...]
DeclineMeetings=... ---` header. Failures inside the run, including a failed
`Import-Module ExchangeOnlineManagement` or an Exchange Online error, are
logged as `Failed to set OOF: <error>` and the task exits with result `0x1`.
Errors about a missing `config.json` or an invalid mailbox occur before the
transcript starts and are not logged. Logs are never rotated.

Use menu option **4 (Run & Log)** to trigger a task and view its output, or
inspect the file directly:

```powershell
Get-Content "C:\Scripts\tasks\Logs\task_output-john.doe.log" -Tail 50
```

`test-deployment.ps1 -LiveTask` prints only the lines written by its own run.

---

## Troubleshooting

| Symptom | Cause / Fix |
|---|---|
| Task result `0x1` | Script error — check the transcript log for a `Failed to set OOF: ...` line. |
| Task result `0x1` but no new lines in the log | `config.json` is missing from `C:\Scripts\tasks` or the mailbox argument is invalid; these errors occur before the transcript starts. Run the task's command manually in an elevated pwsh to see them. |
| `Connect-ExchangeOnline` fails with certificate error | Verify the thumbprint in `config.json` matches a cert in `Cert:\LocalMachine\My`. Ensure the cert's public key is uploaded to the Entra ID app. |
| Log shows `Failed to set OOF: The specified module 'ExchangeOnlineManagement' was not loaded` | Module not visible to SYSTEM. From an elevated shell run `Install-Module ExchangeOnlineManagement -Scope AllUsers`. |
| `register-ooftask.ps1`: `ExchangeOnlineManagement is not installed for AllUsers (checked: ...)` | Same cause: the pre-flight check found no module under Program Files. Install with `-Scope AllUsers`, then register again (use `-WhatIf` to re-check without registering). |
| Menu says *Registration failed (exit code 1)* | Read the error printed above it (e.g. the AllUsers check, an invalid mailbox, or StartDate not before EndDate). |
| Log shows `Failed to set OOF: ... couldn't be found` | The mailbox address does not exist in Exchange Online. Fix the address and register the task again. |
| OOF not applied / keeps resetting | Confirm the app has `Exchange.ManageAsApp` permission **and** an Exchange admin role assigned. Check that the mailbox address is correct. |
| `Access denied` when registering tasks | Run `exo-persistent-oof.ps1` (or `register-ooftask.ps1`) from an elevated (Administrator) PowerShell session. The menu shows `WARNING: Not running as Administrator` when not elevated. |
| `test-deployment.ps1`: `Exists: Logs` fails on a fresh install | Create `C:\Scripts\tasks\logs` (otherwise it is only created by the first task run). |
| Task does not appear after EndDate | Expected: the trigger expires the day after EndDate and Task Scheduler deletes the expired task about a day later. |
