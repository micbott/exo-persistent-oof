# OOF Scheduled-Task Manager

Unattended **Out-of-Office (OOF)** management for Exchange Online mailboxes - PowerShell 7 and Windows Task Scheduler.

A daily scheduled task re-applies the OOF auto-reply on each configured
mailbox for a fixed date window, ensuring the setting survives any manual or
system reset. When the window expires the script disables the auto-reply and
the task deletes itself.

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
8. [File Overview](#file-overview)
9. [Logs](#logs)
10. [Troubleshooting](#troubleshooting)

---

## Prerequisites

| Component | Minimum version |
|---|---|
| **Windows** | Server 2016 / Windows 10+ |
| **PowerShell** | 7.x (`pwsh.exe`) |
| **ExchangeOnlineManagement** module | 3.x |

Install the EXO module system-wide (required for SYSTEM-level tasks):

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

1. Copy (or clone) the repository to `C:\Scripts\tasks\` on the target host.
2. Ensure the folder structure looks like this:

```
C:\Scripts\tasks\
    config.json              # <-- you create this (see below)
    exo-persistent-oof.ps1
    register-ooftask.ps1
    set-outofoffice.ps1
    .gitignore
    README.md
    Messages\
        default.txt          # fallback OOF message
        <usertag>.txt        # per-user OOF messages (optional)
    Logs\                    # created automatically on first run
```

3. Create `config.json` (next section).
4. Edit or add message files under `Messages\`.

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

OOF messages are plain-text (or HTML) files stored in `Messages\`.

The scripts resolve messages in this order:

1. **Per-user file** — `Messages\<usertag>.txt`
   (e.g. `Messages\john.doe.txt` for `john.doe@contoso.com`).
2. **Default file** — `Messages\default.txt`.
3. **Hardcoded fallback** — a generic one-liner built into the script.

The *usertag* is the local part of the mailbox address (before `@`),
lowercased, with any characters invalid for filenames replaced by `_`.

### Example message file

```
Thank you for your message. Please note that this mailbox is not being actively monitored.

For general enquiries, please refer to youremail@example.com

```

The same text is used for both internal and external recipients.

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
| 1 | **Register** | Prompts for mailbox, start date, duration (days), and daily run time. Computes end date, previews the task, and registers it on confirmation. Suggests a staggered run time (last existing trigger + 5 min). |
| 2 | **List** | Shows all `Set-OutOfOffice*` tasks with state, schedule, next/last run, result code, and the first 3 lines of the OOF message that would be applied. |
| 3 | **Remove** | Numbered multi-select deletion with confirmation. |
| 4 | **Run & Log** | Triggers a task immediately, waits for completion (up to 60 s), and displays the last 30 lines of its transcript log. |
| 5 | **Show OOF** | Connects to Exchange Online and displays the live `Get-MailboxAutoReplyConfiguration` output for a mailbox. |
| 6 | **Exit** | Leaves the menu. |

### Direct registration (without the menu)

```powershell
.\register-ooftask.ps1 `
    -Mailbox   "user@contoso.com" `
    -StartDate "2026-03-01" `
    -EndDate   "2026-03-31" `
    -RunTime   "00:30"
```

- `-RunTime` defaults to `00:30`.
- Stagger times across mailboxes (e.g. `00:30`, `00:35`, `00:40`) to avoid
  Exchange Online throttling.
- The task trigger expires one day after `-EndDate`, giving the script a final
  run to detect the window has passed and disable the auto-reply. The task is
  then automatically deleted by Task Scheduler.

---

## File Overview

| File | Purpose |
|---|---|
| `config.json` | Entra ID authentication settings (AppId, Org, Thumbprint). Not committed to git. |
| `set-outofoffice.ps1` | Core script called by the scheduled task. Connects to EXO, applies or disables the OOF auto-reply, writes a transcript log. |
| `register-ooftask.ps1` | Creates a per-user Windows Scheduled Task that runs `set-outofoffice.ps1` daily as SYSTEM. |
| `exo-persistent-oof.ps1` | Interactive 6-option controller menu for day-to-day management. |
| `Messages\default.txt` | Fallback OOF message used when no per-user file exists. |
| `Messages\<usertag>.txt` | Per-user OOF message (optional). |
| `Logs\task_output-<usertag>.log` | Transcript log for each mailbox (created automatically). |
| `.gitignore` | Excludes `Logs/`, `config.json`, and per-user message files from version control. Only `Messages\default.txt` is tracked. |

---

## Logs

Each mailbox gets its own transcript log at:

```
C:\Scripts\tasks\Logs\task_output-<usertag>.log
```

Logs are appended on every run. Use menu option **4 (Run & Log)** to trigger
a task and view its output, or inspect the file directly:

```powershell
Get-Content "C:\Scripts\tasks\Logs\task_output-john.doe.log" -Tail 50
```

---

## Troubleshooting

| Symptom | Cause / Fix |
|---|---|
| Task result `0x1` | Script error — check the transcript log for details. |
| `Connect-ExchangeOnline` fails with certificate error | Verify the thumbprint in `config.json` matches a cert in `Cert:\LocalMachine\My`. Ensure the cert's public key is uploaded to the Entra ID app. |
| `The term 'Connect-ExchangeOnline' is not recognized` | `ExchangeOnlineManagement` module not installed system-wide. Run `Install-Module ExchangeOnlineManagement -Scope AllUsers`. |
| OOF not applied / keeps resetting | Confirm the app has `Exchange.ManageAsApp` permission **and** an Exchange admin role assigned. Check that the mailbox address is correct. |
| `Access denied` when registering tasks | Run `exo-persistent-oof.ps1` (or `register-ooftask.ps1`) from an elevated (Administrator) PowerShell session. |
| Task does not appear after EndDate | Expected — tasks auto-delete one day after EndDate. |
