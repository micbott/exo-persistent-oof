#Requires -Modules ExchangeOnlineManagement
<#
.NAME
    exo-persistent-oof -- set-outofoffice.ps1

.SYNOPSIS
    Applies or disables a Scheduled OOF auto-reply on an Exchange Online mailbox.

.DESCRIPTION
    Connects to Exchange Online using certificate-based authentication against
    a Microsoft Entra ID App Registration (no user interaction required).
    The OOF is configured in "Scheduled" mode with immutable start/end dates
    that are passed in as parameters at task-registration time.

    On every run the script:
      - Loads the OOF message from a per-user text file under Messages\, falling
        back to Messages\default.txt if no per-user file exists. 
      - If the current date is past EndDate, disables the OOF and exits cleanly.
      - Otherwise (re-)applies the Scheduled OOF with the fixed date window,
        ensuring the auto-reply survives any manual or system reset.

    Each mailbox gets its own transcript log in Logs\task_output-<usertag>.log.

.PARAMETER Mailbox
    The email address of the mailbox to configure OOF on.

.PARAMETER StartDate
    The OOF start date/time in ISO 8601 format (e.g. 2026-02-25T00:00:00).

.PARAMETER EndDate
    The OOF end date/time in ISO 8601 format (e.g. 2026-03-27T23:59:00).

.PARAMETER DeclineMeetings
    Whether to auto-decline future meeting requests while OOF is active.
    Defaults to $true. Set to $false for shared/media mailboxes that should
    keep accepting meetings (e.g. ctao-media@…).

.EXAMPLE
    .\set-outofoffice.ps1 -Mailbox "user@contoso.com" -StartDate "2026-03-01T00:00:00" -EndDate "2026-03-31T23:59:00"

.EXAMPLE
    .\set-outofoffice.ps1 -Mailbox "ctao-media@contoso.com" -StartDate "2026-03-01T00:00:00" -EndDate "2026-03-31T23:59:00" -DeclineMeetings $false

.NOTES
    Author  : Michelangelo Bottura
    License : MIT
    Source  : https://github.com/micbott/exo-persistent-oof

    Prerequisites:
      1. A Microsoft Entra ID App Registration with the "Exchange.ManageAsApp"
         application permission granted and admin-consented.
      2. The app must be assigned the "Exchange Administrator" role (or a custom
         RBAC role with mailbox auto-reply permissions) in Exchange Online.
      3. A certificate installed in Cert:\LocalMachine\My on this host, with its
         public key (.cer) uploaded to the App Registration.
      4. The ExchangeOnlineManagement module installed system-wide:
         Install-Module ExchangeOnlineManagement -Scope AllUsers
#>
param(
    [Parameter(Mandatory)][string]   $Mailbox,
    [Parameter(Mandatory)][datetime] $StartDate,
    [Parameter(Mandatory)][datetime] $EndDate,
    [bool] $DeclineMeetings = $true
)

# ── Configuration ────────────────────────────────────────────────────────────
$configPath = Join-Path $PSScriptRoot "Config.json"
if (-not (Test-Path $configPath)) { Write-Error "Config file not found: $configPath"; exit 1 }
$config            = Get-Content -Path $configPath -Raw | ConvertFrom-Json
$AppId             = $config.AppId
$orgId             = $config.Organization
$CertificateThumb  = $config.CertificateThumbprint

# Derive a safe user tag from the mailbox (local part before @, lowercase,
# any chars invalid for file names replaced with _)
if ($Mailbox -notmatch '@') { Write-Error "Invalid mailbox: $Mailbox"; exit 1 }
$userTag = ($Mailbox.Split('@')[0]).ToLower() -replace '[\\/:*?"<>|]', '_'

# Ensure log directory exists
$logDir = "C:\scripts\tasks\Logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

Start-Transcript -Path "$logDir\task_output-$userTag.log" -Append

Write-Output "--- Run at $(Get-Date -Format 'u') for $Mailbox  [OOF window $StartDate -> $EndDate]  DeclineMeetings=$DeclineMeetings ---"

# ── OOF message content ─────────────────────────────────────────────────────
# Look for a per-user message file, fall back to default
$msgDir      = "C:\Scripts\tasks\Messages"
$userFile    = Join-Path $msgDir "$userTag.txt"
$defaultFile = Join-Path $msgDir "default.txt"

if (Test-Path $userFile) {
    Write-Output "Loading per-user message from $userFile"
    $InternalMessage = Get-Content -Path $userFile -Raw
}
elseif (Test-Path $defaultFile) {
    Write-Output "No per-user message found -- loading default from $defaultFile"
    $InternalMessage = Get-Content -Path $defaultFile -Raw
}
else {
    Write-Warning "No message file found in $msgDir -- using hardcoded fallback."
    $InternalMessage = "Thank you for your message. This mailbox is not being actively monitored."
}

$ExternalMessage = $InternalMessage
# ─────────────────────────────────────────────────────────────────────────────

try {
    # Connect to Exchange Online using certificate auth (no prompt)
    Connect-ExchangeOnline `
        -AppId            $AppId `
        -CertificateThumbprint $CertificateThumb `
        -Organization     $orgId `
        -ShowBanner:$false

    # If we are past the end date, disable OOF and exit cleanly
    if ((Get-Date) -gt $EndDate) {
        Write-Output "EndDate ($EndDate) has passed -- disabling OOF on $Mailbox."
        Set-MailboxAutoReplyConfiguration -Identity $Mailbox -AutoReplyState Disabled
        Write-Output "OOF disabled successfully on $Mailbox."
    }
    else {
        # Build the parameters for Set-MailboxAutoReplyConfiguration
        $params = @{
            Identity         = $Mailbox
            AutoReplyState   = "Scheduled"
            StartTime        = $StartDate
            EndTime          = $EndDate
            InternalMessage  = $InternalMessage
            ExternalMessage  = $ExternalMessage
            ExternalAudience = "All"
            AutoDeclineFutureRequestsWhenOOF = $DeclineMeetings
        }

        Set-MailboxAutoReplyConfiguration @params

        Write-Output "OOF auto-reply successfully set on $Mailbox (Scheduled $StartDate -> $EndDate, DeclineMeetings=$DeclineMeetings)"
    }
}
catch {
    Write-Error "Failed to set OOF: $_"
    exit 1
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Stop-Transcript -ErrorAction SilentlyContinue
}
