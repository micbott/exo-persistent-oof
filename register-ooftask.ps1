<#
.NAME
    exo-persistent-oof -- register-ooftask.ps1

.SYNOPSIS
    Registers a daily Windows Scheduled Task that runs set-outofoffice.ps1
    for a given mailbox over a fixed date range. Must be run elevated.

.DESCRIPTION
    Creates (or overwrites) a per-user scheduled task named
    "Set-OutOfOffice-<usertag>" that invokes set-outofoffice.ps1 with the
    specified -Mailbox, -StartDate, and -EndDate parameters baked in.

    The task triggers expire one day after EndDate so the OOF script gets a
    final run to detect the window has passed and disable the auto-reply.
    After expiry the task is automatically deleted by Task Scheduler.

    A -RunTime parameter controls the daily execution time (default 00:30).
    When multiple mailboxes are registered, stagger the times (e.g. 00:30,
    00:35, 00:40) to avoid Exchange Online throttling.

.PARAMETER Mailbox
    The email address of the mailbox to configure OOF on.

.PARAMETER StartDate
    The OOF start date (e.g. 2026-02-25).

.PARAMETER EndDate
    The OOF end date (e.g. 2026-03-27). The task will expire 1 day after this.

.PARAMETER RunTime
    The daily trigger time in HH:mm format (default "00:30"). Stagger across
    mailboxes to avoid throttling.

.PARAMETER DeclineMeetings
    Whether to auto-decline future meeting requests while OOF is active.
    Defaults to $true. Set to $false for shared/media mailboxes that should
    keep accepting meetings (e.g. ctao-media@…).

.EXAMPLE
    .\register-ooftask.ps1 -Mailbox "user@contoso.com" -StartDate "2026-02-25" -EndDate "2026-03-27"

.EXAMPLE
    .\register-ooftask.ps1 -Mailbox "user@contoso.com" -StartDate "2026-02-25" -EndDate "2026-03-27" -RunTime "01:15"

.EXAMPLE
    .\register-ooftask.ps1 -Mailbox "ctao-media@contoso.com" -StartDate "2026-02-25" -EndDate "2026-03-27" -DeclineMeetings $false

.NOTES
    Author  : Michelangelo Bottura
    License : MIT
    Source  : https://github.com/micbott/exo-persistent-oof
#>
param(
    [Parameter(Mandatory)][string]   $Mailbox,
    [Parameter(Mandatory)][datetime] $StartDate,
    [Parameter(Mandatory)][datetime] $EndDate,
    [string] $RunTime = "00:30",
    [bool]   $DeclineMeetings = $true
)

# ── Configuration ────────────────────────────────────────────────────────────
# Derive a safe user tag from the mailbox (local part before @, lowercase,
# any chars invalid for task names replaced with _)
if ($Mailbox -notmatch '@') { Write-Error "Invalid mailbox: $Mailbox"; exit 1 }
if ($StartDate -ge $EndDate) { Write-Error "StartDate ($StartDate) must be before EndDate ($EndDate)."; exit 1 }
$userTag = ($Mailbox.Split('@')[0]).ToLower() -replace '[\\/:*?"<>|]', '_'

$TaskName        = "Set-OutOfOffice-$userTag"
$TaskDescription = "Applies OOF auto-reply on $Mailbox daily at $RunTime ($($StartDate.ToString('yyyy-MM-dd')) -> $($EndDate.ToString('yyyy-MM-dd')))"
$ScriptPath      = "C:\Scripts\tasks\Set-OutOfOffice.ps1"

# The task trigger expires 1 day after EndDate so the script can do a final
# run, detect the date has passed, disable the OOF, and then the task is removed.
$TaskExpiry      = $EndDate.AddDays(1)
# ─────────────────────────────────────────────────────────────────────────────

# Build the action -- runs pwsh and passes fixed params (logging handled by Start-Transcript in the script)
# Note: -File mode requires the colon-attached syntax (-Param:$Value) for [bool] parameters;
# a space-separated -Param $Value is treated as a string literal and fails type conversion.
$scriptArgs = "-NonInteractive -ExecutionPolicy Bypass -File `"$ScriptPath`" -Mailbox `"$Mailbox`" -StartDate `"$($StartDate.ToString('yyyy-MM-ddTHH:mm:ss'))`" -EndDate `"$($EndDate.ToString('yyyy-MM-ddTHH:mm:ss'))`" -DeclineMeetings:`$$DeclineMeetings"
$action = New-ScheduledTaskAction `
    -Execute   "pwsh.exe" `
    -Argument  $scriptArgs

# Trigger: daily at the specified RunTime, expiring 1 day after EndDate
$trigger = New-ScheduledTaskTrigger -Daily -At $RunTime
$trigger.EndBoundary = $TaskExpiry.ToString("yyyy-MM-ddTHH:mm:ss")

# Settings
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -DeleteExpiredTaskAfter (New-TimeSpan -Days 1) `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

# Run as SYSTEM so no user logon is required
$principal = New-ScheduledTaskPrincipal `
    -UserId  "SYSTEM" `
    -LogonType ServiceAccount `
    -RunLevel Highest

# Register (or overwrite) the task
Register-ScheduledTask `
    -TaskName    $TaskName `
    -Description $TaskDescription `
    -Action      $action `
    -Trigger     $trigger `
    -Settings    $settings `
    -Principal   $principal `
    -Force

Write-Output "Scheduled task '$TaskName' registered successfully."
Write-Output "  Mailbox         : $Mailbox"
Write-Output "  OOF             : $($StartDate.ToString('yyyy-MM-dd')) -> $($EndDate.ToString('yyyy-MM-dd'))"
Write-Output "  Runs daily at $RunTime as SYSTEM."
Write-Output "  DeclineMeetings : $DeclineMeetings"
Write-Output "  Task expires    : $($TaskExpiry.ToString('yyyy-MM-dd'))"
Write-Output "  Script          : $ScriptPath"
