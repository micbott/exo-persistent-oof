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

    Before registering, the script verifies that ExchangeOnlineManagement is
    installed with -Scope AllUsers, since the task runs as SYSTEM and cannot
    see modules installed for the current user only.

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

.EXAMPLE
    .\register-ooftask.ps1 -Mailbox "user@contoso.com" -StartDate "2026-02-25" -EndDate "2026-03-27" -WhatIf
    Runs every validation (including the AllUsers module check) and shows the
    task that would be registered, without registering it.

.NOTES
    Author  : Michelangelo Bottura
    License : MIT
    Source  : https://github.com/micbott/exo-persistent-oof
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
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

# Pre-flight: the task runs as SYSTEM, which only sees AllUsers module paths
# under Program Files. A CurrentUser install would make every daily run fail.
# Check the folders directly rather than Get-Module -ListAvailable, so the
# result does not depend on which shell (powershell.exe / pwsh) runs this script.
$allUsersModulePaths = @(
    Join-Path $env:ProgramFiles 'PowerShell\Modules\ExchangeOnlineManagement'
    Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules\ExchangeOnlineManagement'
)
if (-not ($allUsersModulePaths | Where-Object { Test-Path $_ })) {
    Write-Error ("ExchangeOnlineManagement is not installed for AllUsers (checked: $($allUsersModulePaths -join ', ')). " +
        "The task runs as SYSTEM and cannot load a CurrentUser install. " +
        "From an elevated shell run: Install-Module ExchangeOnlineManagement -Scope AllUsers")
    exit 1
}

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

# Register (or overwrite) the task -- skipped under -WhatIf, all checks above still run
if ($PSCmdlet.ShouldProcess($TaskName, "Register scheduled task (runs as SYSTEM daily at $RunTime)")) {
    Register-ScheduledTask `
        -TaskName    $TaskName `
        -Description $TaskDescription `
        -Action      $action `
        -Trigger     $trigger `
        -Settings    $settings `
        -Principal   $principal `
        -Force `
        -ErrorAction Stop
    Write-Output "Scheduled task '$TaskName' registered successfully."
}
else {
    # ShouldProcess is also false when -Confirm is answered No
    $label = if ($WhatIfPreference) { "[WhatIf] Would register" } else { "Skipped registering" }
    Write-Output "$label '$TaskName' (nothing was changed)."
    Write-Output "  Command         : pwsh.exe $scriptArgs"
}
Write-Output "  Mailbox         : $Mailbox"
Write-Output "  OOF             : $($StartDate.ToString('yyyy-MM-dd')) -> $($EndDate.ToString('yyyy-MM-dd'))"
Write-Output "  Runs daily at $RunTime as SYSTEM."
Write-Output "  DeclineMeetings : $DeclineMeetings"
Write-Output "  Task expires    : $($TaskExpiry.ToString('yyyy-MM-dd'))"
Write-Output "  Script          : $ScriptPath"
