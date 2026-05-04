<#
.NAME
    exo-persistent-oof -- exo-persistent-oof.ps1

.SYNOPSIS
    Interactive controller menu for managing OOF scheduled tasks.

.DESCRIPTION
    Presents a looping menu with the following options:

      1) Register  -- Prompt for mailbox, start date, duration, and daily run
                      time; compute the end date; preview the command;
                      optionally execute register-ooftask.ps1 immediately.
                      The run time defaults to 5 minutes after the latest
                      existing OOF task trigger (or 00:30 if none exist).
      2) List      -- Show every OOF scheduled task (Set-OutOfOffice*) with
                      name, daily run time, next run, last run, last result,
                      state, and the OOF message that would be applied.
      3) Remove    -- Numbered multi-select removal with confirmation.
      4) Run & Log -- Trigger a task immediately and display its transcript log.
      5) Show OOF  -- Connect to Exchange Online and display the live OOF
                      configuration for a mailbox with an active task.
      6) Exit      -- Leave the menu.

    Input validation, edge-case handling (no tasks, missing scripts, bad input),
    and confirmation prompts are built in.
    Must be run elevated (Run as Administrator).

.EXAMPLE
    pwsh -ExecutionPolicy Bypass -File "C:\Scripts\tasks\exo-persistent-oof.ps1"

.NOTES
    Author  : Michelangelo Bottura
    License : MIT
    Source  : https://github.com/micbott/exo-persistent-oof

    All scripts are expected under C:\Scripts\tasks\.
    Message files live in C:\Scripts\tasks\Messages\ (per-user and default).
    No Unicode special characters -- ASCII only.
#>

# ── Paths ────────────────────────────────────────────────────────────────────
$ScriptRoot     = "C:\Scripts\tasks"
$RegisterScript = Join-Path $ScriptRoot "Register-OofTask.ps1"
$MessageDir     = Join-Path $ScriptRoot "Messages"

# ── Exchange Online auth (lazy-loaded from Config.json on first need) ────────
$script:ExoAppId      = $null
$script:ExoOrg        = $null
$script:ExoThumbprint = $null
$script:ExoConfigLoaded = $false

function Ensure-ExoConfig {
    if ($script:ExoConfigLoaded) { return $true }
    $cfgPath = Join-Path $ScriptRoot "Config.json"
    if (-not (Test-Path $cfgPath)) {
        Write-Host "  ERROR: Config file not found: $cfgPath" -ForegroundColor Red
        return $false
    }
    $exoConfig = Get-Content -Path $cfgPath -Raw | ConvertFrom-Json
    $script:ExoAppId      = $exoConfig.AppId
    $script:ExoOrg        = $exoConfig.Organization
    $script:ExoThumbprint = $exoConfig.CertificateThumbprint
    $script:ExoConfigLoaded = $true
    return $true
}
# ─────────────────────────────────────────────────────────────────────────────

# ── Helper: derive user tag from mailbox (same logic as all other scripts) ───
function Get-UserTag ([string]$Mailbox) {
    ($Mailbox.Split('@')[0]).ToLower() -replace '[\\/:*?"<>|]', '_'
}

# ── Helper: resolve the OOF message that would be applied for a user tag ─────
function Get-OofMessage ([string]$UserTag) {
    $userFile    = Join-Path $MessageDir "$UserTag.txt"
    $defaultFile = Join-Path $MessageDir "default.txt"

    if (Test-Path $userFile) {
        return (Get-Content -Path $userFile -Raw).TrimEnd()
    }
    elseif (Test-Path $defaultFile) {
        return (Get-Content -Path $defaultFile -Raw).TrimEnd()
    }
    else {
        return "(no message file found -- hardcoded fallback would be used)"
    }
}

# ── Helper: extract the mailbox from the task action arguments ───────────────
function Get-MailboxFromTask ([Microsoft.Management.Infrastructure.CimInstance]$Task) {
    $actionArgs = $Task.Actions | Select-Object -First 1 -ExpandProperty Arguments
    if ($actionArgs -match '-Mailbox\s+"([^"]+)"') {
        return $Matches[1]
    }
    return $null
}

# ── Helper: parse date params from the task action arguments ─────────────────
function Get-DatesFromTask ([Microsoft.Management.Infrastructure.CimInstance]$Task) {
    $actionArgs = $Task.Actions | Select-Object -First 1 -ExpandProperty Arguments
    $start = $null; $end = $null
    if ($actionArgs -match '-StartDate\s+"([^"]+)"') { $start = $Matches[1] }
    if ($actionArgs -match '-EndDate\s+"([^"]+)"')   { $end   = $Matches[1] }
    return @{ StartDate = $start; EndDate = $end }
}

# ── Helper: extract the DeclineMeetings flag from the task action arguments ──
function Get-DeclineMeetingsFromTask ([Microsoft.Management.Infrastructure.CimInstance]$Task) {
    $actionArgs = $Task.Actions | Select-Object -First 1 -ExpandProperty Arguments
    # Matches both colon syntax (-DeclineMeetings:$False) and space syntax (-DeclineMeetings $False)
    if ($actionArgs -match '-DeclineMeetings[:\s]+\$(\w+)') {
        return $Matches[1] -eq 'True'
    }
    # Legacy tasks without the flag default to $true (original behaviour)
    return $true
}

# ── Helper: get the daily trigger time (HH:mm) from a task ──────────────────
function Get-TriggerTime ([Microsoft.Management.Infrastructure.CimInstance]$Task) {
    $trig = $Task.Triggers | Select-Object -First 1
    if ($trig -and $trig.StartBoundary) {
        try {
            $dt = [datetime]::Parse($trig.StartBoundary)
            return $dt.ToString("HH:mm")
        }
        catch { return $null }
    }
    return $null
}

# ── Helper: compute the next available run time (last trigger + 5 min) ───────
function Get-NextAvailableTime {
    $tasks = @(Get-ScheduledTask | Where-Object { $_.TaskName -like "Set-OutOfOffice*" })
    if ($tasks.Count -eq 0) { return "00:30" }

    $times = @()
    foreach ($t in $tasks) {
        $tt = Get-TriggerTime $t
        if ($tt) {
            try { $times += [datetime]::ParseExact($tt, "HH:mm", $null) }
            catch {}
        }
    }

    if ($times.Count -eq 0) { return "00:30" }

    $latest = ($times | Sort-Object)[-1]
    $next   = $latest.AddMinutes(5)
    return $next.ToString("HH:mm")
}

# ── Helper: wait for a keypress before returning to the menu ─────────────────
function Pause-ForKey {
    Write-Host ""
    Write-Host "  Press any key to return to the menu..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ── 1) Register ──────────────────────────────────────────────────────────────
function Invoke-Register {
    Write-Host ""

    # Check that the registration script exists
    if (-not (Test-Path $RegisterScript)) {
        Write-Host "ERROR: Registration script not found at:" -ForegroundColor Red
        Write-Host "  $RegisterScript" -ForegroundColor Red
        Pause-ForKey
        return
    }

    # --- Mailbox ---
    $mailbox = (Read-Host "  Mailbox address (e.g. user@contoso.com)").Trim()
    if ([string]::IsNullOrWhiteSpace($mailbox) -or $mailbox -notmatch '@') {
        Write-Host "  Invalid mailbox address." -ForegroundColor Red
        Pause-ForKey
        return
    }

    # --- Start date ---
    $startRaw = (Read-Host "  Start date (yyyy-MM-dd) [default: today]").Trim()
    if ([string]::IsNullOrWhiteSpace($startRaw)) {
        $startDate = (Get-Date).Date
    }
    else {
        try { $startDate = [datetime]::ParseExact($startRaw, "yyyy-MM-dd", $null) }
        catch {
            Write-Host "  Invalid date format. Use yyyy-MM-dd." -ForegroundColor Red
            Pause-ForKey
            return
        }
    }

    # --- Duration ---
    $durationRaw = (Read-Host "  Duration in days [default: 30]").Trim()
    if ([string]::IsNullOrWhiteSpace($durationRaw)) {
        $durationDays = 30
    }
    else {
        $durationDays = $durationRaw -as [int]
        if (-not $durationDays -or $durationDays -le 0) {
            Write-Host "  Duration must be a positive integer." -ForegroundColor Red
            Pause-ForKey
            return
        }
    }

    # --- Run time ---
    $suggestedTime = Get-NextAvailableTime
    $timeRaw = (Read-Host "  Daily run time HH:mm [default: $suggestedTime]").Trim()
    if ([string]::IsNullOrWhiteSpace($timeRaw)) {
        $runTime = $suggestedTime
    }
    else {
        # Validate HH:mm format
        try {
            $null = [datetime]::ParseExact($timeRaw, "HH:mm", $null)
            $runTime = $timeRaw
        }
        catch {
            Write-Host "  Invalid time format. Use HH:mm (e.g. 01:30)." -ForegroundColor Red
            Pause-ForKey
            return
        }
    }

    # --- Decline meetings ---
    $declineRaw = (Read-Host "  Auto-decline future meetings while OOF? (Y/N) [default: Y]").Trim()
    if ($declineRaw -eq 'N' -or $declineRaw -eq 'n') {
        $declineMeetings = $false
    }
    else {
        $declineMeetings = $true
    }

    # --- Compute end date ---
    $endDate = $startDate.AddDays($durationDays)

    # --- Preview ---
    $userTag  = Get-UserTag $mailbox
    $taskName = "Set-OutOfOffice-$userTag"
    $message  = Get-OofMessage $userTag

    Write-Host ""
    Write-Host "  ── Preview ─────────────────────────────────────────" -ForegroundColor Cyan
    Write-Host "  Task name       : $taskName"
    Write-Host "  Mailbox         : $mailbox"
    Write-Host "  OOF window      : $($startDate.ToString('yyyy-MM-dd')) -> $($endDate.ToString('yyyy-MM-dd'))"
    Write-Host "  Duration        : $durationDays day(s)"
    Write-Host "  Daily run       : $runTime"
    Write-Host "  DeclineMeetings : $declineMeetings"
    Write-Host "  Task expiry     : $($endDate.AddDays(1).ToString('yyyy-MM-dd'))  (EndDate + 1 day)"
    $msgLines = ($message -split '\r?\n') | Where-Object { $_.Trim() -ne '' } | Select-Object -First 3
    Write-Host "  Message    :" -ForegroundColor DarkGray
    foreach ($ln in $msgLines) { Write-Host "    $ln" -ForegroundColor DarkGray }
    Write-Host ""
    Write-Host "  Command:" -ForegroundColor DarkGray
    Write-Host ('   & "{0}" -Mailbox "{1}" -StartDate "{2}" -EndDate "{3}" -RunTime "{4}" -DeclineMeetings ${5}' -f $RegisterScript, $mailbox, $startDate.ToString('yyyy-MM-dd'), $endDate.ToString('yyyy-MM-dd'), $runTime, $declineMeetings) -ForegroundColor DarkGray
    Write-Host ""

    # Check if task already exists
    $existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  WARNING: A task named '$taskName' already exists and will be overwritten." -ForegroundColor Yellow
    }

    $confirm = (Read-Host "  Register this task now? (Y/N)").Trim()
    if ($confirm -ne 'Y' -and $confirm -ne 'y') {
        Write-Host "  Cancelled." -ForegroundColor Yellow
        Pause-ForKey
        return
    }

    Write-Host ""
    try {
        & $RegisterScript -Mailbox $mailbox -StartDate $startDate.ToString('yyyy-MM-dd') -EndDate $endDate.ToString('yyyy-MM-dd') -RunTime $runTime -DeclineMeetings $declineMeetings
        Write-Host ""
        Write-Host "  Task registered successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "  ERROR: $_" -ForegroundColor Red
    }
    Pause-ForKey
}

# ── 2) List ──────────────────────────────────────────────────────────────────
function Invoke-List {
    Write-Host ""
    $tasks = @(Get-ScheduledTask | Where-Object { $_.TaskName -like "Set-OutOfOffice*" })

    if ($tasks.Count -eq 0) {
        Write-Host "  No OOF scheduled tasks found." -ForegroundColor Yellow
        Pause-ForKey
        return
    }

    Write-Host "  Found $($tasks.Count) OOF task(s):" -ForegroundColor Cyan
    Write-Host ""

    foreach ($task in $tasks) {
        $info    = $task | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue
        $mailbox = Get-MailboxFromTask $task
        $dates   = Get-DatesFromTask $task
        $trigTime = Get-TriggerTime $task
        $userTag = if ($mailbox) { Get-UserTag $mailbox } else { $null }
        $decline = Get-DeclineMeetingsFromTask $task

        # Format last-run result as hex (0x0 = success)
        $lastResult = if ($info -and $null -ne $info.LastTaskResult) {
            "0x{0:X}" -f $info.LastTaskResult
        } else { "N/A" }

        Write-Host "  ── $($task.TaskName) ──────────────────────────────────" -ForegroundColor Cyan
        Write-Host "    State           : $($task.State)"
        Write-Host "    Mailbox         : $(if ($mailbox) { $mailbox } else { '(unknown)' })"
        Write-Host "    OOF window      : $(if ($dates.StartDate) { $dates.StartDate } else { 'N/A' }) -> $(if ($dates.EndDate) { $dates.EndDate } else { 'N/A' })"
        Write-Host "    Daily run       : $(if ($trigTime) { $trigTime } else { 'N/A' })"
        Write-Host "    DeclineMeetings : $decline"
        Write-Host "    Next run        : $(if ($info -and $info.NextRunTime -and $info.NextRunTime -ne [datetime]::MinValue) { $info.NextRunTime.ToString('yyyy-MM-dd HH:mm') } else { 'N/A' })"
        Write-Host "    Last run        : $(if ($info -and $info.LastRunTime -and $info.LastRunTime.Year -gt 1999) { $info.LastRunTime.ToString('yyyy-MM-dd HH:mm') } else { 'Never' })"
        Write-Host "    Last result     : $lastResult"

        if ($userTag) {
            $msg = Get-OofMessage $userTag
            $msgLines = ($msg -split '\r?\n') | Where-Object { $_.Trim() -ne '' } | Select-Object -First 3
            Write-Host "    Message    :" -ForegroundColor DarkGray
            foreach ($ln in $msgLines) { Write-Host "      $ln" -ForegroundColor DarkGray }
        }

        Write-Host ""
    }
    Pause-ForKey
}

# ── 3) Remove ────────────────────────────────────────────────────────────────
function Invoke-Remove {
    Write-Host ""
    $tasks = @(Get-ScheduledTask | Where-Object { $_.TaskName -like "Set-OutOfOffice*" })

    if ($tasks.Count -eq 0) {
        Write-Host "  No OOF scheduled tasks found." -ForegroundColor Yellow
        Pause-ForKey
        return
    }

    # Display numbered list
    Write-Host "  Existing OOF scheduled tasks:" -ForegroundColor Cyan
    Write-Host ""
    for ($i = 0; $i -lt $tasks.Count; $i++) {
        $t = $tasks[$i]
        Write-Host "    [$($i + 1)] $($t.TaskName)  (State: $($t.State))" -ForegroundColor White
    }
    Write-Host ""

    # Prompt for selection
    Write-Host "  Enter the numbers to delete (comma-separated), 'A' for all, or 'Q' to cancel:" -ForegroundColor Cyan
    $selection = (Read-Host "  Selection").Trim()

    if ($selection -eq 'Q' -or $selection -eq 'q') {
        Write-Host "  Cancelled." -ForegroundColor Yellow
        Pause-ForKey
        return
    }

    # Resolve selection
    if ($selection -eq 'A' -or $selection -eq 'a') {
        $selected = $tasks
    }
    else {
        $indices = @($selection -split ',' | ForEach-Object {
            $num = $_.Trim() -as [int]
            if ($num -and $num -ge 1 -and $num -le $tasks.Count) { $num - 1 }
            else { Write-Warning "  Invalid selection: $_" }
        } | Where-Object { $null -ne $_ })
        $selected = @($indices | ForEach-Object { $tasks[$_] })
    }

    if ($selected.Count -eq 0) {
        Write-Host "  No valid tasks selected." -ForegroundColor Yellow
        Pause-ForKey
        return
    }

    # Confirm and delete
    Write-Host ""
    Write-Host "  The following tasks will be deleted:" -ForegroundColor Red
    Write-Host ""
    $selected | ForEach-Object { Write-Host "    - $($_.TaskName)" -ForegroundColor White }
    Write-Host ""

    $confirm = (Read-Host "  Are you sure? (Y/N)").Trim()
    if ($confirm -ne 'Y' -and $confirm -ne 'y') {
        Write-Host "  Cancelled." -ForegroundColor Yellow
        Pause-ForKey
        return
    }

    foreach ($task in $selected) {
        try {
            Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false
            Write-Host "    Deleted: $($task.TaskName)" -ForegroundColor Green
        }
        catch {
            Write-Host "    ERROR deleting $($task.TaskName): $_" -ForegroundColor Red
        }
    }

    Write-Host ""
    Write-Host "  Done. $($selected.Count) task(s) removed." -ForegroundColor Cyan
    Pause-ForKey
}

# ── 4) Run & Log ─────────────────────────────────────────────────────────────
function Invoke-RunAndLog {
    Write-Host ""
    $tasks = @(Get-ScheduledTask | Where-Object { $_.TaskName -like "Set-OutOfOffice*" })

    if ($tasks.Count -eq 0) {
        Write-Host "  No OOF scheduled tasks found." -ForegroundColor Yellow
        Pause-ForKey
        return
    }

    # Display numbered list
    Write-Host "  Select a task to run:" -ForegroundColor Cyan
    Write-Host ""
    for ($i = 0; $i -lt $tasks.Count; $i++) {
        $t = $tasks[$i]
        Write-Host "    [$($i + 1)] $($t.TaskName)  (State: $($t.State))" -ForegroundColor White
    }
    Write-Host ""

    $selRaw = (Read-Host "  Task number (or 'Q' to cancel)").Trim()
    if ($selRaw -eq 'Q' -or $selRaw -eq 'q') {
        Write-Host "  Cancelled." -ForegroundColor Yellow
        Pause-ForKey
        return
    }

    $selNum = $selRaw -as [int]
    if (-not $selNum -or $selNum -lt 1 -or $selNum -gt $tasks.Count) {
        Write-Host "  Invalid selection." -ForegroundColor Red
        Pause-ForKey
        return
    }

    $task = $tasks[$selNum - 1]

    # Derive user tag from the task's mailbox argument to find the log file
    $mailbox = Get-MailboxFromTask $task
    if (-not $mailbox) {
        Write-Host "  Could not determine mailbox from task." -ForegroundColor Red
        Pause-ForKey
        return
    }
    $userTag = Get-UserTag $mailbox
    $logFile = Join-Path $ScriptRoot "Logs\task_output-$userTag.log"

    # Start the task
    Write-Host ""
    Write-Host "  Starting task '$($task.TaskName)'..." -ForegroundColor Cyan
    try {
        Start-ScheduledTask -TaskName $task.TaskName
    }
    catch {
        Write-Host "  ERROR starting task: $_" -ForegroundColor Red
        Pause-ForKey
        return
    }

    # Wait for the task to finish (poll every 2 seconds, up to 60 seconds)
    Write-Host "  Waiting for task to complete..." -ForegroundColor DarkGray
    $timeout = 60
    $elapsed = 0
    while ($elapsed -lt $timeout) {
        Start-Sleep -Seconds 2
        $elapsed += 2
        $state = (Get-ScheduledTask -TaskName $task.TaskName -ErrorAction SilentlyContinue).State
        if ($state -ne 'Running') { break }
        Write-Host "    ... still running ($elapsed s)" -ForegroundColor DarkGray
    }

    if ($state -eq 'Running') {
        Write-Host "  Task is still running after $timeout seconds. Check the log manually." -ForegroundColor Yellow
    }
    else {
        $info = Get-ScheduledTaskInfo -TaskName $task.TaskName -ErrorAction SilentlyContinue
        $result = if ($info) { "0x{0:X}" -f $info.LastTaskResult } else { "N/A" }
        Write-Host "  Task finished. Last result: $result" -ForegroundColor Cyan
    }

    # Show the transcript log
    Write-Host ""
    if (Test-Path $logFile) {
        Write-Host "  ── Log: $logFile ──" -ForegroundColor Cyan
        Write-Host ""
        # Show the last 30 lines of the log
        Get-Content -Path $logFile -Tail 30 | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
    }
    else {
        Write-Host "  Log file not found: $logFile" -ForegroundColor Yellow
    }

    Pause-ForKey
}

# ── 5) Show live OOF config ──────────────────────────────────────────────────
function Invoke-ShowOofConfig {
    Write-Host ""
    $tasks = @(Get-ScheduledTask | Where-Object { $_.TaskName -like "Set-OutOfOffice*" })

    if ($tasks.Count -eq 0) {
        Write-Host "  No OOF scheduled tasks found." -ForegroundColor Yellow
        Pause-ForKey
        return
    }

    # Display numbered list
    Write-Host "  Select a mailbox to query:" -ForegroundColor Cyan
    Write-Host ""
    $mailboxes = @()
    for ($i = 0; $i -lt $tasks.Count; $i++) {
        $mb = Get-MailboxFromTask $tasks[$i]
        $mailboxes += $mb
        $display = if ($mb) { $mb } else { "(unknown)" }
        Write-Host "    [$($i + 1)] $display  ($($tasks[$i].TaskName))" -ForegroundColor White
    }
    Write-Host ""

    $selRaw = (Read-Host "  Mailbox number (or 'Q' to cancel)").Trim()
    if ($selRaw -eq 'Q' -or $selRaw -eq 'q') {
        Write-Host "  Cancelled." -ForegroundColor Yellow
        Pause-ForKey
        return
    }

    $selNum = $selRaw -as [int]
    if (-not $selNum -or $selNum -lt 1 -or $selNum -gt $tasks.Count) {
        Write-Host "  Invalid selection." -ForegroundColor Red
        Pause-ForKey
        return
    }

    $mailbox = $mailboxes[$selNum - 1]
    if (-not $mailbox) {
        Write-Host "  Could not determine mailbox from task." -ForegroundColor Red
        Pause-ForKey
        return
    }

    # Load EXO credentials (lazy)
    if (-not (Ensure-ExoConfig)) {
        Pause-ForKey
        return
    }

    # Connect to Exchange Online
    Write-Host ""
    Write-Host "  Connecting to Exchange Online..." -ForegroundColor DarkGray
    try {
        Connect-ExchangeOnline `
            -AppId            $script:ExoAppId `
            -CertificateThumbprint $script:ExoThumbprint `
            -Organization     $script:ExoOrg `
            -ShowBanner:$false

        $oofConfig = Get-MailboxAutoReplyConfiguration -Identity $mailbox

        Write-Host ""
        Write-Host "  ── Live OOF config for $mailbox ──" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "    AutoReplyState : $($oofConfig.AutoReplyState)"
        Write-Host "    StartTime      : $($oofConfig.StartTime)"
        Write-Host "    EndTime        : $($oofConfig.EndTime)"
        Write-Host "    ExternalAudience : $($oofConfig.ExternalAudience)"
        Write-Host "    AutoDeclineFutureRequestsWhenOOF : $($oofConfig.AutoDeclineFutureRequestsWhenOOF)"
        Write-Host ""
        Write-Host "    Internal message:" -ForegroundColor DarkGray
        $intLines = ($oofConfig.InternalMessage -replace '<[^>]+>', '' -split '\r?\n') | Where-Object { $_.Trim() -ne '' } | Select-Object -First 5
        foreach ($ln in $intLines) { Write-Host "      $($ln.Trim())" -ForegroundColor DarkGray }
        Write-Host ""
        Write-Host "    External message:" -ForegroundColor DarkGray
        $extLines = ($oofConfig.ExternalMessage -replace '<[^>]+>', '' -split '\r?\n') | Where-Object { $_.Trim() -ne '' } | Select-Object -First 5
        foreach ($ln in $extLines) { Write-Host "      $($ln.Trim())" -ForegroundColor DarkGray }
    }
    catch {
        Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }

    Pause-ForKey
}

# ── Main menu loop ───────────────────────────────────────────────────────────
function Show-Menu {
    Write-Host ""
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host "       OOF Scheduled-Task Manager" -ForegroundColor Cyan
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    1) Register a new OOF task"
    Write-Host "    2) List existing OOF tasks"
    Write-Host "    3) Remove OOF tasks"
    Write-Host "    4) Run task & show log"
    Write-Host "    5) Show live OOF config"
    Write-Host "    6) Exit"
    Write-Host ""
}

# ── Entry point ──────────────────────────────────────────────────────────────
# Check for elevation
$isAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

while ($true) {
    Clear-Host
    if (-not $isAdmin) {
        Write-Host ""
        Write-Host "  WARNING: Not running as Administrator. Register/Remove may fail." -ForegroundColor Yellow
    }
    Show-Menu
    $choice = (Read-Host "  Choose an option (1-6)").Trim()

    switch ($choice) {
        '1' { Invoke-Register }
        '2' { Invoke-List }
        '3' { Invoke-Remove }
        '4' { Invoke-RunAndLog }
        '5' { Invoke-ShowOofConfig }
        '6' {
            Write-Host ""
            Write-Host "  Goodbye." -ForegroundColor Cyan
            Write-Host ""
            break
        }
        default {
            Write-Host "  Invalid option. Please enter 1-6." -ForegroundColor Red
        }
    }

    # Exit the while loop when option 6 is chosen
    if ($choice -eq '6') { break }
}
