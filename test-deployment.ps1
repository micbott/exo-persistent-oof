<#
.NAME
    exo-persistent-oof -- test-deployment.ps1

.SYNOPSIS
    Checks that the deployed OOF scripts match this repo and can run as SYSTEM,
    and optionally runs one scheduled task end-to-end.

.DESCRIPTION
    Run from the repo checkout. Offline checks (no Exchange Online contact):
      - pwsh.exe is installed where the scheduled task expects it
      - ExchangeOnlineManagement is installed for AllUsers (visible to SYSTEM)
      - each deployed .ps1 matches the repo copy (line endings ignored)
      - each deployed .ps1 parses
      - config.json, Messages\default.txt and Logs\ exist in the deploy folder
      - status of every Set-OutOfOffice-* scheduled task

    With -LiveTask, it also starts that task, waits for it to finish, and checks
    the exit code and this run's transcript lines. This applies the OOF settings
    in Exchange Online, the same as the task's nightly run. The live run is
    skipped when any offline check fails, and refused for tasks that decline
    meetings unless -AllowDeclineMeetings is given.

    Exits 0 when no check failed, 1 otherwise.

.PARAMETER DeployPath
    Folder the scheduled tasks run from (default C:\Scripts\tasks).

.PARAMETER LiveTask
    User tag of the task to run live (e.g. "ctao-media" for Set-OutOfOffice-ctao-media).

.PARAMETER AllowDeclineMeetings
    Allow a live run of a task with -DeclineMeetings:$True. Such a run can
    decline real meeting requests in that mailbox.

.PARAMETER TimeoutSeconds
    How long to wait for the live task to finish (default 120).

.EXAMPLE
    .\test-deployment.ps1
    Offline checks only.

.EXAMPLE
    .\test-deployment.ps1 -LiveTask ctao-media
    Offline checks, then a live run of Set-OutOfOffice-ctao-media.

.EXAMPLE
    .\test-deployment.ps1 -LiveTask ctao-media -WhatIf
    Offline checks, and shows the live run without starting the task.

.NOTES
    Author  : Michelangelo Bottura
    License : MIT
    Source  : https://github.com/micbott/exo-persistent-oof
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
param(
    [string] $DeployPath = "C:\Scripts\tasks",
    [string] $LiveTask,
    [switch] $AllowDeclineMeetings,
    [ValidateRange(10, 1800)]
    [int]    $TimeoutSeconds = 120
)

# ── Configuration ────────────────────────────────────────────────────────────
$ScriptFiles = @('set-outofoffice.ps1', 'register-ooftask.ps1', 'exo-persistent-oof.ps1')
$PwshPath    = Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'
# Same AllUsers paths as the pre-flight check in register-ooftask.ps1 -- keep in sync
$ExoModulePaths = @(
    Join-Path $env:ProgramFiles 'PowerShell\Modules\ExchangeOnlineManagement'
    Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules\ExchangeOnlineManagement'
)
# Success lines written by set-outofoffice.ps1
$SuccessPattern = 'OOF auto-reply successfully set|OOF disabled successfully'
# ─────────────────────────────────────────────────────────────────────────────

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
function Add-Result ([string]$Check, [string]$Status, [string]$Detail) {
    $results.Add([PSCustomObject]@{ Check = $Check; Status = $Status; Detail = $Detail })
}

# Same rule as Get-DeclineMeetingsFromTask in exo-persistent-oof.ps1: a task without the
# flag uses set-outofoffice.ps1's default ($true), so it DOES decline meetings.
function Get-DeclineMeetings ($Task) {
    $match = [regex]::Match($Task.Actions[0].Arguments, '-DeclineMeetings[:\s]+\$(\w+)')
    if ($match.Success) { return $match.Groups[1].Value }
    return 'True'
}

if (-not (Test-Path $DeployPath)) {
    Write-Error "Deploy folder not found: $DeployPath"
    exit 1
}

# 1. pwsh.exe -- the task action runs "pwsh.exe" as SYSTEM
if (Test-Path $PwshPath) { Add-Result 'pwsh.exe' 'Pass' $PwshPath }
else                     { Add-Result 'pwsh.exe' 'Fail' "Not found: $PwshPath" }

# 2. ExchangeOnlineManagement installed for AllUsers
$exoFound = $ExoModulePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($exoFound) { Add-Result 'EXO module (AllUsers)' 'Pass' $exoFound }
else           { Add-Result 'EXO module (AllUsers)' 'Fail' "Not in: $($ExoModulePaths -join ', '). Run elevated: Install-Module ExchangeOnlineManagement -Scope AllUsers" }

# 3 + 4. Deployed scripts match the repo and parse
foreach ($file in $ScriptFiles) {
    $repoFile     = Join-Path $PSScriptRoot $file
    $deployedFile = Join-Path $DeployPath $file

    if (-not (Test-Path $deployedFile)) {
        Add-Result "Match: $file" 'Fail' "Not deployed: $deployedFile"
        continue
    }

    # Compare with CRLF normalised so a line-ending-only difference doesn't fail
    $repoText     = (Get-Content -Path $repoFile -Raw) -replace "`r`n", "`n"
    $deployedText = (Get-Content -Path $deployedFile -Raw) -replace "`r`n", "`n"
    if ($repoText -ceq $deployedText) {
        Add-Result "Match: $file" 'Pass' 'Deployed copy matches repo'
    }
    else {
        $deployedTime = (Get-Item $deployedFile).LastWriteTime.ToString('yyyy-MM-dd HH:mm')
        Add-Result "Match: $file" 'Fail' "Deployed copy differs from repo (deployed file modified $deployedTime)"
    }

    $tokens = $null; $parseErrors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($deployedFile, [ref]$tokens, [ref]$parseErrors) | Out-Null
    if ($parseErrors) { Add-Result "Parse: $file" 'Fail' ($parseErrors | Select-Object -First 1).Message }
    else              { Add-Result "Parse: $file" 'Pass' 'No parse errors' }
}

# 5. Supporting files (existence only -- config.json holds secrets and is never read)
foreach ($item in 'config.json', 'Messages\default.txt', 'Logs') {
    $path = Join-Path $DeployPath $item
    if (Test-Path $path) { Add-Result "Exists: $item" 'Pass' $path }
    else                 { Add-Result "Exists: $item" 'Fail' "Missing: $path" }
}

# 6. Status of existing OOF tasks
$tasks = @(Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -like "Set-OutOfOffice-*" })
foreach ($t in $tasks) {
    $info    = $t | Get-ScheduledTaskInfo
    $decline = Get-DeclineMeetings $t
    $detail  = "State=$($t.State) LastResult=$($info.LastTaskResult) NextRun=$($info.NextRunTime) DeclineMeetings=$decline"
    # 267011 (0x41303) = task has not run yet
    $status  = if ($info.LastTaskResult -in 0, 267011) { 'Info' } else { 'Warn' }
    Add-Result "Task: $($t.TaskName)" $status $detail
}

# ── Live run ─────────────────────────────────────────────────────────────────
$newLog = $null
if ($LiveTask) {
    $taskName = "Set-OutOfOffice-$LiveTask"
    $task     = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    $decline  = if ($task) { Get-DeclineMeetings $task }
    $isAdmin  = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
                ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $task) {
        Add-Result "Live: $taskName" 'Fail' 'Task not found'
    }
    elseif ($decline -eq 'True' -and -not $AllowDeclineMeetings) {
        Add-Result "Live: $taskName" 'Fail' 'Refused -- task declines meetings; pass -AllowDeclineMeetings to run it anyway'
    }
    elseif ($results.Status -contains 'Fail') {
        Add-Result "Live: $taskName" 'Fail' 'Skipped -- fix the failed offline checks first'
    }
    elseif ($PSCmdlet.ShouldProcess($taskName, 'Start scheduled task (applies OOF in Exchange Online)')) {
        if (-not $isAdmin) { Add-Result 'Live: elevation' 'Warn' 'Not elevated -- starting a SYSTEM task may fail' }

        $mailbox  = [regex]::Match($task.Actions[0].Arguments, '-Mailbox\s+"([^"]+)"').Groups[1].Value
        $userTag  = ($mailbox.Split('@')[0]).ToLower() -replace '[\\/:*?"<>|]', '_'
        $logFile  = Join-Path $DeployPath "Logs\task_output-$userTag.log"
        $logStart = if (Test-Path $logFile) { (Get-Item $logFile).Length } else { 0 }

        $prevRunTime = (Get-ScheduledTaskInfo -TaskName $taskName).LastRunTime
        $startedAt   = Get-Date
        Start-ScheduledTask -TaskName $taskName -ErrorAction Stop
        do {
            Start-Sleep -Seconds 5
            $state = (Get-ScheduledTask -TaskName $taskName).State
        } while ($state -eq 'Running' -and ((Get-Date) - $startedAt).TotalSeconds -lt $TimeoutSeconds)

        $info = Get-ScheduledTaskInfo -TaskName $taskName
        if ($state -eq 'Running') {
            Add-Result "Live: $taskName" 'Fail' "Still running after $TimeoutSeconds s"
        }
        # LastRunTime drifts tens of seconds from the real start, so compare it with its previous value, not the clock
        elseif ($info.LastRunTime -le $prevRunTime) {
            Add-Result "Live: $taskName" 'Fail' "Task did not run (LastRunTime $($info.LastRunTime) is not after previous $prevRunTime)"
        }
        elseif ($info.LastTaskResult -ne 0) {
            Add-Result "Live: $taskName" 'Fail' "LastTaskResult=$($info.LastTaskResult)"
        }
        else {
            Add-Result "Live: $taskName" 'Pass' "LastTaskResult=0 in $([int]((Get-Date) - $startedAt).TotalSeconds) s"
        }

        # Read only this run's transcript; the task may still hold the file, so share ReadWrite
        if (Test-Path $logFile) {
            $stream = [System.IO.File]::Open($logFile, 'Open', 'Read', 'ReadWrite')
            try {
                $null   = $stream.Seek($logStart, 'Begin')
                $reader = [System.IO.StreamReader]::new($stream)
                $newLog = $reader.ReadToEnd()
            }
            finally { $stream.Dispose() }
        }
        if ($newLog -match $SuccessPattern) { Add-Result 'Live: log' 'Pass' $Matches[0] }
        else                                { Add-Result 'Live: log' 'Fail' "No success line in new output of $logFile" }
    }
}

# ── Report ───────────────────────────────────────────────────────────────────
Write-Output ($results | Format-Table -AutoSize -Wrap | Out-String)
if ($newLog) {
    Write-Output "── New transcript lines ──"
    Write-Output $newLog
}

$failed = @($results | Where-Object Status -eq 'Fail').Count
Write-Output "$failed check(s) failed."
exit ([int]($failed -gt 0))
