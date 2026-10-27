# CLAUDE.md — exo-persistent-oof

PowerShell 7 toolkit that **re-applies Exchange Online out-of-office (OOF) auto-replies daily** through Windows Task Scheduler. Re-setting the Scheduled OOF window each day resets Exchange's "already replied" tracker, so senders get one reply **per day** instead of one for the whole period. After the window ends, the run disables the OOF and Task Scheduler deletes the expired task.

## Architecture

```
exo-persistent-oof.ps1  → interactive menu (register / list / remove / run / show OOF / exit)
  └─ register-ooftask.ps1  → creates task "Set-OutOfOffice-<usertag>" (runs as SYSTEM, daily)
       └─ set-outofoffice.ps1  → connects to EXO (cert auth), applies or disables OOF, Start-Transcript log
test-deployment.ps1     → run from the repo: checks the deployed copy, optional live run of one task
config.json             → AppId, Organization, CertificateThumbprint (secret, gitignored; never print values)
messages/<usertag>.txt  → per-mailbox message; fallback messages/default.txt, then a hardcoded string
logs/task_output-<usertag>.log → one appended transcript per mailbox, never rotated
```

Production runs from `C:\Scripts\tasks\`. **The menu calls the deployed `register-ooftask.ps1`, not the repo copy**, so a change is live only after it's copied there. Run `.\test-deployment.ps1` (from the repo) after every copy.

## Conventions and gotchas

- **User tag:** `($Mailbox.Split('@')[0]).ToLower() -replace '[\\/:*?"<>|]', '_'`. Used for task names, message files and log files, and must stay identical in every script.
- **Bool arguments via `pwsh -File`** need the colon syntax `-DeclineMeetings:$False`; with a space the value arrives as a string and fails. `register-ooftask.ps1` builds the task arguments this way.
- **The menu parses task arguments with regexes** (`-Mailbox\s+"([^"]+)"`, `-DeclineMeetings[:\s]+\$(\w+)`, dates). If the argument format in `register-ooftask.ps1` changes, update those regexes and the ones in `test-deployment.ps1`. A task without `-DeclineMeetings` means `$true`, the parameter's default.
- **`-DeclineMeetings`** defaults to `$true`. Shared mailboxes that must keep accepting meetings use `$false`.
- **Dates:** `yyyy-MM-dd` means 00:00, so EndDate is exclusive. The run on EndDate disables the OOF. The trigger expires at EndDate+1, and Task Scheduler deletes the task about one day later.
- **Paths are hardcoded** to `C:\Scripts\tasks` (`set-outofoffice.ps1` Logs/Messages, `register-ooftask.ps1` `$ScriptPath`, `exo-persistent-oof.ps1` `$ScriptRoot`). Only `config.json` is resolved through `$PSScriptRoot`.
- **The EXO module must be installed with `-Scope AllUsers`**, because SYSTEM can't see CurrentUser installs. `register-ooftask.ps1` refuses to register otherwise.
- **Don't add `#Requires -Modules ExchangeOnlineManagement`** to `set-outofoffice.ps1`. It fails before the transcript starts, so nothing is logged. Keep `Import-Module ... -ErrorAction Stop` inside the `try`.
- Errors before `Start-Transcript` (missing config.json, mailbox without `@`) exit 1 without writing a log line.
- **Scripts signal failure with `Write-Error` + `exit 1`, which `try/catch` does not see.** A caller that runs a script with `&` (like the menu calling `register-ooftask.ps1`) must reset `$global:LASTEXITCODE = 0` before the call and check it after. Give external cmdlets whose failure matters (`Connect-ExchangeOnline`, `Set-MailboxAutoReplyConfiguration`, `Register-ScheduledTask`) `-ErrorAction Stop`, or a non-terminating error is followed by a "success" message.
- Stagger `-RunTime` across mailboxes (00:30, 00:35, …) to avoid EXO throttling.

## Code style (repo conventions)

Match the existing scripts:
- Comment-based help with `.NAME`, `.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, `.EXAMPLE`; identical `.NOTES` (Author / License / Source) in every file.
- Typed **PascalCase** parameters (`$Mailbox`, `$StartDate`); kebab-case lowercase script names.
- Output with `Write-Output`; fatal errors with `Write-Error "..."; exit 1`. No PSFramework dependency.
- `Write-Host` only in the interactive menu (`exo-persistent-oof.ps1`).
- `# ── Section ──` banner comments; vertical alignment of named arguments when there are three or more.
- State-changing scripts support `-WhatIf` (`SupportsShouldProcess`, `ConfirmImpact = 'Low'` so the menu gets no prompts). Wrap only the mutating call in `ShouldProcess`.
- EXO: certificate auth only (cert in `Cert:\LocalMachine\My`), `Disconnect-ExchangeOnline -Confirm:$false` in `finally`.

## Testing and safety

There are no unit tests or CI yet. Verify with:
1. A parser check on every changed `.ps1` (`[System.Management.Automation.Language.Parser]::ParseFile`).
2. `register-ooftask.ps1 ... -WhatIf`, which runs all validation and registers nothing.
3. `.\test-deployment.ps1` (offline). Use `-LiveTask <usertag>` only for a live end-to-end run.

Rules:
- **A live run changes a real mailbox** in production Exchange Online. Ask before any `Start-ScheduledTask` or live `test-deployment.ps1` run. Never live-run a task with DeclineMeetings=True: it can decline real meeting requests.
- **Test anything that could register or start a task with `-WhatIf`, never with stub functions.** A function stub for `Register-ScheduledTask` once failed to intercept and created a real task.
- Copying to `C:\Scripts\tasks\` is a production deployment. Compare first; git history is the backup.

## Planned changes

1. **`$PSScriptRoot` migration:** replace the hardcoded `C:\Scripts\tasks` paths. `register-ooftask.ps1` must bake the *absolute* resolved script path into the task action, because Task Scheduler has no script-relative context.
2. **Shared EXO session in the menu:** connect lazily on first need through a helper that tracks `$script:ExoConnected`, reuse the session, and disconnect once on exit. Verify mailboxes with `Get-EXOMailbox` before registering. Only "not found" means absent; throttling or other errors are inconclusive and must stop registration. Multiple matches must throw.
