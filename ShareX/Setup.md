# ShareX Auto-Rename and Clipboard Rescue - Setup Steps

Script: `C:\temp\Invoke-ShareXRename_withSnip.ps1` (v3.3)
Accounts: a standard account for day-to-day use and a separate admin account for elevation.
The scheduled tasks run as the **standard** account. Setup is run from the **admin** session.

## 1. Put the script where the standard account can read it

- Save the script as `C:\temp\Invoke-ShareXRename_withSnip.ps1` (or move it to somewhere permanent such as `C:\Scripts\`).
- Do not use the admin profile or an admin-only folder.
- If the file came from a download or a browser copy, unblock it:

```powershell
Unblock-File C:\temp\Invoke-ShareXRename_withSnip.ps1
```

- The scheduled tasks store this exact path. If the script is moved or renamed later, re-run both setup commands in step 4.

## 2. Set the API key (standard account, normal non-elevated PowerShell)

```powershell
[Environment]::SetEnvironmentVariable('ANTHROPIC_API_KEY','<your key>','User')
```

Open a fresh PowerShell window afterwards.

Check the key without printing it:

```powershell
'Process','User','Machine' | ForEach-Object {
    $K = [Environment]::GetEnvironmentVariable('ANTHROPIC_API_KEY', $_)
    If ($K) { '{0}: length {1}, starts "{2}", ends "{3}"' -f $_, $K.Length, $K.Substring(0,7), $K.Substring($K.Length-4) } Else { "${_}: not set" }
}
```

- The script reads Process scope first, then User, then Machine.
- If Process and User show different values, the Process one wins. Clear it:

```powershell
Remove-Item Env:\ANTHROPIC_API_KEY
```

- If a PowerShell profile sets it, find it with `Select-String -Path $PROFILE -Pattern 'ANTHROPIC'` and remove that line.

Test the key directly (should list models, not return a 401):

```powershell
$K = [Environment]::GetEnvironmentVariable('ANTHROPIC_API_KEY','User').Trim()
Invoke-RestMethod -Uri 'https://api.anthropic.com/v1/models' -Headers @{ 'x-api-key' = $K; 'anthropic-version' = '2023-06-01' } | Select-Object -ExpandProperty data | Select-Object -First 3 id
```

## 3. Smoke test (standard account, non-elevated)

Rename test. This really renames files and really calls the API, so try it on copies first:

```powershell
powershell.exe -ExecutionPolicy Bypass -File C:\temp\Invoke-ShareXRename_withSnip.ps1
```

Clipboard watcher test:

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File C:\temp\Invoke-ShareXRename_withSnip.ps1 -WatchClipboard
```

- Take two snips back to back.
- Check for `Snip_*.png` files in `Documents\ShareX\Screenshots\yyyy-MM`.
- Press Ctrl+C to stop.

## 4. Register the scheduled tasks (admin account, elevated PowerShell)

Keep the standard account logged on at the desktop while doing this. The script works out the task account from the logged-on console user.

Clipboard watcher (hidden, starts at every logon and starts now):

```powershell
powershell.exe -ExecutionPolicy Bypass -File C:\temp\Invoke-ShareXRename_withSnip.ps1 -SetupWatcherTask
```

Scheduled rename. The day list must go through `-Command`, because `-File` passes it as one literal string and validation fails:

```powershell
powershell.exe -ExecutionPolicy Bypass -Command "& 'C:\temp\Invoke-ShareXRename_withSnip.ps1' -SetupScheduledTask -ScheduleDays Monday,Tuesday,Wednesday,Thursday,Friday -ScheduleStartHour 8 -ScheduleTimesPerDay 3"
```

- Adjust the days, start hour and runs per day as needed. Runs are spread evenly from the start hour towards midnight.
- Look for the log line `Scheduled tasks will run as: DATA3\<standard user>`. If it names any other account, stop and re-run with `-TaskUser 'DATA3\<standarduser>'`.
- Only pass `-ScreenshotFolder` if the ShareX folder is not the default. Passing it bakes the path into the task.
- Re-running a setup command overwrites the existing task.

## 5. Verify

```powershell
Get-ScheduledTask -TaskPath '\ShareXRename\' | Select-Object TaskName, State
Get-ScheduledTaskInfo -TaskPath '\ShareXRename\' -TaskName 'ShareXRename'
Start-ScheduledTask -TaskPath '\ShareXRename\' -TaskName 'ShareXRename'
```

Then check the screenshot root folder for:

- `ShareXRename_<timestamp>.log` - every rename action (CMTrace format)
- `ShareXRename_Uncertain_<timestamp>.log` - files the AI was unsure about (not renamed)
- `ShareXClipboardWatcher_<yyyyMMdd>.log` - every clipboard image saved

If a task shows Ready but never writes a log, check file permissions on the script, then whether Group Policy overrides the execution policy (unconfirmed possibility).

## 6. Day to day

Retry uncertain files (standard account, non-elevated):

```powershell
.\Invoke-ShareXRename_withSnip.ps1 -RetryUncertain
```

Retry a single uncertain log only:

```powershell
.\Invoke-ShareXRename_withSnip.ps1 -RetryUncertain -UncertainLogPath '<path to log>'
```

- To change the schedule, re-run the `-SetupScheduledTask` command from the admin session.
- Normal runs skip files already listed in an uncertain log, so they are not re-sent every time.

## 7. Remove everything (admin, elevated)

```powershell
Unregister-ScheduledTask -TaskPath '\ShareXRename\' -TaskName 'ShareXRename' -Confirm:$false
Unregister-ScheduledTask -TaskPath '\ShareXRename\' -TaskName 'ShareXClipboardWatcher' -Confirm:$false
```

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `HTTP 401` on every file | API key rejected, often a stale Process-scope value | Step 2 check, then `Remove-Item Env:\ANTHROPIC_API_KEY` |
| `ScheduleDays ... does not belong to the set` | Used `-File` with a comma list | Use the `-Command` form in step 4 |
| `No mapping between account names and security IDs` | Task account resolved to SYSTEM | Use v3.1 or later, or pass `-TaskUser` |
| `1/01/1601` dates in CMTrace | Old timezone-bias bug | Use v3.3 |
| Task Ready but no logs | Script path unreadable, or policy override | Check permissions on the script, then Group Policy |
