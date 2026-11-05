
Screenshot Auto-Rename 
Runs a scheuled tasks monitoring a specific folder for any new scrreenshots uses claude (can be changed to any other LLM) to investigate the image and re-name the screen shoot to something apprioprate.  


**To get going**

Save the script as Invoke-ShareXRename.ps1. It’s pure ASCII, so the encoding makes no difference.
Try the watcher in the foreground first: powershell.exe -STA -File .\Invoke-ShareXRename.ps1 -WatchClipboard. Snip twice without saving, then check the yyyy-MM folder for two PNGs.
In an elevated prompt, run .\Invoke-ShareXRename.ps1 -SetupWatcherTask to make it permanent.
Check ANTHROPIC_API_KEY is set at User scope, then re-run your -SetupScheduledTask command to replace the old task.
