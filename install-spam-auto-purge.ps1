$taskName = "FrostyTagYT-AutoPurgeSpam"
$script = "C:\Users\luked\frostytagyt-press\block-junk-senders.ps1"

schtasks /Delete /TN $taskName /F 2>$null | Out-Null
schtasks /Create /TN $taskName /TR "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$script`"" /SC MINUTE /MO 5 /F | Out-Null

Write-Output "TASK_INSTALLED: $taskName (every 5 minutes via schtasks)"
Write-Output "SCRIPT: $script"