$taskName = "FrostyTagYT-DailyMediaCheck"
$script = "C:\Users\luked\frostytagyt-press\check-media-replies.ps1"

schtasks /Delete /TN $taskName /F 2>$null | Out-Null
schtasks /Create /TN $taskName /TR "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$script`"" /SC DAILY /ST 09:00 /F | Out-Null

Write-Output "INSTALLED: $taskName (every day at 9:00 AM)"
Write-Output "Also runs: block junk spam every 5 min - media emails are protected"