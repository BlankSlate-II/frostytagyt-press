$Site = "https://blankslate-ii.github.io/frostytagyt-press/"
$Email = "lukedonovanyoung2020@outlook.com"
$uploadvr = @"
Story tip for UploadVR editorial:

FrostyTagYT (@frostytagyt) is a Gorilla Tag VR creator on TikTok with 28.4K followers who also develops community mods and tools.

Press kit: ${Site}press-kit.html
TikTok: https://www.tiktok.com/@frostytagyt

FrostyTagYT
$Email
"@
Set-Clipboard -Value $uploadvr
Start-Process "https://www.uploadvr.com/contact/"
Start-Sleep -Seconds 2
Start-Process "https://aa-creator-portal.pages.dev/"
Start-Sleep -Seconds 1
Start-Process "https://www.helpareporter.com/sources/"
Write-Output "FORMS_OPENED"