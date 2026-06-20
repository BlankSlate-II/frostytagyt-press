$Site = "https://blankslate-ii.github.io/frostytagyt-press/"

Write-Host ""
Write-Host "VR Scout contact page is DEAD (404)." -ForegroundColor Yellow
Write-Host "Using working replacements for links 3 and 4:" -ForegroundColor Cyan
Write-Host "  3) UploadVR contact form" -ForegroundColor White
Write-Host "  4) Another Axiom Creator Program (Gorilla Tag devs)" -ForegroundColor White
Write-Host ""

# Link 3 - UploadVR form
$uploadvr = @"
Subject: Story Tip: Gorilla Tag TikTok Creator and Mod Developer

Hi UploadVR team,

Story tip for FrostyTagYT (@frostytagyt) - Gorilla Tag VR creator on TikTok with 28.4K followers who also develops community mods and tools.

Press kit: ${Site}press-kit.html
TikTok: https://www.tiktok.com/@frostytagyt
Email: lukedonovanyoung2020@outlook.com

Available for interview.

FrostyTagYT
"@
Set-Clipboard -Value $uploadvr
Start-Process "https://www.uploadvr.com/contact/"
Write-Host "[3/4] UploadVR form opened - select SEND US A STORY, then Ctrl+V" -ForegroundColor Green
Start-Sleep -Seconds 3

# Link 4 - Another Axiom Creator Portal (TikToker application)
$aaNote = @"
Another Axiom Creator Program - apply as TikToker/Live Streamer

Your info for the application:
- Handle: @frostytagyt
- Followers: 28.4K
- Niche: Gorilla Tag VR, TikTok LIVE, mod development
- Press kit: ${Site}press-kit.html
- Email: lukedonovanyoung2020@outlook.com
- TikTok: https://www.tiktok.com/@frostytagyt
"@
Set-Clipboard -Value $aaNote
Start-Process "https://aa-creator-portal.pages.dev/"
Start-Process "https://www.anotheraxiom.com/aa-creator"
Write-Host "[4/4] Another Axiom Creator Portal opened - apply as TikToker" -ForegroundColor Green

Write-Host ""
Write-Host "Done. Submit UploadVR form + AA Creator application." -ForegroundColor Cyan
Write-Host "Note: AA Creator Program is not a news article - still need outlets to PUBLISH about you for TikTok." -ForegroundColor Yellow
pause