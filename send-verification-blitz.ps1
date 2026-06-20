# FrostyTagYT — send EVERYTHING still available for TikTok verification
# Skips Road to VR + UploadVR email (already pitched — do not duplicate)

$Site = "https://blankslate-ii.github.io/frostytagyt-press/"
$Email = "lukedonovanyoung2020@outlook.com"

Write-Host ""
Write-Host "=== FROSTYTAGYT VERIFICATION BLITZ ===" -ForegroundColor Cyan
Write-Host "Goal: get REAL published articles that mention @frostytagyt" -ForegroundColor White
Write-Host ""

$vrNewsBody = @"
Hi Virtual Reality News / Next Reality editorial team,

I'm pitching a creator profile that fits your Gorilla Tag coverage — including your recent piece on Meta betting VR's future on teens who grew up in Gorilla Tag.

FrostyTagYT (@frostytagyt on TikTok) is a Gorilla Tag VR creator with 28.4K followers who also develops and releases community mods and tools (wrist watch mods, lobby utilities). Rare combo: entertainer + builder in the VR-native creator generation your team already covers.

Possible angles:
- Profile: a mid-tier TikTok creator building a career through Gorilla Tag LIVE streaming in 2026
- How mod developers are shaping Gorilla Tag culture beyond clips and memes
- The VR-native teen creator wave — one creator's story inside Meta's bet

Press kit: ${Site}press-kit.html
Creator profile: ${Site}creator-profile.html
TikTok: https://www.tiktok.com/@frostytagyt

Available for interview, quotes, screenshots, or mod demos.

Thanks for considering,
FrostyTagYT
$Email
"@

$uploadvrForm = @"
Story tip for UploadVR editorial:

FrostyTagYT (@frostytagyt) is a Gorilla Tag VR creator on TikTok with 28.4K followers who also develops community mods and tools — wrist watch mods, lobby utilities, and other player-facing releases.

Angles: TikTok LIVE streaming in VR gaming, Gorilla Tag's creator economy, mod development by content creators, the VR-native teen creator wave.

Press kit: ${Site}press-kit.html
Creator profile: ${Site}creator-profile.html
TikTok: https://www.tiktok.com/@frostytagyt

Happy to provide interview, demos, or additional info.

FrostyTagYT
$Email
"@

$mixedNews = @"
News tip (English):

FrostyTagYT (@frostytagyt) — Gorilla Tag VR TikTok creator, 28.4K followers, mod developer (community tools + wrist watch mods). Story about TikTok LIVE streaming and grassroots modding in Gorilla Tag / Quest VR.

Press kit: ${Site}press-kit.html
TikTok: https://www.tiktok.com/@frostytagyt
Email: $Email
"@

# --- Step 1: Email Virtual Reality News / Next Reality ---
Write-Host "[1/5] Sending email to Virtual Reality News (contact@wonderhowto.com)..." -ForegroundColor Yellow
try {
    $ol = New-Object -ComObject Outlook.Application
    $mail = $ol.CreateItem(0)
    $mail.To = "contact@wonderhowto.com"
    $mail.Subject = "Story Pitch for Virtual Reality News - Gorilla Tag TikTok Creator + Mod Developer"
    $mail.Body = $vrNewsBody
    $mail.Send()
    Write-Host "      SENT via Outlook" -ForegroundColor Green
} catch {
    Write-Host "      Outlook send failed — opening mailto instead" -ForegroundColor Red
    $mailto = "mailto:contact@wonderhowto.com?subject=$([uri]::EscapeDataString('Story Pitch for Virtual Reality News - Gorilla Tag TikTok Creator + Mod Developer'))&body=$([uri]::EscapeDataString($vrNewsBody))"
    Start-Process $mailto
}

Start-Sleep -Seconds 2

# --- Step 2: UploadVR form (email already sent — form only) ---
Write-Host "[2/5] UploadVR contact form (Send Us A Story)..." -ForegroundColor Yellow
Set-Clipboard -Value $uploadvrForm
Start-Process "https://www.uploadvr.com/contact/"
Write-Host "      Form opened — message copied. Select SEND US A STORY, paste, submit." -ForegroundColor Green
Start-Sleep -Seconds 3

# --- Step 3: Another Axiom Creator Program ---
Write-Host "[3/5] Another Axiom Creator Program application..." -ForegroundColor Yellow
$aaNote = @"
Another Axiom Creator Program application notes:

Handle: @frostytagyt
Followers: 28.4K (TikTok)
Niche: Gorilla Tag VR, TikTok LIVE, mod development
Press kit: ${Site}press-kit.html
Email: $Email
TikTok: https://www.tiktok.com/@frostytagyt

Note: Creator program acceptance is NOT a news article — still pitch outlets separately.
"@
Set-Clipboard -Value $aaNote
Start-Process "https://aa-creator-portal.pages.dev/"
Write-Host "      Portal opened — notes copied to clipboard." -ForegroundColor Green
Start-Sleep -Seconds 2

# --- Step 4: mixed.news submit (optional extra outlet) ---
Write-Host "[4/5] MIXED.NEWS submit form (Russian VR site — bonus outlet)..." -ForegroundColor Yellow
Set-Clipboard -Value $mixedNews
Start-Process "https://mixed.news/addnews"
Write-Host "      Form opened — English tip copied. Submit if form allows." -ForegroundColor Green
Start-Sleep -Seconds 2

# --- Step 5: HARO signup (journalists find YOU) ---
Write-Host "[5/5] Opening HARO — sign up as a source (journalists request quotes)..." -ForegroundColor Yellow
Start-Process "https://www.helpareporter.com/sources/"
Write-Host "      Sign up free. Reply to Gorilla Tag / VR / TikTok creator queries." -ForegroundColor Green

Write-Host ""
Write-Host "=== DO NOT RE-EMAIL ===" -ForegroundColor Red
Write-Host "  Road to VR (tips@ / contact@) — already pitched, wait for reply" -ForegroundColor DarkYellow
Write-Host "  UploadVR email (tips@ / dheaney@) — already pitched" -ForegroundColor DarkYellow
Write-Host ""
Write-Host "=== WAITING GAME ===" -ForegroundColor Cyan
Write-Host "  Published articles needed: 1-4 URLs that mention @frostytagyt by name" -ForegroundColor White
Write-Host "  Typical response time: 1-2 weeks. No follow-ups for 7 days." -ForegroundColor White
Write-Host "  When articles publish, paste URLs into TikTok verification." -ForegroundColor White
Write-Host ""
Write-Host "Press kit: $Site" -ForegroundColor Gray
pause