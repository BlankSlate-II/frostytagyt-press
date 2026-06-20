# FrostyTagYT - opens all 4 pitches (minimal clicks)
# Double-click SEND-PITCHES.bat

$Site = "https://blankslate-ii.github.io/frostytagyt-press/"
$myEmail = "lukedonovanyoung2020@outlook.com"

$pitches = @(
    @{
        Name = "Road to VR"
        To = "tips@roadtovr.com"
        Subject = "Story Pitch: Gorilla Tag Creator Building Mods + 28K TikTok Audience"
        Body = @"
Hi Road to VR team,

I'm pitching a story about a rising creator in the Gorilla Tag VR ecosystem.

I'm FrostyTagYT (@frostytagyt on TikTok), a Gorilla Tag content creator with 28.4K followers. I also develop and release mods and community tools for Gorilla Tag players.

Story angles:
- Mid-tier TikTok creators building careers through Gorilla Tag LIVE streaming
- Mod developers in VR gaming who ship tools, not just clips
- Gorilla Tag's TikTok community in 2026

Press kit: ${Site}press-kit.html
TikTok: https://www.tiktok.com/@frostytagyt

Thanks,
FrostyTagYT
$myEmail
"@
        FormUrl = $null
    },
    @{
        Name = "UploadVR"
        Subject = "Story Tip: Gorilla Tag Mod Developer with 28.4K TikTok Following"
        Body = @"
Hi UploadVR team,

FrostyTagYT (@frostytagyt) is a Gorilla Tag VR creator on TikTok with 28.4K followers who also develops community mods and tools.

Press kit: ${Site}press-kit.html
TikTok: https://www.tiktok.com/@frostytagyt

FrostyTagYT
$myEmail
"@
        FormUrl = "https://www.uploadvr.com/contact/"
    },
    @{
        Name = "VR Scout"
        Subject = "Creator Spotlight: FrostyTagYT - Gorilla Tag TikTok LIVE Streamer"
        Body = @"
Hi VR Scout team,

Pitch: FrostyTagYT (@frostytagyt) - 28.4K TikTok followers, LIVE streams, mod development for Gorilla Tag.

Press kit: ${Site}press-kit.html
TikTok: https://www.tiktok.com/@frostytagyt

FrostyTagYT
$myEmail
"@
        FormUrl = "https://vrscout.com/contact/"
    },
    @{
        Name = "Game Developer"
        Subject = "Pitch: Mod Developer Building Audience in Gorilla Tag"
        Body = @"
Hi Game Developer team,

FrostyTagYT (@frostytagyt) - TikTok creator (28.4K) in Gorilla Tag VR who ships mods and community tools.

Press kit: ${Site}press-kit.html
TikTok: https://www.tiktok.com/@frostytagyt

FrostyTagYT
$myEmail
"@
        FormUrl = "https://www.gamedeveloper.com/about/contact-us"
    }
)

Write-Host "Opening all 4 pitches..." -ForegroundColor Green

# Pitch 1: mailto link (opens your email app with everything filled in)
$p1 = $pitches[0]
$mailto = "mailto:$($p1.To)?subject=$([uri]::EscapeDataString($p1.Subject))&body=$([uri]::EscapeDataString($p1.Body))"
Start-Process $mailto
Write-Host "[1/4] Road to VR email opened - click SEND" -ForegroundColor Cyan
Start-Sleep -Seconds 2

# Pitches 2-4: copy + open form
for ($i = 1; $i -lt 4; $i++) {
    $p = $pitches[$i]
    Set-Clipboard -Value $p.Body
    Start-Process $p.FormUrl
    Write-Host "[$($i+1)/4] $($p.Name) - message copied, form opened - paste and submit" -ForegroundColor Cyan
    Start-Sleep -Seconds 1
}

Write-Host ""
Write-Host "All 4 ready. Total clicks needed: SEND + paste/submit x3" -ForegroundColor Green