$Site = "https://blankslate-ii.github.io/frostytagyt-press/"

$pitches = @(
    @{
        Name = "VR Scout"
        Url = "https://vrscout.com/contact/"
        Body = @"
Subject: Creator Spotlight Pitch: FrostyTagYT - Gorilla Tag TikTok LIVE Streamer

Hi VR Scout team,

I'd like to pitch a creator spotlight on FrostyTagYT (@frostytagyt).

28.4K TikTok followers, regular Gorilla Tag LIVE streams, and active mod development for the VR community. Strong fit for VR Scout's creator audience.

Press kit: ${Site}press-kit.html
Profile: ${Site}creator-profile.html
TikTok: https://www.tiktok.com/@frostytagyt
Email: lukedonovanyoung2020@outlook.com

Available for interview.

FrostyTagYT
"@
    },
    @{
        Name = "Game Developer"
        Url = "https://www.gamedeveloper.com/about/contact-us"
        Body = @"
Subject: Pitch: Grassroots Mod Developer Building Audience in Gorilla Tag

Hi Game Developer team,

Pitch for a community/creator story:

FrostyTagYT (@frostytagyt) - TikTok creator (28.4K followers) in the Gorilla Tag VR space who also ships mods and community tools, not just content. Relevant to creator economies and live service communities.

Press kit: ${Site}press-kit.html
TikTok: https://www.tiktok.com/@frostytagyt
Email: lukedonovanyoung2020@outlook.com

Happy to provide quotes or technical details.

FrostyTagYT
"@
    }
)

Write-Host "Opening your LAST 2 outlet pitches (VR Scout + Game Developer)..." -ForegroundColor Cyan
Write-Host "Paste with Ctrl+V on each form and submit.`n" -ForegroundColor Yellow

foreach ($p in $pitches) {
    Set-Clipboard -Value $p.Body
    Start-Process $p.Url
    Write-Host "[OPENED] $($p.Name) - message copied to clipboard" -ForegroundColor Green
    Start-Sleep -Seconds 2
}

Write-Host "`nDone. Submit both forms = 4 outlets total pitched." -ForegroundColor Cyan
Write-Host "Wait for them to PUBLISH articles, then paste real URLs in TikTok."
pause