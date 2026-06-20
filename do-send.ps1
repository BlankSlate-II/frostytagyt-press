$Site = "https://blankslate-ii.github.io/frostytagyt-press/"
$ol = New-Object -ComObject Outlook.Application
$mail = $ol.CreateItem(0)
$mail.To = "tips@roadtovr.com"
$mail.Subject = "Story Pitch: Gorilla Tag Creator Building Mods + 28K TikTok Audience"
$mail.Body = @"
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
lukedonovanyoung2020@outlook.com
"@
$mail.Send()
Write-Output "SENT_OK"