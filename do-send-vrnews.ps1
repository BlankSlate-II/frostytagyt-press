$Site = "https://blankslate-ii.github.io/frostytagyt-press/"
$Email = "lukedonovanyoung2020@outlook.com"
$body = @"
Hi Virtual Reality News / Next Reality editorial team,

I'm pitching a creator profile that fits your Gorilla Tag coverage — including your recent piece on Meta betting VR's future on teens who grew up in Gorilla Tag.

FrostyTagYT (@frostytagyt on TikTok) is a Gorilla Tag VR creator with 28.4K followers who also develops and releases community mods and tools (wrist watch mods, lobby utilities). Rare combo: entertainer + builder in the VR-native creator generation your team already covers.

Press kit: ${Site}press-kit.html
Creator profile: ${Site}creator-profile.html
TikTok: https://www.tiktok.com/@frostytagyt

Available for interview, quotes, screenshots, or mod demos.

Thanks for considering,
FrostyTagYT
$Email
"@
$ol = New-Object -ComObject Outlook.Application
$mail = $ol.CreateItem(0)
$mail.To = "contact@wonderhowto.com"
$mail.Subject = "Story Pitch for Virtual Reality News - Gorilla Tag TikTok Creator + Mod Developer"
$mail.Body = $body
$mail.Send()
Write-Output "SENT: contact@wonderhowto.com"