$Site = "https://blankslate-ii.github.io/frostytagyt-press/"
$ol = New-Object -ComObject Outlook.Application

$emails = @(
    @{
        To = "tips@roadtovr.com"
        Subject = "Story Pitch: Gorilla Tag Creator Building Mods + 28K TikTok Audience"
        Body = "Hi Road to VR team,`n`nPitch: FrostyTagYT (@frostytagyt) is a Gorilla Tag VR creator on TikTok with 28.4K followers who also develops community mods and tools.`n`nPress kit: ${Site}press-kit.html`nTikTok: https://www.tiktok.com/@frostytagyt`n`nAvailable for interview.`n`nThanks,`nFrostyTagYT`nlukedonovanyoung2020@outlook.com"
    },
    @{
        To = "dheaney@uploadvr.com"
        Subject = "Pitch: Gorilla Tag Mod Developer with 28.4K TikTok Following"
        Body = "Hi David,`n`nPitch: FrostyTagYT (@frostytagyt) is a Gorilla Tag VR creator on TikTok with 28.4K followers who also develops community mods and tools - community-driven creator story for UploadVR.`n`nPress kit: ${Site}press-kit.html`nTikTok: https://www.tiktok.com/@frostytagyt`n`nHappy to provide interview or demos.`n`nFrostyTagYT`nlukedonovanyoung2020@outlook.com"
    },
    @{
        To = "tips@uploadvr.com"
        Subject = "Story Tip: Gorilla Tag TikTok Creator and Mod Developer"
        Body = "Hi UploadVR editorial,`n`nStory tip: FrostyTagYT (@frostytagyt) - 28.4K TikTok followers, Gorilla Tag LIVE streams, mod development.`n`nPress kit: ${Site}press-kit.html`nTikTok: https://www.tiktok.com/@frostytagyt`n`nFrostyTagYT`nlukedonovanyoung2020@outlook.com"
    },
    @{
        To = "contact@roadtovr.com"
        Subject = "Media Inquiry: FrostyTagYT Gorilla Tag Creator Profile"
        Body = "Hi Road to VR,`n`nMedia inquiry regarding FrostyTagYT (@frostytagyt), Gorilla Tag creator with 28.4K TikTok followers and mod developer in the VR gaming space.`n`nPress kit: ${Site}press-kit.html`nCreator profile: ${Site}creator-profile.html`nTikTok: https://www.tiktok.com/@frostytagyt`n`nFrostyTagYT`nlukedonovanyoung2020@outlook.com"
    }
)

$results = @()
foreach ($e in $emails) {
    $mail = $ol.CreateItem(0)
    $mail.To = $e.To
    $mail.Subject = $e.Subject
    $mail.Body = $e.Body
    $mail.Send()
    $results += "SENT: $($e.To)"
    Start-Sleep -Milliseconds 800
}

$results -join "`n"
Write-Output "ALL_DONE"