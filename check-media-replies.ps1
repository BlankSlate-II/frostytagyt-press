# Daily check: find media outlet replies, rescue from junk, build reply drafts

Add-Type -AssemblyName System.Web

$Site = "https://blankslate-ii.github.io/frostytagyt-press/"
$Email = "lukedonovanyoung2020@outlook.com"
$MediaFolderName = "FrostyTagYT Media Replies"
$MaxScan = 40
$Root = $PSScriptRoot
$SeenFile = Join-Path $Root "media-replies-seen.json"
$ReportFile = Join-Path $Root "media-replies-report.html"
$LogFile = Join-Path $Root "media-check.log"

$MediaDomains = @(
    "roadtovr.com", "uploadvr.com", "wonderhowto.com",
    "reality.news", "technologyadvice.com"
)

function Is-MediaSender($address) {
    if (-not $address) { return $false }
    $a = $address.ToLower()
    foreach ($d in $MediaDomains) {
        if ($a -like "*$d*") { return $true }
    }
    return $false
}

function Get-BodyPreview($item) {
    try {
        $b = $item.Body
        if (-not $b) { return "" }
        $b = $b -replace "<[^>]+>", " "
        $b = $b -replace "\s+", " "
        if ($b.Length -gt 600) { $b = $b.Substring(0, 600) + "..." }
        return $b.Trim()
    } catch { return "" }
}

function H([string]$s) { return [System.Web.HttpUtility]::HtmlEncode($s) }

function Get-ReplyDraft($subject, $body) {
    $text = ("$subject $body").ToLower()
    $handle = '@frostytagyt'
    $tiktok = 'https://www.tiktok.com/@frostytagyt'

    if ($text -match "not interested|unable to cover|pass on|decline|no coverage|not a fit") {
        return @{
            Type = "Pass - optional thank-you only"
            Subject = "Re: $subject"
            Body = "Hi,`n`nThank you for getting back to me. If you need a Gorilla Tag / VR creator source later, I'm happy to help.`n`nBest,`nFrostyTagYT`n$Email"
        }
    }
    if ($text -match "interview|call|chat|speak with|zoom|phone|questions for you") {
        return @{
            Type = "Interview request - reply ASAP"
            Subject = "Re: $subject"
            Body = "Hi,`n`nThank you - I'd love to interview.`n`nI'm FrostyTagYT ($handle), Gorilla Tag VR creator with 28.4K TikTok followers. I also develop community mods and tools for the game.`n`nPress kit: ${Site}press-kit.html`nTikTok: $tiktok`n`nI'm flexible on timing - let me know what works.`n`nBest,`nFrostyTagYT`n$Email"
        }
    }
    if ($text -match "more info|additional information|send us|details|assets|photos|screenshots|metrics|followers|stats|press kit|bio") {
        return @{
            Type = "They need more info - send everything"
            Subject = "Re: $subject"
            Body = "Hi,`n`nHappy to send more:`n`nCreator: FrostyTagYT ($handle)`nFollowers: 28.4K on TikTok`nNiche: Gorilla Tag VR, TikTok LIVE, mod development`n`nPress kit: ${Site}press-kit.html`nProfile: ${Site}creator-profile.html`nTikTok: $tiktok`n`nAvailable for quotes, screenshots, mod demos, or a call.`n`nThanks,`nFrostyTagYT`n$Email"
        }
    }
    if ($text -match "received your|we got your|thank you for (reaching|contacting|submitting)|will review|in our queue") {
        return @{
            Type = "Auto-reply - wait, do NOT email again yet"
            Subject = ""
            Body = "No reply needed. They acknowledged your pitch. Wait 5-7 days."
        }
    }
    if ($text -match "publish|article|story|feature|writer assigned|run it|going live|draft") {
        return @{
            Type = "POSITIVE - they may publish!"
            Subject = "Re: $subject"
            Body = "Hi,`n`nThat's great - thank you!`n`nPlease use $handle on first reference: $tiktok`nFollower count: 28.4K`n`nPress kit for fact-checking: ${Site}press-kit.html`n`nLet me know if you need quotes or images before publish.`n`nBest,`nFrostyTagYT`n$Email"
        }
    }
    return @{
        Type = "General reply"
        Subject = "Re: $subject"
        Body = "Hi,`n`nThank you for responding.`n`nFrostyTagYT ($handle) - Gorilla Tag VR creator, 28.4K TikTok followers, mod developer.`n`nPress kit: ${Site}press-kit.html`nTikTok: $tiktok`n`nHappy to answer questions, send assets, or do a quick call.`n`nBest,`nFrostyTagYT`n$Email"
    }
}

function Load-Seen {
    if (Test-Path $SeenFile) {
        $j = Get-Content $SeenFile -Raw | ConvertFrom-Json
        return [System.Collections.Generic.HashSet[string]]::new([string[]]@($j.ids))
    }
    return [System.Collections.Generic.HashSet[string]]::new()
}

function Save-Seen($set) {
    @{ ids = @($set) } | ConvertTo-Json | Set-Content $SeenFile -Encoding UTF8
}

$ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$found = New-Object System.Collections.Generic.List[object]
$rescued = 0

try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)
    $junk = $ns.GetDefaultFolder(23)

    $mediaFolder = $null
    foreach ($f in $inbox.Folders) {
        if ($f.Name -eq $MediaFolderName) { $mediaFolder = $f; break }
    }
    if (-not $mediaFolder) { $mediaFolder = $inbox.Folders.Add($MediaFolderName) }

    $jMax = [Math]::Min($junk.Items.Count, 15)
    for ($i = $jMax; $i -ge 1; $i--) {
        try {
            $item = $junk.Items.Item($i)
            if (Is-MediaSender $item.SenderEmailAddress) {
                $item.Move($mediaFolder) | Out-Null
                $rescued++
            }
        } catch {}
    }

    $seenIds = Load-Seen
    foreach ($folder in @($inbox, $mediaFolder)) {
        $limit = [Math]::Min($folder.Items.Count, $MaxScan)
        for ($i = 1; $i -le $limit; $i++) {
            try {
                $item = $folder.Items.Item($i)
                if ($seenIds.Contains($item.EntryID)) { continue }
                if (-not (Is-MediaSender $item.SenderEmailAddress)) { continue }

                $subj = $item.Subject
                $body = Get-BodyPreview $item
                $draft = Get-ReplyDraft $subj $body

                $found.Add([PSCustomObject]@{
                    From = $item.SenderEmailAddress
                    Subject = $subj
                    Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
                    Preview = $body
                    DraftType = $draft.Type
                    DraftSubject = $draft.Subject
                    DraftBody = $draft.Body
                    Unread = $item.UnRead
                    Id = $item.EntryID
                }) | Out-Null
                $null = $seenIds.Add($item.EntryID)
            } catch {}
        }
    }
    Save-Seen $seenIds

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("<!DOCTYPE html><html lang=`"en`"><head><meta charset=`"UTF-8`"><title>Media Replies $ts</title>")
    [void]$sb.AppendLine("<style>body{font-family:Segoe UI,sans-serif;background:#0d1117;color:#e6edf3;max-width:800px;margin:2rem auto;padding:0 1rem}")
    [void]$sb.AppendLine(".card{background:#161b22;border:1px solid #30363d;border-radius:12px;padding:1.25rem;margin:1rem 0}")
    [void]$sb.AppendLine(".urgent{border-color:#34d399}pre{background:#0d1117;padding:1rem;border-radius:8px;white-space:pre-wrap}")
    [void]$sb.AppendLine("button{background:#238636;color:#fff;border:none;padding:0.5rem 1rem;border-radius:8px;cursor:pointer}")
    [void]$sb.AppendLine("h1{color:#7dd3fc}.muted{color:#8b949e}</style>")
    [void]$sb.AppendLine("<script>function copyReply(b){var p=b.parentElement.querySelector('pre');navigator.clipboard.writeText(p.textContent).then(()=>b.textContent='Copied!')}</script></head><body>")
    [void]$sb.AppendLine("<h1>Media Reply Check</h1>")
    [void]$sb.AppendLine("<p class=`"muted`">$ts | Rescued from junk: $rescued | New: $($found.Count)</p>")
    [void]$sb.AppendLine("<p>Only emails from <strong>Road to VR, UploadVR, Virtual Reality News</strong> show here.</p>")

    if ($found.Count -eq 0) {
        [void]$sb.AppendLine("<div class=`"card`"><p>No new media replies. Normal wait: 1-2 weeks. Check <strong>FrostyTagYT Media Replies</strong> in Outlook.</p></div>")
    } else {
        foreach ($m in $found) {
            $cls = if ($m.DraftType -match "ASAP|POSITIVE|need more") { "card urgent" } else { "card" }
            [void]$sb.AppendLine("<div class=`"$cls`"><h3>$(H $m.From)</h3>")
            [void]$sb.AppendLine("<p><strong>Subject:</strong> $(H $m.Subject)</p>")
            [void]$sb.AppendLine("<p><strong>When:</strong> $($m.Received) | <strong>Action:</strong> $(H $m.DraftType)</p>")
            [void]$sb.AppendLine("<p class=`"muted`">$(H $m.Preview)</p><h4>Suggested reply</h4><pre>$(H $m.DraftBody)</pre>")
            if ($m.DraftSubject) {
                [void]$sb.AppendLine("<p><strong>Subject line:</strong> $(H $m.DraftSubject)</p>")
            }
            [void]$sb.AppendLine("<button onclick=`"copyReply(this)`">Copy reply to send in Outlook</button></div>")
        }
    }
    [void]$sb.AppendLine("</body></html>")
    Set-Content -Path $ReportFile -Value $sb.ToString() -Encoding UTF8

    $line = "$ts | new=$($found.Count) rescued=$rescued"
    Add-Content -Path $LogFile -Value $line
    Write-Output $line
    Write-Output "REPORT: $ReportFile"
    Start-Process $ReportFile
} catch {
    Add-Content -Path $LogFile -Value "$ts | ERROR: $($_.Exception.Message)"
    Write-Error $_.Exception.Message
    exit 1
}