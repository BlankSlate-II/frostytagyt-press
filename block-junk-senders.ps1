# Block spam senders found in Junk Email, then permanently delete junk + trash
# Keeps FrostyTagYT Media Replies alone — only touches Junk + Deleted Items

$MaxItems = 80
$log = Join-Path $PSScriptRoot "purge-spam.log"
$ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Never block media outlets we are waiting on
$MediaDomains = @(
    "roadtovr.com"
    "uploadvr.com"
    "wonderhowto.com"
    "reality.news"
    "technologyadvice.com"
)

function Is-MediaSender($address) {
    if (-not $address) { return $false }
    $a = $address.ToLower()
    foreach ($d in $MediaDomains) {
        if ($a -like "*@$d" -or $a -like "*$d*") { return $true }
    }
    return $false
}

function Get-SenderSmtp($item) {
    try {
        if ($item.SenderEmailType -eq "SMTP") { return $item.SenderEmailAddress }
        if ($item.SenderEmailAddress -match "@") { return $item.SenderEmailAddress }
    } catch {}
    return $null
}

function Add-BlockedSender($ns, $smtp) {
    if (-not $smtp -or $smtp -notmatch "@") { return $false }
    try {
        $rcp = $ns.CreateRecipient($smtp)
        $rcp.Resolve() | Out-Null
        # Outlook junk list API (works in classic Outlook)
        $ns.AddToBlockedSendersList($rcp)
        return $true
    } catch {
        return $false
    }
}

function Clear-Folder($folder, $max) {
    $n = 0
    while ($folder.Items.Count -gt 0 -and $n -lt $max) {
        try { $folder.Items.Remove(1) } catch { break }
        $n++
    }
    return $n
}

try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $junk = $ns.GetDefaultFolder(23)
    $trash = $ns.GetDefaultFolder(3)

    $blocked = 0
    $skippedMedia = 0
    $seen = @{}

    $count = [Math]::Min($junk.Items.Count, $MaxItems)
    for ($i = $count; $i -ge 1; $i--) {
        try {
            $item = $junk.Items.Item($i)
            $smtp = Get-SenderSmtp $item
            if ($smtp -and -not $seen.ContainsKey($smtp)) {
                $seen[$smtp] = $true
                if (Is-MediaSender $smtp) {
                    $skippedMedia++
                } elseif (Add-BlockedSender $ns $smtp) {
                    $blocked++
                }
            }
        } catch {}
    }

    # Rescue media replies before purging junk
    $inbox = $ns.GetDefaultFolder(6)
    $mediaFolder = $null
    foreach ($f in $inbox.Folders) {
        if ($f.Name -eq "FrostyTagYT Media Replies") { $mediaFolder = $f; break }
    }
    $rescued = 0
    if ($mediaFolder) {
        $rMax = [Math]::Min($junk.Items.Count, 15)
        for ($i = $rMax; $i -ge 1; $i--) {
            try {
                $item = $junk.Items.Item($i)
                if (Is-MediaSender (Get-SenderSmtp $item)) {
                    $item.Move($mediaFolder) | Out-Null
                    $rescued++
                }
            } catch {}
        }
    }

    $junkPurged = Clear-Folder $junk $MaxItems
    $trashPurged = Clear-Folder $trash $MaxItems

    $line = "$ts | blocked=$blocked rescued=$rescued media_skipped=$skippedMedia | junk_purged=$junkPurged junk_left=$($junk.Items.Count) | trash_purged=$trashPurged trash_left=$($trash.Items.Count)"
    Add-Content -Path $log -Value $line -ErrorAction SilentlyContinue
    Write-Output $line
} catch {
    $err = "$ts | ERROR: $($_.Exception.Message)"
    Add-Content -Path $log -Value $err -ErrorAction SilentlyContinue
    Write-Error $err
    exit 1
}