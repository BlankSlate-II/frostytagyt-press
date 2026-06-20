$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6)
$inbox.Display()
Start-Sleep -Seconds 2

$opened = $false
foreach ($i in 35..45) {
    try {
        $dlg = $ol.Dialogs.Item($i)
        if ($dlg) {
            $dlg.Show()
            Write-Output "OPENED_DIALOG_$i"
            $opened = $true
            break
        }
    } catch {}
}

if (-not $opened) {
    # Fallback: open Outlook options page for rules (classic)
    Start-Process "outlook.exe"
    Write-Output "OPENED_OUTLOOK_MANUAL: File > Manage Rules & Alerts"
}