# FrostyTagYT media inbox: regular folder + Search Folder (auto-collects replies)

$FolderName = "FrostyTagYT Media Replies"
$SearchName = "FrostyTagYT Media Search"

$FromAddresses = @(
    "tips@roadtovr.com",
    "contact@roadtovr.com",
    "dheaney@uploadvr.com",
    "tips@uploadvr.com",
    "contact@wonderhowto.com"
)

$Domains = @("roadtovr.com", "uploadvr.com", "wonderhowto.com", "reality.news")

$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6)

# Regular folder (for manual drag or if you add a rule in Outlook UI later)
$mediaFolder = $null
foreach ($f in $inbox.Folders) {
    if ($f.Name -eq $FolderName) { $mediaFolder = $f; break }
}
if (-not $mediaFolder) {
    $mediaFolder = $inbox.Folders.Add($FolderName)
    Write-Output "CREATED_FOLDER: $FolderName"
} else {
    Write-Output "FOLDER_EXISTS: $FolderName"
}

# Remove old search folder if present
$searchRoot = $ns.GetDefaultFolder(18) # olFolderSearchFolders
foreach ($f in $searchRoot.Folders) {
    if ($f.Name -eq $SearchName) {
        $f.Delete()
        Write-Output "REMOVED_OLD_SEARCH"
    }
}

# Build filter: from specific addresses OR sender domain contains
$fromParts = $FromAddresses | ForEach-Object { "from:$_" }
$domainParts = $Domains | ForEach-Object { "from:@$_" }
$filter = ($fromParts + $domainParts) -join " OR "

$searchFolder = $searchRoot.Folders.Add($SearchName)
$searchFolder.Scope = "Inbox"
$searchFolder.Filter = $filter
$searchFolder.Save()
$searchFolder.Display()

Write-Output "SEARCH_FOLDER: $SearchName"
Write-Output "FILTER: $filter"
Write-Output "FOLDER_PATH: $($mediaFolder.FolderPath)"
Write-Output "SUCCESS"