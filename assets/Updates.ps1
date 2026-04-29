Write-Host "Welcome on Updates"
Write-Host "Searching, downloading and installing pending Windows updates."
Write-Host "This can take a long time depending on what is pending - please wait."

$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host "`n=== $Message ==="
}

function Test-IsAdmin {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Format-MB {
    param([long]$Bytes)
    if ($Bytes -le 0) { return "0 MB" }
    return ("{0:N1} MB" -f ($Bytes / 1MB))
}

if (-not (Test-IsAdmin)) {
    Write-Host "ERROR: This script must run as Administrator."
    exit 1
}

# ----------------------------------------------------------------------------
# Search for pending updates via the Windows Update Agent COM API.
# ----------------------------------------------------------------------------

Write-Step "Searching for available updates"
$session = $null
$searchResult = $null
try {
    $session = New-Object -ComObject Microsoft.Update.Session
    $session.ClientApplicationID = "Rutherford Assistant"
    $searcher = $session.CreateUpdateSearcher()
    $searcher.Online = $true

    Write-Host "Querying Windows Update servers..."
    $searchResult = $searcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
    $count = $searchResult.Updates.Count
    Write-Host "Found $count update(s) available."
}
catch {
    Write-Host "ERROR: Update search failed - $($_.Exception.Message)"
    exit 1
}

if ($searchResult.Updates.Count -eq 0) {
    Write-Step "Summary"
    Write-Host "No pending updates."
    Write-Host "Windows is already up to date."
    Write-Host "Updates run complete!"
    exit 0
}

# ----------------------------------------------------------------------------
# Filter to updates that are eligible (EULA accepted automatically).
# ----------------------------------------------------------------------------

Write-Step "Preparing update list"
$updatesToProcess = New-Object -ComObject Microsoft.Update.UpdateColl
$totalDownloadBytes = 0L

foreach ($update in $searchResult.Updates) {
    if ($update.IsHidden) { continue }
    if (-not $update.EulaAccepted) {
        try { $update.AcceptEula() } catch { }
    }
    [void]$updatesToProcess.Add($update)
    $totalDownloadBytes += [long]$update.MaxDownloadSize

    $kbList = @()
    foreach ($kb in $update.KBArticleIDs) { $kbList += "KB$kb" }
    $kbStr = if ($kbList.Count -gt 0) { " [" + ($kbList -join ", ") + "]" } else { "" }
    $sizeStr = Format-MB -Bytes $update.MaxDownloadSize
    Write-Host ("  - " + $update.Title + " (" + $sizeStr + ")" + $kbStr)
}

Write-Host ("Total download size: " + (Format-MB -Bytes $totalDownloadBytes))
Write-Host ("Updates queued for processing: " + $updatesToProcess.Count)

if ($updatesToProcess.Count -eq 0) {
    Write-Step "Summary"
    Write-Host "No eligible updates to install."
    Write-Host "Updates run complete!"
    exit 0
}

# ----------------------------------------------------------------------------
# Download phase
# ----------------------------------------------------------------------------

Write-Step "Downloading updates"
try {
    $downloader = $session.CreateUpdateDownloader()
    $downloader.Updates = $updatesToProcess
    Write-Host "Download started. This can take several minutes..."
    $downloadResult = $downloader.Download()

    switch ($downloadResult.ResultCode) {
        2 { Write-Host "All updates downloaded successfully." }
        3 { Write-Host "Download finished with warnings." }
        4 { Write-Host "ERROR: Download failed."; exit 1 }
        5 { Write-Host "ERROR: Download was aborted."; exit 1 }
        default { Write-Host "Download finished with code $($downloadResult.ResultCode)." }
    }
}
catch {
    Write-Host "ERROR: Download phase failed - $($_.Exception.Message)"
    exit 1
}

# Filter to updates that actually downloaded successfully
$downloaded = New-Object -ComObject Microsoft.Update.UpdateColl
foreach ($update in $updatesToProcess) {
    if ($update.IsDownloaded) { [void]$downloaded.Add($update) }
}

if ($downloaded.Count -eq 0) {
    Write-Step "Summary"
    Write-Host "ERROR: No updates were successfully downloaded."
    exit 1
}

Write-Host ("Updates ready to install: " + $downloaded.Count)

# ----------------------------------------------------------------------------
# Install phase
# ----------------------------------------------------------------------------

Write-Step "Installing updates"
$rebootRequired = $false
try {
    $installer = $session.CreateUpdateInstaller()
    $installer.Updates = $downloaded
    Write-Host "Install started. Do not power off the machine..."
    $installResult = $installer.Install()

    switch ($installResult.ResultCode) {
        2 { Write-Host "All updates installed successfully." }
        3 { Write-Host "Install finished with warnings." }
        4 { Write-Host "ERROR: Install failed." }
        5 { Write-Host "ERROR: Install was aborted." }
        default { Write-Host "Install finished with code $($installResult.ResultCode)." }
    }

    if ($installResult.RebootRequired) {
        $rebootRequired = $true
        Write-Host "REBOOT REQUIRED to finalize the installed updates."
    }
}
catch {
    Write-Host "ERROR: Install phase failed - $($_.Exception.Message)"
    exit 1
}

# ----------------------------------------------------------------------------
# Per-update result detail (helps the report)
# ----------------------------------------------------------------------------

Write-Step "Per-update results"
for ($index = 0; $index -lt $downloaded.Count; $index++) {
    $update = $downloaded.Item($index)
    $perResult = $installResult.GetUpdateResult($index)
    $code = $perResult.ResultCode
    $statusText = switch ($code) {
        0 { "Not started" }
        1 { "In progress" }
        2 { "Installed" }
        3 { "Installed with warnings" }
        4 { "Failed" }
        5 { "Aborted" }
        default { "Code $code" }
    }
    Write-Host ("  - " + $statusText + " : " + $update.Title)
}

# ----------------------------------------------------------------------------
# Summary
# ----------------------------------------------------------------------------

Write-Step "Summary"
Write-Host "Updates processed: $($downloaded.Count)"
Write-Host "Total download size: $(Format-MB -Bytes $totalDownloadBytes)"
if ($rebootRequired) {
    Write-Host "Reboot required: yes"
} else {
    Write-Host "Reboot required: no"
}
Write-Host "Updates run complete!"

if ($rebootRequired) {
    exit 2
}
exit 0
