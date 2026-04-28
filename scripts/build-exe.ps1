param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$distRoot = Join-Path $repoRoot "dist"
$packageRoot = Join-Path $distRoot "package"
$assetsSource = Join-Path $repoRoot "assets"
$launcherSource = Join-Path $repoRoot "LaRocheLauncher.bat"
$scriptSource = Join-Path $assetsSource "RutherfordLauncher.ps1"
$exeTarget = Join-Path $packageRoot "RutherfordAssistant.exe"
$zipTarget = Join-Path $distRoot "RutherfordAssistant-portable.zip"

if (-not (Test-Path $scriptSource)) {
    throw "Missing launcher script: $scriptSource"
}

if (-not (Test-Path $assetsSource)) {
    throw "Missing assets folder: $assetsSource"
}

if (Test-Path $distRoot) {
    Remove-Item -Path $distRoot -Recurse -Force
}

New-Item -Path $packageRoot -ItemType Directory -Force | Out-Null

Write-Host "Installing ps2exe..."
Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
Import-Module ps2exe -Force

Write-Host "Compiling RutherfordAssistant.exe..."
Invoke-ps2exe `
    -inputFile $scriptSource `
    -outputFile $exeTarget `
    -noConsole `
    -STA `
    -requireAdmin `
    -DPIAware `
    -title "Rutherford Assistant" `
    -product "Rutherford Assistant" `
    -company "Rutherford" `
    -description "Portable Windows launcher for Rutherford setup and network scripts." `
    -version "1.0.0"

Write-Host "Copying launcher files..."
Copy-Item -Path $launcherSource -Destination (Join-Path $packageRoot "LaRocheLauncher.bat") -Force
Copy-Item -Path $assetsSource -Destination (Join-Path $packageRoot "assets") -Recurse -Force

if (Test-Path $zipTarget) {
    Remove-Item -Path $zipTarget -Force
}

Write-Host "Creating portable zip..."
Compress-Archive -Path (Join-Path $packageRoot "*") -DestinationPath $zipTarget -Force

Write-Host "Build complete:"
Write-Host "EXE: $exeTarget"
Write-Host "ZIP: $zipTarget"
