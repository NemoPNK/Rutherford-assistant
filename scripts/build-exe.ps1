param(
    [string]$Configuration = "Release",
    [string]$Version       = "1.1.0",

    # Optional Authenticode signing.
    # Provide either a pfx path + password, or a certificate thumbprint already in the user store.
    [string]$SigningPfxPath        = "",
    [string]$SigningPfxPassword    = "",
    [string]$SigningCertThumbprint = "",
    [string]$TimestampUrl          = "http://timestamp.digicert.com"
)

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

$repoRoot       = Split-Path -Parent $PSScriptRoot
$distRoot       = Join-Path $repoRoot "dist"
$packageRoot    = Join-Path $distRoot "package"
$assetsSource   = Join-Path $repoRoot "assets"
$scriptSource   = Join-Path $assetsSource "RutherfordLauncher.ps1"
$exeTarget      = Join-Path $packageRoot "RutherfordAssistant.exe"
$zipTarget      = Join-Path $distRoot "RutherfordAssistant.zip"
$checksumTarget = Join-Path $distRoot "RutherfordAssistant.sha256"

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

# ---------------------------------------------------------------------------
# Pre-build inventory (helps debug missing manifests / checks)
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "=== Build inventory ==="
Write-Host "Repo root  : $repoRoot"
Write-Host "Version    : $Version"

$manifestFiles = @(Get-ChildItem -Path $assetsSource -Filter "*.manifest.json" -File -ErrorAction SilentlyContinue)
Write-Host "Script manifests ($($manifestFiles.Count)):"
foreach ($m in $manifestFiles) {
    $base = $m.BaseName
    if ($base.EndsWith(".manifest")) { $base = $base.Substring(0, $base.Length - ".manifest".Length) }
    $sister = Join-Path $m.DirectoryName ("$base.ps1")
    $status = if (Test-Path $sister) { "ok" } else { "MISSING SCRIPT" }
    Write-Host "  - $($m.Name) -> $base.ps1 [$status]"
}

$checkRoot = Join-Path $assetsSource "checks"
$checkFiles = @(Get-ChildItem -Path $checkRoot -Filter "*.check.ps1" -File -ErrorAction SilentlyContinue)
Write-Host "Audit checks ($($checkFiles.Count)):"
foreach ($c in $checkFiles) {
    Write-Host "  - $($c.Name)"
}

if ($manifestFiles.Count -eq 0) {
    Write-Warning "No script manifest found. The launcher will start with no action buttons."
}
if ($checkFiles.Count -eq 0) {
    Write-Warning "No audit check found. The LaRoche Audit panel will be empty."
}
Write-Host ""

# ---------------------------------------------------------------------------
# ps2exe compile
# ---------------------------------------------------------------------------

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
    -version $Version

if (-not (Test-Path $exeTarget)) {
    throw "ps2exe did not produce $exeTarget"
}

# ---------------------------------------------------------------------------
# Optional code signing
# ---------------------------------------------------------------------------

$cert = $null

if ($SigningPfxPath -and (Test-Path $SigningPfxPath)) {
    Write-Host "Loading signing cert from PFX: $SigningPfxPath"
    if ($SigningPfxPassword) {
        $securePassword = ConvertTo-SecureString -String $SigningPfxPassword -AsPlainText -Force
        $cert = Get-PfxCertificate -FilePath $SigningPfxPath -Password $securePassword
    }
    else {
        $cert = Get-PfxCertificate -FilePath $SigningPfxPath
    }
}
elseif ($SigningCertThumbprint) {
    Write-Host "Loading signing cert by thumbprint: $SigningCertThumbprint"
    $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$SigningCertThumbprint" -ErrorAction SilentlyContinue
    if (-not $cert) {
        $cert = Get-ChildItem -Path "Cert:\LocalMachine\My\$SigningCertThumbprint" -ErrorAction SilentlyContinue
    }
}

if ($cert) {
    Write-Host "Signing $exeTarget with $($cert.Subject)"
    $sig = Set-AuthenticodeSignature -FilePath $exeTarget -Certificate $cert -TimestampServer $TimestampUrl -HashAlgorithm SHA256
    if ($sig.Status -ne "Valid") {
        Write-Warning "Signature status: $($sig.Status). Continuing without failing the build."
    }
    else {
        Write-Host "Signature: Valid (subject = $($cert.Subject))"
    }
}
else {
    Write-Host "No signing cert provided; producing unsigned EXE."
    Write-Host "  -> SmartScreen may show a warning the first few times the EXE is launched."
    Write-Host "  -> To sign, pass -SigningPfxPath / -SigningPfxPassword or -SigningCertThumbprint."
}

# ---------------------------------------------------------------------------
# Stage package contents
# ---------------------------------------------------------------------------

Write-Host "Copying assets folder..."
Copy-Item -Path $assetsSource   -Destination (Join-Path $packageRoot "assets") -Recurse -Force

# Strip macOS metadata that sometimes makes it into the repo
Get-ChildItem -Path $packageRoot -Recurse -Force -Filter ".DS_Store" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue

# Build a CONTENTS.txt manifest listing every file shipped in the zip
$contentsTarget = Join-Path $packageRoot "CONTENTS.txt"
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("Rutherford Assistant portable package") | Out-Null
$lines.Add("Version: $Version") | Out-Null
$lines.Add("Built: $((Get-Date).ToString('s'))") | Out-Null
$lines.Add("") | Out-Null
$lines.Add("Files:") | Out-Null
Get-ChildItem -Path $packageRoot -Recurse -File | ForEach-Object {
    $relative = $_.FullName.Substring($packageRoot.Length + 1)
    $size = "{0,10}" -f $_.Length
    $lines.Add("$size  $relative") | Out-Null
}
Set-Content -Path $contentsTarget -Value $lines -Encoding UTF8

# ---------------------------------------------------------------------------
# Zip + checksum
# ---------------------------------------------------------------------------

if (Test-Path $zipTarget) {
    Remove-Item -Path $zipTarget -Force
}

Write-Host "Creating portable zip..."
Compress-Archive -Path (Join-Path $packageRoot "*") -DestinationPath $zipTarget -Force

$zipHash = (Get-FileHash -Path $zipTarget -Algorithm SHA256).Hash
$exeHash = (Get-FileHash -Path $exeTarget -Algorithm SHA256).Hash

@(
    "$zipHash  RutherfordAssistant.zip"
    "$exeHash  package/RutherfordAssistant.exe"
) | Set-Content -Path $checksumTarget -Encoding ASCII

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "=== Build complete ==="
Write-Host "EXE      : $exeTarget"
Write-Host "ZIP      : $zipTarget"
Write-Host "Checksum : $checksumTarget"
Write-Host "EXE SHA256 : $exeHash"
Write-Host "ZIP SHA256 : $zipHash"

if (-not $cert) {
    Write-Host ""
    Write-Host "Reminder: this build is unsigned. SmartScreen may show 'Windows protected your PC'"
    Write-Host "the first time the EXE runs. Click 'More info' -> 'Run anyway' to bypass once,"
    Write-Host "or sign the EXE with a code-signing certificate to remove the prompt."
}
