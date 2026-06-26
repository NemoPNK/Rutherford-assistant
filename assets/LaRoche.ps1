Write-Host "Welcome on La Roche"
Write-Host "Thank you for your patience, it may take some time."

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

function Ensure-RegistryValue {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$true)]$Value,
        [Parameter(Mandatory=$true)][string]$PropertyType
    )

    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }

    New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $PropertyType -Force | Out-Null
}

# Appx removals that failed after retry are collected here and reported in the summary.
$script:AppxFailures = @()

function Remove-AppxEverywhere {
    param(
        [Parameter(Mandatory=$true)][string]$Name
    )

    $installed   = @(Get-AppxPackage -Name $Name -AllUsers -ErrorAction SilentlyContinue)
    $provisioned = @(Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -eq $Name })

    if ($installed.Count -eq 0 -and $provisioned.Count -eq 0) {
        Write-Host "Appx ${Name}: not present."
        return
    }

    Write-Host "Removing Appx: $Name ..."

    foreach ($attempt in 1..2) {
        foreach ($pkg in $installed) {
            try {
                Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction Stop
            }
            catch {
                Write-Host "  attempt ${attempt}: $($pkg.PackageFullName): $($_.Exception.Message)"
            }
        }

        foreach ($prov in $provisioned) {
            try {
                Remove-AppxProvisionedPackage -Online -PackageName $prov.PackageName -ErrorAction Stop | Out-Null
            }
            catch {
                Write-Host "  attempt ${attempt} (provisioned): $($prov.PackageName): $($_.Exception.Message)"
            }
        }

        # Verify instead of assuming success
        $installed   = @(Get-AppxPackage -Name $Name -AllUsers -ErrorAction SilentlyContinue)
        $provisioned = @(Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -eq $Name })

        if ($installed.Count -eq 0 -and $provisioned.Count -eq 0) {
            Write-Host "Appx ${Name}: removed (verified, attempt $attempt)."
            return
        }

        Start-Sleep -Seconds 2
    }

    $leftover = @()
    if ($installed.Count -gt 0)   { $leftover += "installed" }
    if ($provisioned.Count -gt 0) { $leftover += "provisioned" }
    Write-Host "WARNING Appx ${Name}: STILL PRESENT ($($leftover -join '+')) after 2 attempts."
    $script:AppxFailures += $Name
}

# Helper functions for Windows 11 Start menu policy
function Get-WindowsBuildNumber {
    try {
        return [int](Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name CurrentBuildNumber -ErrorAction Stop).CurrentBuildNumber
    }
    catch {
        return 0
    }
}

function Set-Windows11StartPolicy {
    $buildNumber = Get-WindowsBuildNumber
    $startExplorerPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer"
    $layoutJson = '{"pinnedList":[]}'

    if (-not (Test-Path $startExplorerPolicyPath)) {
        New-Item -Path $startExplorerPolicyPath -Force | Out-Null
    }

    # --- Recommended section ---
    # Per Microsoft Policy CSP doc, HideRecommendedSection is only honored on
    # Windows 11 22H2 (build 22621) and later. On older builds the value is ignored.
    if ($buildNumber -ge 22621) {
        New-ItemProperty -Path $startExplorerPolicyPath -Name "HideRecommendedSection" -Value 1 -PropertyType DWord -Force | Out-Null
        Write-Host "HideRecommendedSection=1 applied (build $buildNumber)."
    }
    else {
        Write-Host "WARNING: build $buildNumber < 22621 - HideRecommendedSection is IGNORED by this Windows version; Recommended section cannot be hidden by policy on this PC."
    }

    # Best-effort user-level Start toggles (Settings > Personalization > Start), any build:
    # disable recommendation content and use the 'More pins' layout.
    Ensure-RegistryValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Start_IrisRecommendations" -Value 0 -PropertyType DWord
    Ensure-RegistryValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Start_Layout" -Value 1 -PropertyType DWord
    Write-Host "User Start toggles set: recommendations content off, 'More pins' layout."

    # --- Pinned apps ---
    # Per Microsoft Policy CSP doc, ConfigureStartPins is supported from Windows 11 21H2
    # (build 22000). The GPO-mapped registry value must contain the JSON itself.
    if ($buildNumber -ge 22000) {
        New-ItemProperty -Path $startExplorerPolicyPath -Name "ConfigureStartPins" -Value $layoutJson -PropertyType String -Force | Out-Null

        # Belt-and-suspenders: also seed the MDM PolicyManager entry with the same JSON.
        $startPolicyManagerPath = "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Start"
        if (-not (Test-Path $startPolicyManagerPath)) {
            New-Item -Path $startPolicyManagerPath -Force | Out-Null
        }
        New-ItemProperty -Path $startPolicyManagerPath -Name "ConfigureStartPins" -Value $layoutJson -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $startPolicyManagerPath -Name "ConfigureStartPins_ProviderSet" -Value 1 -PropertyType DWord -Force | Out-Null

        Write-Host "ConfigureStartPins applied with empty pinnedList (build $buildNumber): pinned section emptied."
    }
    else {
        Write-Host "WARNING: build $buildNumber < 22000 - ConfigureStartPins unsupported; pinned apps left as-is."
    }
}

if (-not (Test-IsAdmin)) {
    Write-Host "This script must be run as Administrator."
    exit 1
}

# Copy folder preinstall/OPS to C:\
$sourcePath = Join-Path $PSScriptRoot "preinstall/OPS"
$destinationPath = "C:\OPS"

if (Test-Path $sourcePath) {
    if (-not (Test-Path $destinationPath)) {
        New-Item -Path $destinationPath -ItemType Directory -Force | Out-Null
    }

    Copy-Item -Path (Join-Path $sourcePath "*") -Destination $destinationPath -Recurse -Force
    Write-Host "OPS folder copied to C:\"
    
    try {
        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.Namespace($destinationPath)

        if ($folder -ne $null) {
            $folder.Self.InvokeVerb("pintohome")
            Write-Host "OPS added to Quick Access."
        }
        else {
            Write-Host "Unable to locate $destinationPath to pin it in Quick Access."
        }
    }
    catch {
        Write-Host "Error pinning folder to Quick Access: $_"
    }
}
else {
    Write-Host "Source folder not found: $sourcePath"
}

# Désactivation de la veille
powercfg /change standby-timeout-ac 0
powercfg /change standby-timeout-dc 0
powercfg /change monitor-timeout-ac 0
powercfg /change monitor-timeout-dc 0
powercfg /hibernate off

# Wallpaper
$wallpaperPath = Join-Path $PSScriptRoot "wallpaper.jpg"
$localWallpaperDirectory = "C:\ProgramData\Rutherford"
$localWallpaperPath = Join-Path $localWallpaperDirectory "wallpaper.jpg"

if (Test-Path $wallpaperPath) {
    if (-not (Test-Path $localWallpaperDirectory)) {
        New-Item -Path $localWallpaperDirectory -ItemType Directory -Force | Out-Null
    }

    Copy-Item -Path $wallpaperPath -Destination $localWallpaperPath -Force
    Write-Host "Wallpaper copied to $localWallpaperPath"

    Add-Type @"
using System.Runtime.InteropServices;
public class Wallpaper {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@

    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name Wallpaper -Value $localWallpaperPath
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -Value "10"
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -Value "0"

    [Wallpaper]::SystemParametersInfo(20, 0, $localWallpaperPath, 0x1 -bor 0x2) | Out-Null

    Write-Host "Wallpaper set"
}
else {
    Write-Host "cant found wallpaper $wallpaperPath"
}

# Block silent auto-install of sponsored/suggested apps via Content Delivery Manager.
# This is what actually prevents removed apps from coming back on Pro edition:
# the CloudContent policies (DisableWindowsConsumerFeatures) are IGNORED on Pro since Win10 1607.
Write-Host "Disabling Content Delivery Manager auto-installs..."

$cdmPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
$cdmValues = [ordered]@{
    "ContentDeliveryAllowed"          = 0
    "FeatureManagementEnabled"        = 0
    "OemPreInstalledAppsEnabled"      = 0
    "PreInstalledAppsEnabled"         = 0
    "PreInstalledAppsEverEnabled"     = 0
    "SilentInstalledAppsEnabled"      = 0
    "SoftLandingEnabled"              = 0
    "SubscribedContentEnabled"        = 0
    "SubscribedContent-338388Enabled" = 0
    "SubscribedContent-338389Enabled" = 0
    "SystemPaneSuggestionsEnabled"    = 0
}

foreach ($valueName in $cdmValues.Keys) {
    Ensure-RegistryValue -Path $cdmPath -Name $valueName -Value $cdmValues[$valueName] -PropertyType DWord
}
Write-Host "Content Delivery Manager disabled for current user."

# Apply the same values to the Default profile so any future account is covered too.
$defaultHive = "C:\Users\Default\NTUSER.DAT"
if (Test-Path $defaultHive) {
    $hiveLoaded = $false
    try {
        reg.exe load "HKU\RutherfordDefault" $defaultHive | Out-Null
        $hiveLoaded = $true
        foreach ($valueName in $cdmValues.Keys) {
            reg.exe add "HKU\RutherfordDefault\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" /v $valueName /t REG_DWORD /d $cdmValues[$valueName] /f | Out-Null
        }
        Write-Host "Content Delivery Manager disabled in Default profile (future accounts)."
    }
    catch {
        Write-Host "Default profile hive update failed: $($_.Exception.Message)"
    }
    finally {
        if ($hiveLoaded) {
            [gc]::Collect()
            [gc]::WaitForPendingFinalizers()
            try { reg.exe unload "HKU\RutherfordDefault" | Out-Null } catch { Write-Host "Hive unload deferred (will release at next reboot)." }
        }
    }
}
else {
    Write-Host "Default profile hive not found - skipped."
}

# Suppression des applications (UWP/Appx) indésirables
# NOTE: Certaines applis (ex: OneDrive / Office desktop) ne sont pas des Appx et nécessitent un traitement séparé.

$appsToRemove = @(
    "Microsoft.MicrosoftSolitaireCollection",
    "Microsoft.XboxApp",
    "Microsoft.XboxGamingOverlay",
    "Microsoft.GamingApp",
    "Microsoft.Xbox.TCUI",
    "Microsoft.XboxIdentityProvider",
    "Microsoft.XboxGameOverlay",
    "Microsoft.XboxSpeechToTextOverlay",

    "7EE7776C.LinkedInforWindows",
    "Facebook.Facebook",
    "TikTok.TikTok",
    "Instagram.Instagram",

    "SpotifyAB.SpotifyMusic",
    "Clipchamp.Clipchamp", 
    "Microsoft.ZuneMusic",
    "Microsoft.ZuneVideo",

    "MicrosoftTeams",
    "MSTeams",
    "Microsoft.SkypeApp",

    "Microsoft.BingNews",
    "Microsoft.BingWeather",
    "Microsoft.WindowsMaps",
    "Microsoft.GetHelp",
    "Microsoft.Getstarted",
    "Microsoft.WindowsFeedbackHub",

    "microsoft.windowscommunicationsapps",
    "Microsoft.Todos",
    "Microsoft.OutlookForWindows",

    "Microsoft.MicrosoftOfficeHub",
    "Microsoft.Office.Desktop",
    "Microsoft.Office.OneNote",

    "5319275A.WhatsAppDesktop",

    "Microsoft.MixedReality.Portal",

    "Microsoft.549981C3F5F10",        # Cortana
    "Microsoft.PowerAutomateDesktop",
    "Microsoft.Windows.DevHome",
    "Microsoft.People",

    "Microsoft.YourPhone",                      # Phone Link
    "MicrosoftWindows.Client.WebExperience",    # Widgets host
    "Microsoft.Copilot",
    "Microsoft.BingSearch",                     # web search in Start
    "Microsoft.WindowsSoundRecorder",
    "Microsoft.MicrosoftJournal",
    "MicrosoftCorporationII.MicrosoftFamily",
    "MicrosoftCorporationII.QuickAssist",

    # Games (preinstalled / OEM)
    "king.com.CandyCrushSaga",
    "king.com.CandyCrushSodaSaga",
    "king.com.BubbleWitch3Saga",
    "king.com.FarmHeroesSaga",
    "A278AB0D.DisneyMagicKingdoms",
    "Microsoft.MinecraftUWP"
)

foreach ($app in $appsToRemove) {
    Remove-AppxEverywhere -Name $app
}

if ($script:AppxFailures.Count -eq 0) {
    Write-Host "Unwanted Appx application removal complete: all targets removed or absent (verified)."
}
else {
    Write-Host "Unwanted Appx removal finished with $($script:AppxFailures.Count) FAILURE(S): $($script:AppxFailures -join ', ')"
}

# Microsoft Store
$RemoveMicrosoftStore = $false
if ($RemoveMicrosoftStore) {
    Remove-AppxEverywhere -Name "Microsoft.WindowsStore"
    Write-Host "Microsoft Store removed."
}

# Widgets / News and Interests policy (Windows 10/11)
Ensure-RegistryValue -Path "HKLM:\Software\Policies\Microsoft\Dsh" -Name "AllowNewsAndInterests" -Value 0 -PropertyType DWord
Write-Host "Widjet removed"

Set-Windows11StartPolicy
Write-Host "Start menu policy processed."

# Uninstall OneDrive (non-Appx)
Write-Host "Uninstalling OneDrive..."
try {
    Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue
} catch { }

$oneDriveSetup64 = "$env:SystemRoot\System32\OneDriveSetup.exe"
$oneDriveSetup32 = "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"

if (Test-Path $oneDriveSetup64) {
    Start-Process -FilePath $oneDriveSetup64 -ArgumentList "/uninstall" -Wait -NoNewWindow
} elseif (Test-Path $oneDriveSetup32) {
    Start-Process -FilePath $oneDriveSetup32 -ArgumentList "/uninstall" -Wait -NoNewWindow
} else {
    Write-Host "OneDriveSetup.exe not found (maybe already removed)."
}

Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\OneDrive" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$env:PROGRAMDATA\Microsoft OneDrive" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$env:SYSTEMDRIVE\OneDriveTemp" -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "OneDrive removal complete."

# Disable some Windows features / consumer experiences
Ensure-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableConsumerFeatures" -Value 1 -PropertyType DWord
Ensure-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableWindowsConsumerFeatures" -Value 1 -PropertyType DWord
Ensure-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableSoftLanding" -Value 1 -PropertyType DWord
Ensure-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableWindowsSpotlightFeatures" -Value 1 -PropertyType DWord
Ensure-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo" -Name "DisabledByGroupPolicy" -Value 1 -PropertyType DWord
Ensure-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot" -Name "TurnOffWindowsCopilot" -Value 1 -PropertyType DWord
Write-Host "Windows consumer features blocked"

# Disable common startup apps for current user
$startupRegistryPaths = @(
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run",
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"
)

$startupAppPatterns = @(
    "Teams",
    "Spotify",
    "OneDrive",
    "Copilot"
)

foreach ($registryPath in $startupRegistryPaths) {
    if (Test-Path $registryPath) {
        $startupValues = Get-ItemProperty -Path $registryPath
        foreach ($property in $startupValues.PSObject.Properties) {
            if ($property.Name -in @('PSPath','PSParentPath','PSChildName','PSDrive','PSProvider')) {
                continue
            }

            foreach ($pattern in $startupAppPatterns) {
                if ($property.Name -like "*$pattern*" -or [string]$property.Value -like "*$pattern*") {
                    try {
                        Remove-ItemProperty -Path $registryPath -Name $property.Name -ErrorAction Stop
                        Write-Host "Startup entry removed: $($property.Name)"
                    }
                    catch {
                        Write-Host "Unable to remove startup entry $($property.Name): $($_.Exception.Message)"
                    }
                    break
                }
            }
        }
    }
}

$startupFolders = @(
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup",
    "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"
)

foreach ($startupFolder in $startupFolders) {
    if (Test-Path $startupFolder) {
        Get-ChildItem -Path $startupFolder -File -ErrorAction SilentlyContinue | ForEach-Object {
            foreach ($pattern in $startupAppPatterns) {
                if ($_.Name -like "*$pattern*") {
                    try {
                        Remove-Item -Path $_.FullName -Force -ErrorAction Stop
                        Write-Host "Startup shortcut removed: $($_.Name)"
                    }
                    catch {
                        Write-Host "Unable to remove startup shortcut $($_.Name): $($_.Exception.Message)"
                    }
                    break
                }
            }
        }
    }
}

Write-Host "Startup apps cleanup complete"

Ensure-RegistryValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowTaskViewButton" -Value 0 -PropertyType DWord

Ensure-RegistryValue -Path "HKCU:\Software\Microsoft\TabletTip\1.7" -Name "EnableDesktopModeAutoInvoke" -Value 1 -PropertyType DWord
Ensure-RegistryValue -Path "HKCU:\Software\Microsoft\TabletTip\1.7" -Name "TipbandDesiredVisibility" -Value 1 -PropertyType DWord

# Show touch keyboard button
Ensure-RegistryValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowTouchKeyboardButton" -Value 1 -PropertyType DWord

Write-Host "Tactile keyboard set"

Write-Step "Final cleanup"
Write-Host "Cleaning temp files..."
Remove-Item -Path "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\Windows\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Final cleanup complete"

Write-Step "Restarting Explorer"
try {
    Stop-Process -Name StartMenuExperienceHost -Force -ErrorAction SilentlyContinue
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Start-Process explorer.exe
}
catch {
    Write-Host "Explorer restart skipped: $($_.Exception.Message)"
}

Write-Step "Summary"
Write-Host "Setup done!"
Write-Host "Standby set to never"
Write-Host "Wallpaper applied"
Write-Host "Unwanted Appx removed"
Write-Host "Widget policy disabled"
Write-Host "OPS copied to C:\"
Write-Host "Touch keyboard configured"
Write-Host "OneDrive removed"
Write-Host "Windows consumer features blocked"
Write-Host "Startup apps cleaned"
Write-Host "Final cleanup done"
Write-Host "Microsoft Store removal toggle available"
Write-Host "Recommended hidden; pinned section handling applied when supported by Windows 11 build"

if ($script:AppxFailures.Count -gt 0) {
    Write-Host "WARNING: the following Appx could NOT be removed and need manual attention: $($script:AppxFailures -join ', ')"
}
