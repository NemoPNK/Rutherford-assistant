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

function Ensure-WindowsCapabilityPresent {
    param([Parameter(Mandatory=$true)][string]$CapabilityName)

    try {
        $capability = Get-WindowsCapability -Online -Name $CapabilityName -ErrorAction Stop
        if ($capability.State -eq "Installed") {
            Write-Host "$CapabilityName already installed."
            return
        }

        Write-Host "Installing capability: $CapabilityName"
        Add-WindowsCapability -Online -Name $CapabilityName -ErrorAction Stop | Out-Null
        Write-Host "$CapabilityName installed."
    }
    catch {
        Write-Host "Skipping capability $CapabilityName : $($_.Exception.Message)"
    }
}

function Remove-WindowsCapabilityIfPresent {
    param([Parameter(Mandatory=$true)][string]$CapabilityName)

    try {
        $capability = Get-WindowsCapability -Online -Name $CapabilityName -ErrorAction Stop
        if ($capability.State -ne "Installed") {
            Write-Host "$CapabilityName already absent."
            return
        }

        Write-Host "Removing capability: $CapabilityName"
        Remove-WindowsCapability -Online -Name $CapabilityName -ErrorAction Stop | Out-Null
        Write-Host "$CapabilityName removed."
    }
    catch {
        Write-Host "Skipping removal for $CapabilityName : $($_.Exception.Message)"
    }
}

function Remove-AppxEverywhere {
    param(
        [Parameter(Mandatory=$true)][string]$Name
    )

    Write-Host "Removing Appx: $Name ..."

    try {
        Get-AppxPackage -Name $Name -AllUsers -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
    } catch { }

    try {
        Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -eq $Name } | ForEach-Object {
            Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -ErrorAction SilentlyContinue | Out-Null
        }
    } catch { }
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
    $startPolicyManagerPath = "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Start"
    $startExplorerPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer"
    $startLayoutDirectory = "C:\ProgramData\Rutherford"
    $startLayoutFile = Join-Path $startLayoutDirectory "StartPins.json"
    $layoutJson = '{"pinnedList":[]}'

    if (-not (Test-Path $startLayoutDirectory)) {
        New-Item -Path $startLayoutDirectory -ItemType Directory -Force | Out-Null
    }

    Set-Content -Path $startLayoutFile -Value $layoutJson -Encoding UTF8

    if (-not (Test-Path $startPolicyManagerPath)) {
        New-Item -Path $startPolicyManagerPath -Force | Out-Null
    }

    if (-not (Test-Path $startExplorerPolicyPath)) {
        New-Item -Path $startExplorerPolicyPath -Force | Out-Null
    }

    New-ItemProperty -Path $startExplorerPolicyPath -Name "HideRecommendedSection" -Value 1 -PropertyType DWord -Force | Out-Null

    if ($buildNumber -ge 26100) {
        New-ItemProperty -Path $startPolicyManagerPath -Name "ConfigureStartPins" -Value $layoutJson -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $startExplorerPolicyPath -Name "ConfigureStartPins" -Value $startLayoutFile -PropertyType String -Force | Out-Null
        Write-Host "Windows 11 Start policy applied: Recommended hidden and Pinned section emptied."
    }
    else {
        Write-Host "Recommended section hidden. Empty pinned layout skipped because ConfigureStartPins is only reliably supported on newer Windows 11 builds."
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

if (Test-Path $wallpaperPath) {

    Add-Type @"
using System.Runtime.InteropServices;
public class Wallpaper {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@

    [Wallpaper]::SystemParametersInfo(20, 0, $wallpaperPath, 0x1 -bor 0x2)

    Write-Host "Wallpaper set"
}
else {
    Write-Host "cant found wallpaper $wallpaperPath"
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
    "Microsoft.News",
    "Microsoft.Weather",
    "Microsoft.WindowsMaps",
    "Microsoft.GetHelp",
    "Microsoft.Getstarted",
    "Microsoft.WindowsFeedbackHub",

    "microsoft.windowscommunicationsapps",
    "Microsoft.Todos",
    "Microsoft.OutlookForWindows",

    "Microsoft.MicrosoftOfficeHub",
    "Microsoft.Office.Desktop",
    "Microsoft.Office.OneNote"
)

foreach ($app in $appsToRemove) {
    Remove-AppxEverywhere -Name $app
}

Write-Host "Unwanted Appx application removal complete."

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

Write-Step "Language configuration"

$capabilitiesToInstall = @(
    "Language.Basic~~~en-US~0.0.1.0",
    "Language.Handwriting~~~en-US~0.0.1.0",
    "Language.OCR~~~en-US~0.0.1.0",
    "Language.Speech~~~en-US~0.0.1.0"
)

foreach ($capabilityName in $capabilitiesToInstall) {
    Ensure-WindowsCapabilityPresent -CapabilityName $capabilityName
}

Set-WinUILanguageOverride -Language en-US
Set-WinDefaultInputMethodOverride -InputTip "0409:00000409"
Set-WinUserLanguageList -LanguageList en-US -Force

Set-Culture en-US
Set-WinSystemLocale en-US
Set-WinHomeLocation -GeoId 244

$LangList = New-WinUserLanguageList en-US
$LangList[0].InputMethodTips.Clear()
$LangList[0].InputMethodTips.Add("0409:00000409")
Set-WinUserLanguageList $LangList -Force

Copy-UserInternationalSettingsToSystem -WelcomeScreen $true -NewUser $true

Write-Host "Machine is now US"

# Remove French language packs and keep English only
$englishOnlyList = New-WinUserLanguageList en-US
$englishOnlyList[0].InputMethodTips.Clear()
$englishOnlyList[0].InputMethodTips.Add("0409:00000409")
Set-WinUserLanguageList $englishOnlyList -Force

$capabilitiesToRemove = @(
    "Language.Basic~~~fr-FR~0.0.1.0",
    "Language.Handwriting~~~fr-FR~0.0.1.0",
    "Language.OCR~~~fr-FR~0.0.1.0",
    "Language.Speech~~~fr-FR~0.0.1.0"
)

foreach ($capabilityName in $capabilitiesToRemove) {
    Remove-WindowsCapabilityIfPresent -CapabilityName $capabilityName
}

Set-WinUILanguageOverride -Language en-US

Write-Host "French removed"

Ensure-RegistryValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowTaskViewButton" -Value 0 -PropertyType DWord

Ensure-RegistryValue -Path "HKCU:\Software\Microsoft\TabletTip\1.7" -Name "EnableDesktopModeAutoInvoke" -Value 1 -PropertyType DWord
Ensure-RegistryValue -Path "HKCU:\Software\Microsoft\TabletTip\1.7" -Name "TipbandDesiredVisibility" -Value 1 -PropertyType DWord

# Show touch keyboard button
Ensure-RegistryValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowTouchKeyboardButton" -Value 1 -PropertyType DWord

Write-Host "Tactile keyboard set"

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
Write-Host "Language set to en-US"
Write-Host "OPS copied to C:\"
Write-Host "Touch keyboard configured"
Write-Host "OneDrive removed"
Write-Host "Microsoft Store removal toggle available"
Write-Host "Recommended hidden; pinned section handling applied when supported by Windows 11 build"