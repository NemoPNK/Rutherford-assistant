Write-Host "Language configuration - Rutherford Assistant"
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

# Max time allowed per capability operation (download from Windows Update can be slow).
# After this delay the operation is abandoned and the script continues.
$script:CapabilityAddTimeoutSeconds    = 600
$script:CapabilityRemoveTimeoutSeconds = 300

function Invoke-CapabilityWithTimeout {
    param(
        [Parameter(Mandatory=$true)][string]$CapabilityName,
        [Parameter(Mandatory=$true)][ValidateSet("Add","Remove")][string]$Operation,
        [Parameter(Mandatory=$true)][int]$TimeoutSeconds
    )

    $job = Start-Job -ScriptBlock {
        param($name, $op)
        if ($op -eq "Add") {
            Add-WindowsCapability -Online -Name $name -ErrorAction Stop | Out-Null
        }
        else {
            Remove-WindowsCapability -Online -Name $name -ErrorAction Stop | Out-Null
        }
    } -ArgumentList $CapabilityName, $Operation

    $heartbeatSeconds = 30
    $elapsed = 0

    while ($true) {
        $finished = Wait-Job -Job $job -Timeout $heartbeatSeconds
        if ($finished) { break }

        $elapsed += $heartbeatSeconds
        if ($elapsed -ge $TimeoutSeconds) {
            Stop-Job -Job $job -ErrorAction SilentlyContinue
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            Write-Host "TIMEOUT: $Operation $CapabilityName abandoned after $TimeoutSeconds s (Windows Update unreachable or download too slow). Continuing."
            return $false
        }

        Write-Host "... $Operation $CapabilityName still in progress ($elapsed s / max $TimeoutSeconds s)"
    }

    $failed = ($job.State -eq "Failed")
    $errorMessage = $null
    if ($failed) {
        $reason = $job.ChildJobs[0].JobStateInfo.Reason
        $errorMessage = if ($reason) { $reason.Message } else { "unknown error" }
    }

    Receive-Job -Job $job -ErrorAction SilentlyContinue | Out-Null
    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue

    if ($failed) {
        Write-Host "Skipping $Operation for $CapabilityName : $errorMessage"
        return $false
    }

    return $true
}

function Ensure-WindowsCapabilityPresent {
    param([Parameter(Mandatory=$true)][string]$CapabilityName)

    try {
        $capability = Get-WindowsCapability -Online -Name $CapabilityName -ErrorAction Stop
        if ($capability.State -eq "Installed") {
            Write-Host "$CapabilityName already installed."
            return $false
        }
    }
    catch {
        Write-Host "Skipping capability $CapabilityName : $($_.Exception.Message)"
        return $false
    }

    Write-Host "Installing capability: $CapabilityName (downloads from Windows Update, max $script:CapabilityAddTimeoutSeconds s)"
    if (Invoke-CapabilityWithTimeout -CapabilityName $CapabilityName -Operation Add -TimeoutSeconds $script:CapabilityAddTimeoutSeconds) {
        Write-Host "$CapabilityName installed."
        return $true
    }
    return $false
}

function Remove-WindowsCapabilityIfPresent {
    param([Parameter(Mandatory=$true)][string]$CapabilityName)

    try {
        $capability = Get-WindowsCapability -Online -Name $CapabilityName -ErrorAction Stop
        if ($capability.State -ne "Installed") {
            Write-Host "$CapabilityName already absent."
            return $false
        }
    }
    catch {
        Write-Host "Skipping removal for $CapabilityName : $($_.Exception.Message)"
        return $false
    }

    Write-Host "Removing capability: $CapabilityName (max $script:CapabilityRemoveTimeoutSeconds s)"
    if (Invoke-CapabilityWithTimeout -CapabilityName $CapabilityName -Operation Remove -TimeoutSeconds $script:CapabilityRemoveTimeoutSeconds) {
        Write-Host "$CapabilityName removed."
        return $true
    }
    return $false
}

if (-not (Test-IsAdmin)) {
    Write-Host "This script must be run as Administrator."
    exit 1
}

Write-Step "Language configuration"

# --- Target locale: single source of truth ---
# Retarget the whole machine to another locale by editing only these values.
$targetLanguage    = "en-US"
$targetInputTip    = "0409:00000409"   # US keyboard layout
$targetGeoId       = 244               # 244 = United States
$languagesToRemove = @("fr-FR")

# Tracks whether anything that requires a restart was actually changed.
$rebootNeeded = $false

# Capture the current system locale to detect a real change later.
try { $previousSystemLocale = (Get-WinSystemLocale).Name } catch { $previousSystemLocale = $null }

$hasLangModule = $null -ne (Get-Command Install-Language -ErrorAction SilentlyContinue)

# 1) Install the target display language (full pack: UI + basic + handwriting + OCR + speech)
$targetInstalled = $false
if ($hasLangModule) {
    try {
        if (Get-InstalledLanguage -Language $targetLanguage -ErrorAction SilentlyContinue) {
            Write-Host "$targetLanguage language pack already installed."
            $targetInstalled = $true
        }
        else {
            Write-Host "Installing $targetLanguage language pack via Install-Language (downloads from Windows Update)..."
            Install-Language -Language $targetLanguage -ErrorAction Stop | Out-Null
            $targetInstalled = $true
            $rebootNeeded = $true
            Write-Host "$targetLanguage language pack installed."
        }
    }
    catch {
        Write-Host "Install-Language failed for $targetLanguage : $($_.Exception.Message) - falling back to capabilities."
    }
}

if (-not $targetInstalled) {
    # Fallback for builds without the LanguagePackManagement module
    foreach ($suffix in @("Basic", "Handwriting", "OCR", "Speech")) {
        if (Ensure-WindowsCapabilityPresent -CapabilityName "Language.$suffix~~~$targetLanguage~0.0.1.0") {
            $rebootNeeded = $true
        }
    }
    try {
        $basic = Get-WindowsCapability -Online -Name "Language.Basic~~~$targetLanguage~0.0.1.0" -ErrorAction Stop
        $targetInstalled = ($basic.State -eq "Installed")
    }
    catch { $targetInstalled = $false }
}

# 2) Apply the locale to the current user and the system
try {
    $langList = New-WinUserLanguageList $targetLanguage
    $langList[0].InputMethodTips.Clear()
    $langList[0].InputMethodTips.Add($targetInputTip)
    Set-WinUserLanguageList $langList -Force

    Set-WinUILanguageOverride -Language $targetLanguage
    Set-WinDefaultInputMethodOverride -InputTip $targetInputTip
    Set-Culture $targetLanguage
    Set-WinSystemLocale $targetLanguage
    Set-WinHomeLocation -GeoId $targetGeoId

    # A system-locale change always needs a restart to take effect.
    if ($previousSystemLocale -ne $targetLanguage) {
        $rebootNeeded = $true
    }

    # Actually change the OS display language (welcome screen, system accounts) - Win11
    if (Get-Command Set-SystemPreferredUILanguage -ErrorAction SilentlyContinue) {
        $prevPreferred = $null
        try { $prevPreferred = Get-SystemPreferredUILanguage } catch { }
        Set-SystemPreferredUILanguage $targetLanguage
        # Only a real change needs a restart - keeps re-runs idempotent: re-running on an
        # already-configured machine no longer falsely demands a reboot.
        if ("$prevPreferred" -ne $targetLanguage) { $rebootNeeded = $true }
    }

    # Propagate to welcome screen and new users (Windows 11 22H2+; skipped on older builds)
    if (Get-Command Copy-UserInternationalSettingsToSystem -ErrorAction SilentlyContinue) {
        Copy-UserInternationalSettingsToSystem -WelcomeScreen $true -NewUser $true
        Write-Host "International settings copied to welcome screen and new users."
    }
    else {
        Write-Host "Copy-UserInternationalSettingsToSystem not available on this build - welcome screen / new users keep current locale."
    }

    Write-Host "Machine locale set to $targetLanguage."
}
catch {
    Write-Host "ERROR in language settings: $($_.Exception.Message) - continuing."
}

# 3) Remove unwanted languages - ONLY if the target is confirmed installed,
#    so we never strip the current language and leave the machine with none.
if ($targetInstalled) {
    foreach ($lang in $languagesToRemove) {
        if ($lang -eq $targetLanguage) { continue }

        if ($hasLangModule -and (Get-Command Uninstall-Language -ErrorAction SilentlyContinue)) {
            try {
                if (Get-InstalledLanguage -Language $lang -ErrorAction SilentlyContinue) {
                    Write-Host "Uninstalling language pack: $lang"
                    Uninstall-Language -Language $lang -ErrorAction Stop | Out-Null
                    $rebootNeeded = $true
                    Write-Host "$lang language pack removed."
                }
            }
            catch {
                Write-Host "Uninstall-Language failed for $lang : $($_.Exception.Message) - falling back to capabilities."
            }
        }

        foreach ($suffix in @("Basic", "Handwriting", "OCR", "Speech")) {
            if (Remove-WindowsCapabilityIfPresent -CapabilityName "Language.$suffix~~~$lang~0.0.1.0") {
                $rebootNeeded = $true
            }
        }
    }

    # Re-assert the target-only user language list after removals
    try {
        $finalList = New-WinUserLanguageList $targetLanguage
        $finalList[0].InputMethodTips.Clear()
        $finalList[0].InputMethodTips.Add($targetInputTip)
        Set-WinUserLanguageList $finalList -Force
        Set-WinUILanguageOverride -Language $targetLanguage
        # Re-assert the default keyboard too, so removing French cannot leave a stale default layout.
        Set-WinDefaultInputMethodOverride -InputTip $targetInputTip
    }
    catch {
        Write-Host "ERROR re-applying $targetLanguage-only language list: $($_.Exception.Message) - continuing."
    }

    Write-Host "Unwanted languages removed ($($languagesToRemove -join ', '))."
}
else {
    Write-Host "WARNING: $targetLanguage not confirmed installed - SKIPPING removal of $($languagesToRemove -join ', ') to avoid leaving the machine without a usable language."
}

# Verify the keyboard layout actually changed (read back the active input methods).
$keyboardOk = $false
try {
    $activeTips = @(Get-WinUserLanguageList | ForEach-Object { $_.InputMethodTips })
    $frenchTips = @($activeTips | Where-Object { $_ -like "*040c*" })
    if (($activeTips -contains $targetInputTip) -and ($frenchTips.Count -eq 0)) {
        $keyboardOk = $true
        Write-Host "Keyboard verified: US layout ($targetInputTip) active, no French layout left."
    }
    else {
        Write-Host "WARNING: keyboard not fully switched yet (active: $($activeTips -join ', ')). Applies fully after sign-out / reboot."
    }
}
catch {
    Write-Host "Could not verify keyboard input method: $($_.Exception.Message)"
}

Write-Step "Summary"
Write-Host "Language set to $targetLanguage"
Write-Host "Keyboard layout set to US ($targetInputTip)$(if ($keyboardOk) { ' - verified' } else { ' - applies after reboot' })"
Write-Host "Home location GeoId set to $targetGeoId"
Write-Host "Removed languages: $($languagesToRemove -join ', ')"

# A restart is required for the language change to fully apply (sign-in screen,
# system accounts, new users). The launcher reads the flag below to show
# "Restart required" in its header; it records the boot time so the launcher can
# auto-clear it once the PC has actually rebooted.
$rebootFlagPath = Join-Path "C:\ProgramData\Rutherford" "reboot-required.json"
if ($rebootNeeded) {
    try {
        $bootId = (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime.ToFileTimeUtc()
        $flagDir = Split-Path -Parent $rebootFlagPath
        if (-not (Test-Path $flagDir)) { New-Item -Path $flagDir -ItemType Directory -Force | Out-Null }
        @{ bootId = "$bootId"; reason = "language"; at = (Get-Date).ToString("s") } |
            ConvertTo-Json | Set-Content -Path $rebootFlagPath -Encoding UTF8
    }
    catch {
        Write-Host "Could not write reboot flag: $($_.Exception.Message)"
    }

    # No popup dialog: the launcher header shows the "Restart required" banner,
    # which is clearer and less intrusive than a modal window.
    Write-Host "Redemarrage necessaire pour appliquer la langue (voir le bandeau dans le header du launcher)."
}
else {
    # Nothing changed that needs a restart - clear any stale flag.
    Remove-Item -Path $rebootFlagPath -Force -ErrorAction SilentlyContinue
    Write-Host "No restart needed - the target locale was already applied."
}

Write-Host "Language configuration done!"
