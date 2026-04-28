# ----------------------------------------------------------------------------
# Bootstrap crash log - MUST be the very first code so we capture even
# startup errors. Tries multiple paths and picks the first writable one.
# ----------------------------------------------------------------------------

$script:CrashLogPath = $null

function Try-WriteLog {
    param([string]$Path)
    try {
        $dir = Split-Path -Parent $Path
        if ($dir -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -Path $dir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        [System.IO.File]::AppendAllText($Path, "")
        return $true
    } catch { return $false }
}

# Build candidate list: next to EXE first (most findable), then standard places
$_logCandidates = @()
try {
    $_exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    if ($_exePath) { $_logCandidates += (Join-Path (Split-Path -Parent $_exePath) "RutherfordAssistant.log") }
} catch { }
if ($PSCommandPath) {
    $_logCandidates += (Join-Path (Split-Path -Parent $PSCommandPath) "RutherfordAssistant.log")
}
$_logCandidates += "C:\ProgramData\Rutherford\launcher.log"
$_logCandidates += (Join-Path $env:TEMP "RutherfordAssistant.log")
$_logCandidates += (Join-Path $env:USERPROFILE "RutherfordAssistant.log")

foreach ($_cand in $_logCandidates) {
    if ([string]::IsNullOrWhiteSpace($_cand)) { continue }
    if (Try-WriteLog $_cand) { $script:CrashLogPath = $_cand; break }
}

function Write-CrashLog {
    param([string]$Message)
    if (-not $script:CrashLogPath) { return }
    try {
        $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        [System.IO.File]::AppendAllText($script:CrashLogPath, ("[$stamp] " + $Message + "`r`n"))
    } catch { }
}

Write-CrashLog "===== Launcher boot ====="
Write-CrashLog ("CrashLogPath = " + $script:CrashLogPath)

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase
Write-CrashLog "WPF assemblies loaded"

# ----------------------------------------------------------------------------
# Path resolution
# Works in both raw .ps1 and ps2exe-compiled .exe contexts.
# In ps2exe, $PSCommandPath / $MyInvocation are usually empty, so we fall back
# to the running process MainModule path (which IS the EXE itself).
# ----------------------------------------------------------------------------

function Get-LauncherSelfPath {
    if ($PSCommandPath) { return $PSCommandPath }
    if ($MyInvocation.MyCommand.Path) { return $MyInvocation.MyCommand.Path }
    try {
        $main = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        if ($main -and (Test-Path -LiteralPath $main)) { return $main }
    }
    catch { }
    return $null
}

$script:LauncherPath = Get-LauncherSelfPath

if ($script:LauncherPath) {
    $script:SelfRoot = Split-Path -Parent $script:LauncherPath
}
else {
    # Last resort: current working directory (when launched from Explorer this
    # is usually where the EXE sits)
    $script:SelfRoot = (Get-Location).Path
}

if (-not $script:SelfRoot) {
    $script:SelfRoot = [System.IO.Directory]::GetCurrentDirectory()
}

$script:AssetsRoot = $null
$script:AppRoot    = $null

# Candidate paths to search for assets/ - first hit wins
$assetCandidates = @(
    (Join-Path $script:SelfRoot "assets"),
    $script:SelfRoot,
    (Join-Path $script:SelfRoot "..\assets"),
    (Join-Path ([System.IO.Directory]::GetCurrentDirectory()) "assets")
)

foreach ($candidate in $assetCandidates) {
    if (-not $candidate) { continue }
    if (-not (Test-Path -LiteralPath $candidate)) { continue }
    $leaf = Split-Path -Leaf $candidate
    if ($leaf -eq "assets") {
        $script:AssetsRoot = (Resolve-Path -LiteralPath $candidate).Path
        $script:AppRoot    = Split-Path -Parent $script:AssetsRoot
        break
    }
    # If folder exists but is not "assets", check if it CONTAINS an "assets" subfolder
    $sub = Join-Path $candidate "assets"
    if (Test-Path -LiteralPath $sub) {
        $script:AppRoot    = (Resolve-Path -LiteralPath $candidate).Path
        $script:AssetsRoot = (Resolve-Path -LiteralPath $sub).Path
        break
    }
}

if (-not $script:AssetsRoot) {
    # Final fallback - keep launcher running so the operator sees a clear message
    $script:AppRoot    = $script:SelfRoot
    $script:AssetsRoot = Join-Path $script:SelfRoot "assets"
}

$script:ChecksRoot         = Join-Path $script:AssetsRoot "checks"
$script:ReportsRoot        = Join-Path $script:AppRoot "reports"
$script:NetworkProfilesPath = Join-Path $script:AssetsRoot "config\network-profiles.json"
$script:StateRoot          = "C:\ProgramData\Rutherford"
$script:StateFilePath      = Join-Path $script:StateRoot "launcher-state.json"

# ----------------------------------------------------------------------------
# ColorLoop Design System palette
# Source: ColorLoop logo colors extracted from the .fig design system
# ----------------------------------------------------------------------------

$script:Palette = @{
    Red         = "#FD0902"
    Yellow      = "#FDD800"
    Green       = "#63B02F"
    Blue        = "#1DB6FF"
    Pink        = "#FCBBEB"
    White       = "#FFFFFF"

    # Surfaces / text
    InkPrimary  = "#0A0A0A"
    InkSoft     = "#52525B"
    InkMuted    = "#A1A1AA"
    Surface     = "#FFFFFF"
    SurfaceDim  = "#F4F4F5"
    Border      = "#E5E7EB"
    HeroBg      = "#0F0F10"
    HeroBorder  = "#232326"
    HeroText    = "#F4F4F5"
    HeroSoft    = "#B3B3BC"

    # Status badges
    OkBg        = "#DCFCE7"; OkFg = "#166534"
    ErrorBg     = "#FEE2E2"; ErrorFg = "#B91C1C"
    WarnBg      = "#FEF3C7"; WarnFg = "#92400E"
    NeutralBg   = "#E5E7EB"; NeutralFg = "#374151"
}

function Get-ColorLoopAccent {
    param([string]$Name)
    switch ($Name) {
        "red"     { return @{ Bg = $script:Palette.Red;    Fg = "#FFFFFF"; Hover = "#E10000" } }
        "yellow"  { return @{ Bg = $script:Palette.Yellow; Fg = "#0A0A0A"; Hover = "#E6C400" } }
        "green"   { return @{ Bg = $script:Palette.Green;  Fg = "#FFFFFF"; Hover = "#558E29" } }
        "blue"    { return @{ Bg = $script:Palette.Blue;   Fg = "#FFFFFF"; Hover = "#0098DA" } }
        "pink"    { return @{ Bg = $script:Palette.Pink;   Fg = "#0A0A0A"; Hover = "#F4A0DA" } }
        "white"   { return @{ Bg = "#FFFFFF";              Fg = "#0A0A0A"; Hover = "#F4F4F5" } }
        "dark"    { return @{ Bg = "#0A0A0A";              Fg = "#FFFFFF"; Hover = "#1F1F1F" } }
        default   { return @{ Bg = $script:Palette.Green;  Fg = "#FFFFFF"; Hover = "#558E29" } }
    }
}

# ----------------------------------------------------------------------------
# Globals
# ----------------------------------------------------------------------------

$script:IsBusy           = $false
$script:CurrentProcess   = $null
$script:CurrentTask      = $null
$script:CurrentTaskKey   = $null
$script:LastReportPath   = $null
$script:RunStartedAt     = $null
$script:CurrentLogLines  = New-Object System.Collections.Generic.List[string]
$script:LogItems         = New-Object 'System.Collections.ObjectModel.ObservableCollection[string]'
$script:Tasks            = New-Object 'System.Collections.Generic.List[hashtable]'
$script:TasksByKey       = @{}
$script:ButtonKeyMap     = @{}   # button ref -> task key (fallback if Tag fails under ps2exe)

# File-based output handling. The script's stdout / stderr are redirected
# to temp files by the OS itself (via Start-Process -RedirectStandardOutput).
# The UI thread polls these files via a DispatcherTimer. This avoids ALL
# Process events on background threads, which we suspect crash ps2exe.
$script:CurrentOutputFile  = $null
$script:CurrentErrorFile   = $null
$script:OutputFilePosition = [int64]0
$script:ErrorFilePosition  = [int64]0
$script:CompletionHandled  = $false
$script:AuditChecks      = @()
$script:State            = @{
    computerName = $env:COMPUTERNAME
    lastUpdated  = $null
    scripts      = @{}
}

# ----------------------------------------------------------------------------
# Admin elevation
# ----------------------------------------------------------------------------

function Test-IsAdmin {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Ensure-Elevated {
    if (Test-IsAdmin) { return }

    $argumentList = "-NoProfile -ExecutionPolicy Bypass -STA -File `"$script:LauncherPath`""
    Start-Process -FilePath "powershell.exe" -ArgumentList $argumentList -Verb RunAs | Out-Null
    exit
}

# ----------------------------------------------------------------------------
# Persistent state
# ----------------------------------------------------------------------------

function Load-State {
    try {
        if (Test-Path $script:StateFilePath) {
            $raw = Get-Content -Path $script:StateFilePath -Raw -Encoding UTF8
            if ([string]::IsNullOrWhiteSpace($raw)) { return }
            $loaded = $raw | ConvertFrom-Json -ErrorAction Stop
            if ($loaded -and $loaded.computerName -eq $env:COMPUTERNAME) {
                $scripts = @{}
                if ($loaded.scripts) {
                    foreach ($prop in $loaded.scripts.PSObject.Properties) {
                        $scripts[$prop.Name] = @{
                            status      = $prop.Value.status
                            completedAt = $prop.Value.completedAt
                            exitCode    = $prop.Value.exitCode
                        }
                    }
                }
                $script:State = @{
                    computerName = $env:COMPUTERNAME
                    lastUpdated  = $loaded.lastUpdated
                    scripts      = $scripts
                }
            }
        }
    }
    catch {
        # Corrupted state file: silently keep defaults
    }
}

function Save-State {
    try {
        if (-not (Test-Path $script:StateRoot)) {
            New-Item -Path $script:StateRoot -ItemType Directory -Force | Out-Null
        }
        $script:State.lastUpdated = (Get-Date).ToString("s")
        $json = $script:State | ConvertTo-Json -Depth 6
        Set-Content -Path $script:StateFilePath -Value $json -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        # Cannot persist (e.g. read-only) - keep going, never crash the launcher
    }
}

function Update-ScriptState {
    param(
        [string]$Key,
        [string]$Status,
        [int]$ExitCode = -1
    )

    if (-not $script:State.scripts) { $script:State.scripts = @{} }
    $script:State.scripts[$Key] = @{
        status      = $Status
        completedAt = (Get-Date).ToString("s")
        exitCode    = $ExitCode
    }
    Save-State
}

# ----------------------------------------------------------------------------
# Task discovery (auto-discovery via *.manifest.json)
# ----------------------------------------------------------------------------

function Discover-Tasks {
    $tasks = New-Object 'System.Collections.Generic.List[hashtable]'

    if ([string]::IsNullOrWhiteSpace($script:AssetsRoot)) { return $tasks }
    if (-not (Test-Path -LiteralPath $script:AssetsRoot)) { return $tasks }

    $manifests = Get-ChildItem -Path $script:AssetsRoot -Filter "*.manifest.json" -File -ErrorAction SilentlyContinue
    foreach ($manifestFile in $manifests) {
        try {
            $raw = Get-Content -Path $manifestFile.FullName -Raw -Encoding UTF8
            $manifest = $raw | ConvertFrom-Json
            if (-not $manifest -or -not $manifest.key) { continue }

            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($manifestFile.Name)
            if ($baseName.EndsWith(".manifest")) {
                $baseName = $baseName.Substring(0, $baseName.Length - ".manifest".Length)
            }

            $scriptPath = Join-Path $manifestFile.DirectoryName ("$baseName.ps1")
            if (-not (Test-Path $scriptPath)) { continue }

            $task = @{
                Key            = [string]$manifest.key
                Label          = if ($manifest.label) { [string]$manifest.label } else { "Run $baseName" }
                Description    = if ($manifest.description) { [string]$manifest.description } else { "" }
                Order          = if ($null -ne $manifest.order) { [int]$manifest.order } else { 999 }
                Primary        = if ($null -ne $manifest.primary) { [bool]$manifest.primary } else { $false }
                AuditAfterRun  = if ($null -ne $manifest.auditAfterRun) { [bool]$manifest.auditAfterRun } else { $false }
                ScriptPath     = $scriptPath
                ManifestPath   = $manifestFile.FullName
                Button         = $null
                StatusBorder   = $null
                StatusText     = $null
            }

            $tasks.Add($task) | Out-Null
        }
        catch {
            # Skip bad manifest, never crash
        }
    }

    # Sort by Order then by Label
    $sorted = $tasks | Sort-Object @{Expression = { $_.Order }}, @{Expression = { $_.Label }}
    $result = New-Object 'System.Collections.Generic.List[hashtable]'
    foreach ($t in $sorted) { $result.Add($t) | Out-Null }
    return $result
}

# ----------------------------------------------------------------------------
# Audit check discovery
# ----------------------------------------------------------------------------

function Discover-AuditChecks {
    $checks = @()
    $script:DiscoveryDiagnostics = @()

    if ([string]::IsNullOrWhiteSpace($script:ChecksRoot)) {
        $script:DiscoveryDiagnostics += "Discover-AuditChecks: ChecksRoot is empty"
        return $checks
    }
    if (-not (Test-Path -LiteralPath $script:ChecksRoot)) {
        $script:DiscoveryDiagnostics += "Discover-AuditChecks: ChecksRoot not found: $script:ChecksRoot"
        return $checks
    }

    $files = @(Get-ChildItem -Path $script:ChecksRoot -Filter "*.check.ps1" -File -ErrorAction SilentlyContinue)
    $script:DiscoveryDiagnostics += "Discover-AuditChecks: $($files.Count) file(s) in $script:ChecksRoot"

    foreach ($file in ($files | Sort-Object Name)) {
        try {
            $content = $null
            try {
                $content = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8 -ErrorAction Stop
            }
            catch {
                $script:DiscoveryDiagnostics += "Read failed for $($file.Name): $($_.Exception.Message)"
                continue
            }

            if ([string]::IsNullOrWhiteSpace($content)) {
                $script:DiscoveryDiagnostics += "Empty: $($file.Name)"
                continue
            }

            # Use [scriptblock]::Create + & to capture the hashtable output reliably
            # in both .ps1 and ps2exe-compiled .exe contexts.
            $sb = [scriptblock]::Create($content)
            $check = & $sb

            if (-not ($check -is [hashtable])) {
                $typeName = if ($null -eq $check) { "<null>" } else { $check.GetType().FullName }
                $script:DiscoveryDiagnostics += "Not a hashtable: $($file.Name) (got $typeName)"
                continue
            }
            if (-not $check.ContainsKey('Test')) {
                $script:DiscoveryDiagnostics += "No 'Test' key: $($file.Name)"
                continue
            }

            if (-not $check.ContainsKey('Order'))    { $check.Order    = 999 }
            if (-not $check.ContainsKey('Label'))    { $check.Label    = $file.BaseName }
            if (-not $check.ContainsKey('Category')) { $check.Category = "General" }
            $check.SourcePath = $file.FullName
            $checks += $check
        }
        catch {
            $script:DiscoveryDiagnostics += "Error in $($file.Name): $($_.Exception.Message)"
        }
    }

    $script:DiscoveryDiagnostics += "Discover-AuditChecks: $($checks.Count) check(s) loaded"
    return @($checks | Sort-Object @{Expression = { $_.Order }}, @{Expression = { $_.Label }})
}

# ----------------------------------------------------------------------------
# Network helpers (unchanged)
# ----------------------------------------------------------------------------

function Read-NetworkProfiles {
    if ([string]::IsNullOrWhiteSpace($script:NetworkProfilesPath)) {
        throw "Network config path is empty (assets folder was not found)."
    }
    if (-not (Test-Path -LiteralPath $script:NetworkProfilesPath)) {
        throw "Missing network config: $script:NetworkProfilesPath"
    }

    $raw = Get-Content -Path $script:NetworkProfilesPath -Raw -Encoding UTF8
    $profiles = $raw | ConvertFrom-Json
    if ($null -eq $profiles) {
        throw "Network config is empty."
    }

    return @($profiles)
}

function Get-AdapterByNameList {
    param([string[]]$Names)

    foreach ($name in $Names) {
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $adapter = Get-NetAdapter -Name $name -ErrorAction SilentlyContinue
        if ($adapter) { return $adapter }
    }

    return $null
}

function Convert-PrefixLengthToMask {
    param([int]$PrefixLength)

    if ($PrefixLength -lt 0 -or $PrefixLength -gt 32) { return "" }

    $bits = ("1" * $PrefixLength).PadRight(32, "0")
    $octets = for ($index = 0; $index -lt 4; $index++) {
        [Convert]::ToInt32($bits.Substring($index * 8, 8), 2)
    }

    return ($octets -join ".")
}

function Escape-Html {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Get-LineSeverity {
    param([string]$Line)

    if ($Line -match "(?i)\b(ERROR|Failed|Exception|not found|can't|cant found)\b") { return "error" }
    if ($Line -match "(?i)\b(completed|configured|processed|copied|added|installed|removed|set|already|ready|success)\b") { return "success" }
    return "neutral"
}

function Get-StepSummaries {
    param(
        [string[]]$Lines,
        [int]$ExitCode
    )

    $steps = New-Object System.Collections.Generic.List[object]
    $currentStep = $null

    foreach ($line in $Lines) {
        $cleanLine = ($line -replace "^\[[^\]]+\]\s*", "").Trim()

        if ($cleanLine -match "^===\s*(.+?)\s*===$") {
            if ($currentStep) {
                $steps.Add([pscustomobject]$currentStep) | Out-Null
            }
            $currentStep = @{
                Title   = $matches[1]
                Status  = "success"
                Details = New-Object System.Collections.Generic.List[string]
            }
            continue
        }

        if (-not $currentStep) { continue }

        if ((Get-LineSeverity $cleanLine) -eq "error") {
            $currentStep.Status = "error"
        }

        if (-not [string]::IsNullOrWhiteSpace($cleanLine)) {
            $currentStep.Details.Add($cleanLine) | Out-Null
        }
    }

    if ($currentStep) {
        $steps.Add([pscustomobject]$currentStep) | Out-Null
    }

    if ($steps.Count -eq 0) {
        $fallbackStatus = if ($ExitCode -eq 0) { "success" } else { "error" }
        $steps.Add([pscustomobject]@{
            Title   = "Script execution"
            Status  = $fallbackStatus
            Details = @("Exit code: $ExitCode")
        }) | Out-Null
    }

    $steps.Add([pscustomobject]@{
        Title   = "Final result"
        Status  = if ($ExitCode -eq 0) { "success" } else { "error" }
        Details = @("Exit code: $ExitCode")
    }) | Out-Null

    return @($steps)
}

function Get-NetworkSnapshot {
    param($Profile)

    $adapter = Get-AdapterByNameList -Names (@($Profile.newName) + @($Profile.possibleCurrentNames))

    if (-not $adapter) {
        return [pscustomobject]@{
            Name         = $Profile.newName
            Mode         = $Profile.mode
            ExpectedIp   = $Profile.ipv4
            ExpectedMask = $Profile.subnetMask
            ActualName   = "Missing"
            ActualIp     = ""
            ActualMask   = ""
            Status       = "error"
            StatusText   = "Missing"
            Detail       = "Adapter not found"
        }
    }

    $ipv4Info = Get-NetIPAddress -InterfaceAlias $adapter.Name -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.IPAddress -notlike "169.254.*" } |
        Select-Object -First 1

    $actualIp = if ($ipv4Info) { $ipv4Info.IPAddress } else { "" }
    $actualMask = if ($ipv4Info) { Convert-PrefixLengthToMask -PrefixLength $ipv4Info.PrefixLength } else { "" }
    $ipInterface = Get-NetIPInterface -InterfaceAlias $adapter.Name -AddressFamily IPv4 -ErrorAction SilentlyContinue
    $dhcpEnabled = $ipInterface -and $ipInterface.Dhcp -eq "Enabled"

    $nameOk = $adapter.Name -eq $Profile.newName
    if ($Profile.mode -eq "dhcp") {
        $statusOk = $nameOk -and $dhcpEnabled
        $detail = if ($statusOk) { "Name and DHCP are correct" } else { "Expected DHCP on adapter $($Profile.newName)" }
    }
    else {
        $statusOk = $nameOk -and $actualIp -eq $Profile.ipv4 -and $actualMask -eq $Profile.subnetMask
        $detail = if ($statusOk) {
            "Name, IP and mask are correct"
        }
        else {
            "Expected $($Profile.ipv4) / $($Profile.subnetMask)"
        }
    }

    return [pscustomobject]@{
        Name         = $Profile.newName
        Mode         = $Profile.mode
        ExpectedIp   = if ($Profile.mode -eq "dhcp") { "DHCP" } else { $Profile.ipv4 }
        ExpectedMask = if ($Profile.mode -eq "dhcp") { "Auto" } else { $Profile.subnetMask }
        ActualName   = $adapter.Name
        ActualIp     = if ($actualIp) { $actualIp } else { "N/A" }
        ActualMask   = if ($actualMask) { $actualMask } else { "N/A" }
        Status       = if ($statusOk) { "success" } else { "error" }
        StatusText   = if ($statusOk) { "OK" } else { "Mismatch" }
        Detail       = $detail
    }
}

function Get-NetworkSnapshots {
    return @(Read-NetworkProfiles | ForEach-Object { Get-NetworkSnapshot -Profile $_ })
}

# ----------------------------------------------------------------------------
# Boot - elevation + discovery
# ----------------------------------------------------------------------------

Ensure-Elevated
Load-State
$script:Tasks = Discover-Tasks
foreach ($t in $script:Tasks) { $script:TasksByKey[$t.Key] = $t }
$script:AuditChecks = Discover-AuditChecks

# ----------------------------------------------------------------------------
# XAML
# ----------------------------------------------------------------------------

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Rutherford Assistant"
        Width="1280"
        Height="900"
        MinWidth="1080"
        MinHeight="780"
        WindowStartupLocation="CenterScreen"
        Background="#050505"
        Foreground="#F4F4F5"
        FontFamily="Segoe UI">
  <Window.Resources>
    <Style x:Key="RoundedButton" TargetType="Button">
      <Setter Property="FocusVisualStyle" Value="{x:Null}" />
      <Setter Property="Cursor" Value="Hand" />
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="ButtonBorder"
                    CornerRadius="18"
                    Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    SnapsToDevicePixels="True">
              <ContentPresenter HorizontalAlignment="Center"
                                VerticalAlignment="Center"
                                TextElement.Foreground="{TemplateBinding Foreground}"
                                TextElement.FontWeight="{TemplateBinding FontWeight}"
                                TextElement.FontSize="{TemplateBinding FontSize}" />
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.5" />
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="RoundedSecondaryButton" TargetType="Button" BasedOn="{StaticResource RoundedButton}">
      <Setter Property="Background" Value="#FFFFFF" />
      <Setter Property="Foreground" Value="#0A0A0A" />
      <Setter Property="BorderBrush" Value="#E5E7EB" />
      <Setter Property="BorderThickness" Value="1" />
      <Setter Property="FontWeight" Value="Bold" />
      <Setter Property="Height" Value="42" />
    </Style>
  </Window.Resources>
  <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
    <Grid Margin="20">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto" />
        <RowDefinition Height="18" />
        <RowDefinition Height="Auto" />
        <RowDefinition Height="18" />
        <RowDefinition Height="*" />
      </Grid.RowDefinitions>

      <Border Grid.Row="0"
              Background="#0F0F10"
              BorderBrush="#232326"
              BorderThickness="1"
              CornerRadius="22"
              Padding="22">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*" />
            <ColumnDefinition Width="330" />
          </Grid.ColumnDefinitions>

          <StackPanel Grid.Column="0">
            <StackPanel Orientation="Horizontal">
              <Ellipse Width="16" Height="16" Fill="#FD0902" Margin="0,0,10,0" />
              <Ellipse Width="16" Height="16" Fill="#FDD800" Margin="0,0,10,0" />
              <Ellipse Width="16" Height="16" Fill="#63B02F" Margin="0,0,10,0" />
              <Ellipse Width="16" Height="16" Fill="#1DB6FF" Margin="0,0,10,0" />
              <Ellipse Width="16" Height="16" Fill="#FCBBEB" />
            </StackPanel>
            <TextBlock Margin="0,16,0,0"
                       Text="Rutherford Assistant"
                       FontSize="32"
                       FontWeight="Bold" />
            <TextBlock Name="HeroSubtitle"
                       Margin="0,10,0,0"
                       Foreground="#B3B3BC"
                       FontSize="14"
                       TextWrapping="Wrap"
                       Text="Portable Windows launcher. The window stays open while scripts run." />
          </StackPanel>

          <Border Grid.Column="1"
                  Background="#151517"
                  BorderBrush="#27272A"
                  BorderThickness="1"
                  CornerRadius="18"
                  Padding="16">
            <StackPanel>
              <TextBlock Text="Operator Flow"
                         FontWeight="Bold"
                         Foreground="#F9F9FA" />
              <TextBlock Margin="0,10,0,0"
                         Foreground="#B3B3BC"
                         TextWrapping="Wrap"
                         Text="1. Open the launcher.&#x0a;2. Click an action button.&#x0a;3. Watch live logs.&#x0a;4. Open the HTML report when finished." />
            </StackPanel>
          </Border>
        </Grid>
      </Border>

      <Grid Grid.Row="2">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="380" />
          <ColumnDefinition Width="18" />
          <ColumnDefinition Width="*" />
        </Grid.ColumnDefinitions>

        <Border Grid.Column="0"
                Background="#FFFFFF"
                BorderBrush="#E5E7EB"
                BorderThickness="1"
                CornerRadius="22"
                Padding="20">
          <StackPanel>
            <TextBlock Text="Actions"
                       Foreground="#111111"
                       FontSize="20"
                       FontWeight="Bold" />
            <TextBlock Name="ActionsHelpText"
                       Margin="0,6,0,0"
                       Foreground="#52525B"
                       TextWrapping="Wrap"
                       Text="One button per script discovered in assets." />

            <StackPanel Name="ActionsPanel" Margin="0,16,0,0" />

            <Button Name="OpenReportButton"
                    Style="{StaticResource RoundedSecondaryButton}"
                    Margin="0,22,0,0"
                    IsEnabled="False"
                    Content="Open Last Report" />

            <Button Name="RefreshNetworkButton"
                    Style="{StaticResource RoundedSecondaryButton}"
                    Margin="0,10,0,0"
                    Content="Refresh Network Status" />

            <Button Name="RefreshAuditButton"
                    Style="{StaticResource RoundedSecondaryButton}"
                    Margin="0,10,0,0"
                    Content="Refresh Setup Audit" />

            <Button Name="ClearLogsButton"
                    Style="{StaticResource RoundedSecondaryButton}"
                    Margin="0,10,0,0"
                    Content="Clear Logs" />

            <TextBlock Margin="0,24,0,0"
                       Foreground="#111111"
                       FontWeight="Bold"
                       Text="Current Report" />
            <TextBlock Name="ReportSummaryText"
                       Margin="0,8,0,0"
                       Foreground="#52525B"
                       TextWrapping="Wrap"
                       Text="No report yet." />

            <TextBlock Margin="0,16,0,0"
                       Foreground="#111111"
                       FontWeight="Bold"
                       Text="State File" />
            <TextBlock Name="StateFileText"
                       Margin="0,6,0,0"
                       Foreground="#52525B"
                       FontSize="11"
                       TextWrapping="Wrap"
                       Text="" />
          </StackPanel>
        </Border>

        <Grid Grid.Column="2">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="18" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="18" />
            <RowDefinition Height="380" />
          </Grid.RowDefinitions>

          <Border Grid.Row="0"
                  Background="#0F0F10"
                  BorderBrush="#232326"
                  BorderThickness="1"
                  CornerRadius="22"
                  Padding="20">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*" />
                <ColumnDefinition Width="260" />
              </Grid.ColumnDefinitions>

              <StackPanel Grid.Column="0">
                <TextBlock Text="Execution"
                           FontSize="20"
                           FontWeight="Bold" />
                <TextBlock Name="CurrentTaskText"
                           Margin="0,10,0,0"
                           FontSize="24"
                           FontWeight="SemiBold"
                           Text="Ready" />
                <TextBlock Name="CurrentStatusText"
                           Margin="0,8,0,0"
                           Foreground="#B3B3BC"
                           Text="Waiting for action." />
              </StackPanel>

              <Border Grid.Column="1"
                      Background="#151517"
                      BorderBrush="#27272A"
                      BorderThickness="1"
                      CornerRadius="18"
                      Padding="14">
                <StackPanel>
                  <TextBlock Text="Computer"
                             FontWeight="Bold" />
                  <TextBlock Name="ComputerNameText"
                             Margin="0,8,0,0"
                             Foreground="#B3B3BC"
                             TextWrapping="Wrap"
                             Text="" />
                  <TextBlock Name="LastUpdatedText"
                             Margin="0,8,0,0"
                             Foreground="#71717A"
                             FontSize="11"
                             TextWrapping="Wrap"
                             Text="" />
                </StackPanel>
              </Border>
            </Grid>
          </Border>

          <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*" />
              <ColumnDefinition Width="18" />
              <ColumnDefinition Width="*" />
            </Grid.ColumnDefinitions>

            <Border Grid.Column="0"
                    Background="#FFFFFF"
                    BorderBrush="#E5E7EB"
                    BorderThickness="1"
                    CornerRadius="22"
                    Padding="20">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto" />
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0"
                             Text="Network Cards"
                             Foreground="#111111"
                             FontSize="20"
                             FontWeight="Bold" />
                  <TextBlock Grid.Column="1"
                             Name="NetworkSummaryText"
                             VerticalAlignment="Center"
                             Foreground="#52525B"
                             Text="" />
                </Grid>
                <TextBlock Margin="0,8,0,0"
                           Foreground="#52525B"
                           TextWrapping="Wrap"
                           Text="Name, IP and mask must match Network.ps1 expectations." />
                <StackPanel Name="NetworkCardsPanel"
                            Margin="0,16,0,0" />
              </StackPanel>
            </Border>

            <Border Grid.Column="2"
                    Background="#FFFFFF"
                    BorderBrush="#E5E7EB"
                    BorderThickness="1"
                    CornerRadius="22"
                    Padding="20">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto" />
                  </Grid.ColumnDefinitions>
                  <TextBlock Grid.Column="0"
                             Text="Setup Audit"
                             Foreground="#111111"
                             FontSize="20"
                             FontWeight="Bold" />
                  <TextBlock Grid.Column="1"
                             Name="AuditSummaryText"
                             VerticalAlignment="Center"
                             Foreground="#52525B"
                             Text="" />
                </Grid>
                <TextBlock Margin="0,8,0,0"
                           Foreground="#52525B"
                           TextWrapping="Wrap"
                           Text="Every modification expected from Setup is verified on this machine." />
                <StackPanel Name="AuditChecksPanel"
                            Margin="0,16,0,0" />
              </StackPanel>
            </Border>
          </Grid>

          <Border Grid.Row="4"
                  Background="#0B0B0C"
                  BorderBrush="#232326"
                  BorderThickness="1"
                  CornerRadius="22"
                  Padding="16">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto" />
                <RowDefinition Height="12" />
                <RowDefinition Height="*" />
              </Grid.RowDefinitions>

              <TextBlock Grid.Row="0"
                         Text="Live Logs"
                         FontSize="20"
                         FontWeight="Bold" />

              <ListBox Grid.Row="2"
                       Name="LogsListBox"
                       Background="#0B0B0C"
                       Foreground="#F4F4F5"
                       BorderThickness="0"
                       FontFamily="Consolas"
                       FontSize="13"
                       ScrollViewer.VerticalScrollBarVisibility="Auto"
                       ScrollViewer.HorizontalScrollBarVisibility="Auto" />
            </Grid>
          </Border>
        </Grid>
      </Grid>

      <Border Grid.Row="4"
              Background="#0F0F10"
              BorderBrush="#232326"
              BorderThickness="1"
              CornerRadius="22"
              Padding="16">
        <TextBlock Foreground="#B3B3BC"
                   TextWrapping="Wrap"
                   Text="Reports are saved next to the launcher in the reports folder. Network expectations come from assets\config\network-profiles.json. Audit checks live in assets\checks and can be added or modified without rebuilding the EXE." />
      </Border>
    </Grid>
  </ScrollViewer>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Global safety net: any unhandled exception on the UI dispatcher is logged
# but MARKED HANDLED so WPF does not tear down the window.
$window.Dispatcher.add_UnhandledException({
    param($sender, $eventArgs)
    try {
        $msg = "Unhandled UI exception: " + $eventArgs.Exception.Message
        Write-CrashLog $msg
        Write-CrashLog ("Stack: " + $eventArgs.Exception.StackTrace)
        if ($script:LogItems) { $script:LogItems.Add(("[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg)) | Out-Null }
    }
    catch { }
    $eventArgs.Handled = $true
})

# Catch fully unhandled domain exceptions too (worker threads etc.)
[AppDomain]::CurrentDomain.add_UnhandledException({
    param($sender, $eventArgs)
    try {
        $ex = $eventArgs.ExceptionObject
        $msg = if ($ex) { "$ex" } else { "unknown" }
        Write-CrashLog ("AppDomain unhandled: " + $msg)
        if ($script:LogItems) { $script:LogItems.Add(("[{0}] AppDomain error: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg)) | Out-Null }
    }
    catch { }
})

$actionsPanel        = $window.FindName("ActionsPanel")
$openReportButton    = $window.FindName("OpenReportButton")
$refreshNetworkButton = $window.FindName("RefreshNetworkButton")
$refreshAuditButton  = $window.FindName("RefreshAuditButton")
$clearLogsButton     = $window.FindName("ClearLogsButton")
$currentTaskText     = $window.FindName("CurrentTaskText")
$currentStatusText   = $window.FindName("CurrentStatusText")
$reportSummaryText   = $window.FindName("ReportSummaryText")
$stateFileText       = $window.FindName("StateFileText")
$logsListBox         = $window.FindName("LogsListBox")
$networkCardsPanel   = $window.FindName("NetworkCardsPanel")
$networkSummaryText  = $window.FindName("NetworkSummaryText")
$auditChecksPanel    = $window.FindName("AuditChecksPanel")
$auditSummaryText    = $window.FindName("AuditSummaryText")
$computerNameText    = $window.FindName("ComputerNameText")
$lastUpdatedText     = $window.FindName("LastUpdatedText")
$actionsHelpText     = $window.FindName("ActionsHelpText")

$logsListBox.ItemsSource = $script:LogItems
$computerNameText.Text   = $env:COMPUTERNAME
$stateFileText.Text      = $script:StateFilePath

$heroSubtitle = $window.FindName("HeroSubtitle")
if ($heroSubtitle) {
    $heroSubtitle.Text = "Assets folder: $script:AssetsRoot"
}

# ----------------------------------------------------------------------------
# Color helpers
# ----------------------------------------------------------------------------

function Get-Brush {
    param([string]$Hex)
    return [System.Windows.Media.BrushConverter]::new().ConvertFromString($Hex)
}

# ----------------------------------------------------------------------------
# Status / log helpers
# ----------------------------------------------------------------------------

function Set-Status {
    param([string]$TaskText, [string]$StatusText)
    $currentTaskText.Text = $TaskText
    $currentStatusText.Text = $StatusText
}

function Apply-StatusVisual {
    param($Border, $Text, [string]$State)

    switch ($State) {
        "Running" {
            $Border.Background = Get-Brush "#FEF3C7"
            $Text.Foreground   = Get-Brush "#92400E"
        }
        "Done" {
            $Border.Background = Get-Brush "#DCFCE7"
            $Text.Foreground   = Get-Brush "#166534"
        }
        "Error" {
            $Border.Background = Get-Brush "#FEE2E2"
            $Text.Foreground   = Get-Brush "#B91C1C"
        }
        default {
            $Border.Background = Get-Brush "#E5E7EB"
            $Text.Foreground   = Get-Brush "#374151"
        }
    }
    $Text.Text = $State
}

function Set-ActionState {
    param([string]$Key, [string]$State)

    if (-not $script:TasksByKey.ContainsKey($Key)) { return }
    $task = $script:TasksByKey[$Key]
    if (-not $task.StatusBorder -or -not $task.StatusText) { return }

    Apply-StatusVisual -Border $task.StatusBorder -Text $task.StatusText -State $State
}

function Append-LogLine {
    param([string]$Line)

    if ([string]::IsNullOrWhiteSpace($Line)) { return }

    $timestampedLine = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Line.TrimEnd()
    $script:CurrentLogLines.Add($timestampedLine) | Out-Null
    $script:LogItems.Add($timestampedLine)
    try { $logsListBox.ScrollIntoView($timestampedLine) } catch { }
}

function Set-ControlsBusyState {
    param([bool]$Busy)

    $script:IsBusy = $Busy
    foreach ($task in $script:Tasks) {
        if ($task.Button) { $task.Button.IsEnabled = -not $Busy }
    }
    $refreshNetworkButton.IsEnabled = -not $Busy
    $refreshAuditButton.IsEnabled   = -not $Busy
    $clearLogsButton.IsEnabled      = -not $Busy
    $openReportButton.IsEnabled     = (-not $Busy) -and [bool]$script:LastReportPath
}

# ----------------------------------------------------------------------------
# Action button construction
# ----------------------------------------------------------------------------

function Build-ActionButtons {
    $actionsPanel.Children.Clear()

    if ($script:Tasks.Count -eq 0) {
        $emptyText = New-Object System.Windows.Controls.TextBlock
        $emptyText.Text = "No script manifest found in assets. Drop a *.ps1 file plus a *.manifest.json next to it to add a button."
        $emptyText.Foreground = Get-Brush "#B91C1C"
        $emptyText.TextWrapping = "Wrap"
        [void]$actionsPanel.Children.Add($emptyText)
        return
    }

    $actionsHelpText.Text = "$($script:Tasks.Count) script(s) discovered in assets."

    $first = $true
    foreach ($task in $script:Tasks) {
        $row = New-Object System.Windows.Controls.Grid
        if (-not $first) { $row.Margin = [System.Windows.Thickness]::new(0, 14, 0, 0) }

        $colMain = New-Object System.Windows.Controls.ColumnDefinition
        [void]$row.ColumnDefinitions.Add($colMain)
        $colStatus = New-Object System.Windows.Controls.ColumnDefinition
        $colStatus.Width = [System.Windows.GridLength]::new(120)
        [void]$row.ColumnDefinitions.Add($colStatus)

        # ColorLoop accent button. Choose color via manifest field; default to green for primary, blue otherwise.
        $colorName = if ($task.Color) { [string]$task.Color } elseif ($task.Primary) { "green" } else { "blue" }
        $accent = Get-ColorLoopAccent -Name $colorName

        $button = New-Object System.Windows.Controls.Button
        $button.Style = $window.FindResource("RoundedButton")
        $button.Height = 56
        $button.FontWeight = "Bold"
        $button.FontSize   = 15
        $button.Content    = $task.Label
        $button.Background = Get-Brush $accent.Bg
        $button.Foreground = Get-Brush $accent.Fg
        $button.BorderThickness = [System.Windows.Thickness]::new(0)
        # NOTE: drop-shadow effect intentionally removed - was suspected
        # to cause ps2exe rendering crashes on click.
        if ($task.Description) { $button.ToolTip = $task.Description }
        [System.Windows.Controls.Grid]::SetColumn($button, 0)
        [void]$row.Children.Add($button)

        $statusBorder = New-Object System.Windows.Controls.Border
        $statusBorder.Margin = [System.Windows.Thickness]::new(12, 0, 0, 0)
        $statusBorder.CornerRadius = [System.Windows.CornerRadius]::new(14)
        $statusBorder.Background = Get-Brush $script:Palette.NeutralBg
        $statusBorder.Padding = [System.Windows.Thickness]::new(10, 0, 10, 0)
        [System.Windows.Controls.Grid]::SetColumn($statusBorder, 1)

        $statusText = New-Object System.Windows.Controls.TextBlock
        $statusText.VerticalAlignment   = "Center"
        $statusText.HorizontalAlignment = "Center"
        $statusText.FontWeight          = "Bold"
        $statusText.Text                = "Not done"
        $statusBorder.Child             = $statusText
        [void]$row.Children.Add($statusBorder)

        $task.Button       = $button
        $task.StatusBorder = $statusBorder
        $task.StatusText   = $statusText

        # Restore persisted state if any
        $persistedState = "Not done"
        if ($script:State.scripts -and $script:State.scripts.ContainsKey($task.Key)) {
            $persistedState = [string]$script:State.scripts[$task.Key].status
            if ([string]::IsNullOrWhiteSpace($persistedState)) { $persistedState = "Not done" }
        }
        Apply-StatusVisual -Border $statusBorder -Text $statusText -State $persistedState

        # Tag the button with the task key. Also store the mapping in a
        # script-level dictionary as a fallback in case ps2exe loses Tag.
        $button.Tag = $task.Key
        try { $script:ButtonKeyMap[$button] = $task.Key } catch { }

        $button.Add_Click({
            param($sender, $e)
            Write-CrashLog "----- Click handler entered -----"
            $clickedKey = ""

            # Try Tag first
            try { $clickedKey = [string]$sender.Tag } catch { Write-CrashLog "Tag read failed" }
            Write-CrashLog ("After Tag try, clickedKey='" + $clickedKey + "', sender type=" + ($(if ($sender) { $sender.GetType().Name } else { "null" })))

            # Fallback: dictionary lookup by button reference
            if ([string]::IsNullOrWhiteSpace($clickedKey)) {
                try {
                    if ($script:ButtonKeyMap.ContainsKey($sender)) {
                        $clickedKey = [string]$script:ButtonKeyMap[$sender]
                    }
                } catch { }
            }

            # Fallback 2: match by button label (Content)
            if ([string]::IsNullOrWhiteSpace($clickedKey)) {
                try {
                    $label = [string]$sender.Content
                    foreach ($t in $script:Tasks) {
                        if ($t.Label -eq $label) { $clickedKey = $t.Key; break }
                    }
                } catch { }
            }

            Write-CrashLog "Click received: '$clickedKey'"
            if ([string]::IsNullOrWhiteSpace($clickedKey)) {
                Write-CrashLog "Click ignored: could not resolve task key from Tag/dict/label"
                try { Append-LogLine "Click ignored: could not resolve task key" } catch { }
                return
            }

            try {
                # NEW path: file-redirect approach. No background-thread events.
                Start-TaskExecutionFileMode -TaskKey $clickedKey
                Write-CrashLog "Start-TaskExecutionFileMode returned for '$clickedKey'"
            }
            catch {
                Write-CrashLog ("Start-TaskExecutionFileMode threw: " + $_.Exception.Message)
                try { Append-LogLine ("Click handler error: " + $_.Exception.Message) } catch { }
                try { Set-Status -TaskText "Launcher error" -StatusText $_.Exception.Message } catch { }
                try { Set-ControlsBusyState -Busy $false } catch { }
                # Never re-throw - the WPF host would close
            }
        })

        [void]$actionsPanel.Children.Add($row)
        $first = $false
    }
}

# ----------------------------------------------------------------------------
# Network card rendering
# ----------------------------------------------------------------------------

function Add-NetworkCardElement {
    param($Snapshot)

    $border = New-Object System.Windows.Controls.Border
    $border.CornerRadius = [System.Windows.CornerRadius]::new(18)
    $border.BorderThickness = [System.Windows.Thickness]::new(1)
    $border.Padding = [System.Windows.Thickness]::new(14)
    $border.Margin = [System.Windows.Thickness]::new(0, 0, 0, 10)
    $border.Background = Get-Brush "#FAFAFA"
    $border.BorderBrush = Get-Brush "#E5E7EB"

    $stack = New-Object System.Windows.Controls.StackPanel
    $border.Child = $stack

    $headerGrid = New-Object System.Windows.Controls.Grid
    $headerGrid.Margin = [System.Windows.Thickness]::new(0, 0, 0, 10)
    [void]$headerGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
    $statusColumn = New-Object System.Windows.Controls.ColumnDefinition
    $statusColumn.Width = [System.Windows.GridLength]::new(92)
    [void]$headerGrid.ColumnDefinitions.Add($statusColumn)

    $nameText = New-Object System.Windows.Controls.TextBlock
    $nameText.Text = $Snapshot.Name
    $nameText.FontWeight = "Bold"
    $nameText.FontSize = 16
    $nameText.Foreground = Get-Brush "#111111"
    [System.Windows.Controls.Grid]::SetColumn($nameText, 0)
    [void]$headerGrid.Children.Add($nameText)

    $statusBorder = New-Object System.Windows.Controls.Border
    $statusBorder.CornerRadius = [System.Windows.CornerRadius]::new(10)
    $statusBorder.Padding = [System.Windows.Thickness]::new(10, 4, 10, 4)
    $statusBorder.HorizontalAlignment = "Right"
    if ($Snapshot.Status -eq "success") {
        $statusBorder.Background = Get-Brush "#DCFCE7"
        $statusTextColor = "#166534"
    }
    else {
        $statusBorder.Background = Get-Brush "#FEE2E2"
        $statusTextColor = "#B91C1C"
    }

    $statusText = New-Object System.Windows.Controls.TextBlock
    $statusText.Text = $Snapshot.StatusText
    $statusText.FontWeight = "Bold"
    $statusText.Foreground = Get-Brush $statusTextColor
    $statusBorder.Child = $statusText
    [System.Windows.Controls.Grid]::SetColumn($statusBorder, 1)
    [void]$headerGrid.Children.Add($statusBorder)

    [void]$stack.Children.Add($headerGrid)

    foreach ($line in @(
        "Expected Name: $($Snapshot.Name)",
        "Actual Name: $($Snapshot.ActualName)",
        "Expected IP: $($Snapshot.ExpectedIp)",
        "Actual IP: $($Snapshot.ActualIp)",
        "Expected Mask: $($Snapshot.ExpectedMask)",
        "Actual Mask: $($Snapshot.ActualMask)",
        "Detail: $($Snapshot.Detail)"
    )) {
        $text = New-Object System.Windows.Controls.TextBlock
        $text.Text = $line
        $text.Margin = [System.Windows.Thickness]::new(0, 2, 0, 0)
        $text.Foreground = Get-Brush "#52525B"
        [void]$stack.Children.Add($text)
    }

    [void]$networkCardsPanel.Children.Add($border)
}

function Refresh-NetworkCards {
    try {
        $networkCardsPanel.Children.Clear()
        $snapshots = @(Get-NetworkSnapshots)
        $okCount = 0
        foreach ($snapshot in $snapshots) {
            if ($snapshot.Status -eq "success") { $okCount++ }
            Add-NetworkCardElement -Snapshot $snapshot
        }
        $networkSummaryText.Text = "$okCount / $($snapshots.Count) OK"
    }
    catch {
        $networkCardsPanel.Children.Clear()
        $errorText = New-Object System.Windows.Controls.TextBlock
        $errorText.Text = "Unable to read network status: $($_.Exception.Message)"
        $errorText.Foreground = Get-Brush "#B91C1C"
        [void]$networkCardsPanel.Children.Add($errorText)
        $networkSummaryText.Text = "Error"
    }
}

# ----------------------------------------------------------------------------
# Audit panel rendering
# ----------------------------------------------------------------------------

function Get-AuditStatusVisual {
    param([string]$Status)

    switch ($Status) {
        "ok"      { return @{ Bg = "#DCFCE7"; Fg = "#166534"; Label = "OK" } }
        "missing" { return @{ Bg = "#FEE2E2"; Fg = "#B91C1C"; Label = "Missing" } }
        "partial" { return @{ Bg = "#FEF3C7"; Fg = "#92400E"; Label = "Partial" } }
        "unknown" { return @{ Bg = "#E5E7EB"; Fg = "#374151"; Label = "Unknown" } }
        default   { return @{ Bg = "#E5E7EB"; Fg = "#374151"; Label = "Unknown" } }
    }
}

function Add-AuditCardElement {
    param($Check, $Result)

    $visual = Get-AuditStatusVisual -Status $Result.Status

    $border = New-Object System.Windows.Controls.Border
    $border.CornerRadius = [System.Windows.CornerRadius]::new(14)
    $border.BorderThickness = [System.Windows.Thickness]::new(1)
    $border.Padding = [System.Windows.Thickness]::new(14, 10, 14, 10)
    $border.Margin = [System.Windows.Thickness]::new(0, 0, 0, 8)
    $border.Background = Get-Brush "#FAFAFA"
    $border.BorderBrush = Get-Brush "#E5E7EB"

    $grid = New-Object System.Windows.Controls.Grid
    $border.Child = $grid

    $colMain = New-Object System.Windows.Controls.ColumnDefinition
    [void]$grid.ColumnDefinitions.Add($colMain)
    $colStatus = New-Object System.Windows.Controls.ColumnDefinition
    $colStatus.Width = [System.Windows.GridLength]::new(96)
    [void]$grid.ColumnDefinitions.Add($colStatus)

    $textStack = New-Object System.Windows.Controls.StackPanel
    [System.Windows.Controls.Grid]::SetColumn($textStack, 0)

    $labelText = New-Object System.Windows.Controls.TextBlock
    $labelText.Text = $Check.Label
    $labelText.FontWeight = "Bold"
    $labelText.FontSize = 14
    $labelText.Foreground = Get-Brush "#111111"
    [void]$textStack.Children.Add($labelText)

    $detailText = New-Object System.Windows.Controls.TextBlock
    $detailText.Text = if ($Result.Detail) { [string]$Result.Detail } else { "" }
    $detailText.Margin = [System.Windows.Thickness]::new(0, 4, 0, 0)
    $detailText.Foreground = Get-Brush "#52525B"
    $detailText.TextWrapping = "Wrap"
    [void]$textStack.Children.Add($detailText)

    $categoryText = New-Object System.Windows.Controls.TextBlock
    $categoryText.Text = "Category: $($Check.Category)"
    $categoryText.Margin = [System.Windows.Thickness]::new(0, 4, 0, 0)
    $categoryText.Foreground = Get-Brush "#A1A1AA"
    $categoryText.FontSize = 11
    [void]$textStack.Children.Add($categoryText)

    [void]$grid.Children.Add($textStack)

    $statusBorder = New-Object System.Windows.Controls.Border
    $statusBorder.CornerRadius = [System.Windows.CornerRadius]::new(10)
    $statusBorder.Padding = [System.Windows.Thickness]::new(10, 4, 10, 4)
    $statusBorder.HorizontalAlignment = "Right"
    $statusBorder.VerticalAlignment   = "Center"
    $statusBorder.Background = Get-Brush $visual.Bg

    $statusText = New-Object System.Windows.Controls.TextBlock
    $statusText.Text = $visual.Label
    $statusText.FontWeight = "Bold"
    $statusText.Foreground = Get-Brush $visual.Fg
    $statusBorder.Child = $statusText

    [System.Windows.Controls.Grid]::SetColumn($statusBorder, 1)
    [void]$grid.Children.Add($statusBorder)

    [void]$auditChecksPanel.Children.Add($border)
}

function Refresh-AuditPanel {
    try {
        $auditChecksPanel.Children.Clear()

        if (-not $script:AuditChecks -or $script:AuditChecks.Count -eq 0) {
            $emptyText = New-Object System.Windows.Controls.TextBlock
            $emptyText.Text = "No audit check found in assets\checks. Drop a *.check.ps1 file there to add one."
            $emptyText.Foreground = Get-Brush "#52525B"
            $emptyText.TextWrapping = "Wrap"
            [void]$auditChecksPanel.Children.Add($emptyText)
            $auditSummaryText.Text = "0 checks"
            return
        }

        $okCount = 0
        $missingCount = 0
        $partialCount = 0
        $unknownCount = 0

        foreach ($check in $script:AuditChecks) {
            $result = $null
            try {
                $result = & $check.Test
            }
            catch {
                $result = @{ Status = "unknown"; Detail = "Check error: $($_.Exception.Message)" }
            }

            if (-not ($result -is [hashtable]) -or -not $result.ContainsKey('Status')) {
                $result = @{ Status = "unknown"; Detail = "Check returned no status" }
            }

            switch ([string]$result.Status) {
                "ok"      { $okCount++ }
                "missing" { $missingCount++ }
                "partial" { $partialCount++ }
                default   { $unknownCount++ }
            }

            Add-AuditCardElement -Check $check -Result $result
        }

        $total = $script:AuditChecks.Count
        $parts = @("$okCount / $total OK")
        if ($missingCount -gt 0) { $parts += "$missingCount missing" }
        if ($partialCount -gt 0) { $parts += "$partialCount partial" }
        if ($unknownCount -gt 0) { $parts += "$unknownCount unknown" }
        $auditSummaryText.Text = ($parts -join " · ")
    }
    catch {
        $auditChecksPanel.Children.Clear()
        $errorText = New-Object System.Windows.Controls.TextBlock
        $errorText.Text = "Unable to run audit: $($_.Exception.Message)"
        $errorText.Foreground = Get-Brush "#B91C1C"
        [void]$auditChecksPanel.Children.Add($errorText)
        $auditSummaryText.Text = "Error"
    }
}

# ----------------------------------------------------------------------------
# Reports
# ----------------------------------------------------------------------------

function Get-ReportFilePath {
    param([string]$TaskName)

    if (-not (Test-Path $script:ReportsRoot)) {
        New-Item -Path $script:ReportsRoot -ItemType Directory -Force | Out-Null
    }

    $safeTask = $TaskName -replace "[^A-Za-z0-9_-]", "_"
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    return Join-Path $script:ReportsRoot "$timestamp-$safeTask-report.html"
}

function Convert-StepSummariesToHtml {
    param($StepSummaries)

    $cards = foreach ($step in $StepSummaries) {
        $statusClass = if ($step.Status -eq "success") { "step-success" } else { "step-error" }
        $statusLabel = if ($step.Status -eq "success") { "Done" } else { "Error" }
        $details = foreach ($detail in @($step.Details)) {
            "<li>{0}</li>" -f (Escape-Html $detail)
        }

        @"
<div class="step-card $statusClass">
  <div class="step-header">
    <div class="step-title">$(Escape-Html $step.Title)</div>
    <div class="step-badge">$statusLabel</div>
  </div>
  <ul>
    $($details -join "`n")
  </ul>
</div>
"@
    }

    return ($cards -join "`n")
}

function Convert-NetworkSnapshotsToHtml {
    param($Snapshots)

    $cards = foreach ($snapshot in $Snapshots) {
        $statusClass = if ($snapshot.Status -eq "success") { "step-success" } else { "step-error" }
        @"
<div class="network-card $statusClass">
  <div class="step-header">
    <div class="step-title">$(Escape-Html $snapshot.Name)</div>
    <div class="step-badge">$(Escape-Html $snapshot.StatusText)</div>
  </div>
  <div class="card-line"><strong>Expected Name:</strong> $(Escape-Html $snapshot.Name)</div>
  <div class="card-line"><strong>Actual Name:</strong> $(Escape-Html $snapshot.ActualName)</div>
  <div class="card-line"><strong>Expected IP:</strong> $(Escape-Html $snapshot.ExpectedIp)</div>
  <div class="card-line"><strong>Actual IP:</strong> $(Escape-Html $snapshot.ActualIp)</div>
  <div class="card-line"><strong>Expected Mask:</strong> $(Escape-Html $snapshot.ExpectedMask)</div>
  <div class="card-line"><strong>Actual Mask:</strong> $(Escape-Html $snapshot.ActualMask)</div>
  <div class="card-line"><strong>Detail:</strong> $(Escape-Html $snapshot.Detail)</div>
</div>
"@
    }

    return ($cards -join "`n")
}

function Convert-LogsToHtml {
    $items = foreach ($line in $script:CurrentLogLines) {
        $severity = Get-LineSeverity -Line $line
        "<li class=`"log-$severity`">{0}</li>" -f (Escape-Html $line)
    }

    return ($items -join "`n")
}

function Write-RunReport {
    param([string]$TaskName, [string]$ScriptPath, [int]$ExitCode)

    try {
        $reportPath = Get-ReportFilePath -TaskName $TaskName
        $finishedAt = Get-Date
        $duration = [int][Math]::Round(($finishedAt - $script:RunStartedAt).TotalSeconds)
        $statusLabel = if ($ExitCode -eq 0) { "Done" } else { "Error" }
        $statusClass = if ($ExitCode -eq 0) { "step-success" } else { "step-error" }
        $stepHtml = Convert-StepSummariesToHtml -StepSummaries (Get-StepSummaries -Lines $script:CurrentLogLines.ToArray() -ExitCode $ExitCode)
        $networkHtml = Convert-NetworkSnapshotsToHtml -Snapshots (Get-NetworkSnapshots)
        $logsHtml = Convert-LogsToHtml

        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>Rutherford Assistant Report</title>
  <style>
    body { background:#050505; color:#f4f4f5; font-family:Segoe UI, sans-serif; margin:0; padding:28px; }
    .shell { max-width:1200px; margin:0 auto; }
    .hero, .section { background:#101012; border:1px solid #232326; border-radius:22px; padding:22px; margin-bottom:18px; }
    .eyebrow { color:#b3b3bc; text-transform:uppercase; letter-spacing:0.12em; font-size:12px; margin:0 0 8px; }
    h1, h2 { margin:0; }
    .summary-grid, .steps-grid, .network-grid { display:grid; gap:14px; }
    .summary-grid { grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); margin-top:16px; }
    .steps-grid, .network-grid { grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); margin-top:16px; }
    .summary-card, .step-card, .network-card { background:#17171A; border:1px solid #2A2A2E; border-radius:18px; padding:16px; }
    .step-success { border-color:#166534; box-shadow:inset 0 0 0 1px rgba(22,101,52,0.25); }
    .step-error { border-color:#B91C1C; box-shadow:inset 0 0 0 1px rgba(185,28,28,0.25); }
    .summary-status { display:inline-block; padding:8px 12px; border-radius:999px; font-weight:700; }
    .step-success .summary-status, .step-success .step-badge { background:#DCFCE7; color:#166534; }
    .step-error .summary-status, .step-error .step-badge { background:#FEE2E2; color:#B91C1C; }
    .step-header { display:flex; justify-content:space-between; gap:12px; align-items:center; margin-bottom:10px; }
    .step-title { font-size:18px; font-weight:700; }
    .step-badge { padding:6px 10px; border-radius:999px; font-size:13px; font-weight:700; }
    .card-line { color:#D4D4D8; margin-top:6px; }
    ul { margin:12px 0 0; padding-left:18px; }
    li { margin-top:6px; }
    .log-success { color:#BBF7D0; }
    .log-error { color:#FECACA; }
    .log-neutral { color:#E4E4E7; }
  </style>
</head>
<body>
  <div class="shell">
    <div class="hero">
      <div class="eyebrow">Rutherford Assistant Report</div>
      <h1>$(Escape-Html $TaskName)</h1>
      <div class="summary-grid">
        <div class="summary-card $statusClass">
          <div class="eyebrow">Status</div>
          <div class="summary-status">$statusLabel</div>
        </div>
        <div class="summary-card">
          <div class="eyebrow">Script</div>
          <div>$(Escape-Html $ScriptPath)</div>
        </div>
        <div class="summary-card">
          <div class="eyebrow">Computer</div>
          <div>$(Escape-Html $env:COMPUTERNAME)</div>
        </div>
        <div class="summary-card">
          <div class="eyebrow">User</div>
          <div>$(Escape-Html $env:USERNAME)</div>
        </div>
        <div class="summary-card">
          <div class="eyebrow">Started</div>
          <div>$(Escape-Html $script:RunStartedAt.ToString("s"))</div>
        </div>
        <div class="summary-card">
          <div class="eyebrow">Finished</div>
          <div>$(Escape-Html $finishedAt.ToString("s"))</div>
        </div>
        <div class="summary-card">
          <div class="eyebrow">Duration</div>
          <div>$duration seconds</div>
        </div>
        <div class="summary-card">
          <div class="eyebrow">Exit Code</div>
          <div>$ExitCode</div>
        </div>
      </div>
    </div>

    <div class="section">
      <div class="eyebrow">Steps</div>
      <h2>Execution Summary</h2>
      <div class="steps-grid">
        $stepHtml
      </div>
    </div>

    <div class="section">
      <div class="eyebrow">Network</div>
      <h2>Adapter Status</h2>
      <div class="network-grid">
        $networkHtml
      </div>
    </div>

    <div class="section">
      <div class="eyebrow">Logs</div>
      <h2>Detailed Output</h2>
      <ul>
        $logsHtml
      </ul>
    </div>
  </div>
</body>
</html>
"@

        Set-Content -Path $reportPath -Value $html -Encoding UTF8
        $script:LastReportPath = $reportPath
        $reportSummaryText.Text = $reportPath
        $openReportButton.IsEnabled = $true
        return $reportPath
    }
    catch {
        Append-LogLine ("Report generation failed: " + $_.Exception.Message)
        return $null
    }
}

# ----------------------------------------------------------------------------
# Task execution
# ----------------------------------------------------------------------------

function Start-TaskExecution {
    param([string]$TaskKey)

    Write-CrashLog "Start-TaskExecution: enter (key='$TaskKey', busy=$script:IsBusy)"

    if ($script:IsBusy) { Write-CrashLog "Start-TaskExecution: aborted, already busy"; return }

    if (-not $script:TasksByKey.ContainsKey($TaskKey)) {
        Write-CrashLog "Start-TaskExecution: unknown task key '$TaskKey'"
        try { Append-LogLine "Unknown task key: '$TaskKey'. Available: $(@($script:TasksByKey.Keys) -join ', ')" } catch { }
        return
    }

    $task = $script:TasksByKey[$TaskKey]
    Write-CrashLog "Start-TaskExecution: task resolved (label='$($task.Label)' script='$($task.ScriptPath)')"

    if (-not (Test-Path $task.ScriptPath)) {
        Write-CrashLog "Start-TaskExecution: script file missing"
        [System.Windows.MessageBox]::Show("Missing script: $($task.ScriptPath)", "Rutherford Assistant") | Out-Null
        return
    }

    $script:LogItems.Clear()
    $script:CurrentLogLines.Clear()
    $script:LastReportPath = $null
    $reportSummaryText.Text = "Report will be generated when the task finishes."
    $openReportButton.IsEnabled = $false
    $script:CurrentTask = $task
    $script:CurrentTaskKey = $TaskKey
    $script:RunStartedAt = Get-Date

    Set-ActionState -Key $TaskKey -State "Running"
    Set-ControlsBusyState -Busy $true
    Set-Status -TaskText $task.Label -StatusText "Running... keep this window open."

    Append-LogLine "Launcher root: $script:AppRoot"
    Append-LogLine "Assets root: $script:AssetsRoot"
    Append-LogLine "Starting task: $($task.Label)"
    Append-LogLine "Script: $($task.ScriptPath)"

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = "powershell.exe"
    $startInfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$($task.ScriptPath)`""
    $startInfo.WorkingDirectory = $script:AssetsRoot
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    $process.EnableRaisingEvents = $true

    # ====================================================================
    # CRITICAL: these handlers run on background threads. They MUST NOT
    # touch the UI directly and MUST NOT call Dispatcher.BeginInvoke.
    # All they do is push to a thread-safe queue. The UI thread drains
    # the queue via $script:OutputDrainTimer. This eliminates cross-thread
    # WPF interactions, the prime suspect for silent ps2exe terminations.
    # ====================================================================

    $process.add_OutputDataReceived({
        param($sender, $eventArgs)
        if ($null -ne $eventArgs.Data) {
            try { $script:OutputQueue.Enqueue([string]$eventArgs.Data) } catch { }
        }
    })

    $process.add_ErrorDataReceived({
        param($sender, $eventArgs)
        if ($null -ne $eventArgs.Data) {
            try { $script:OutputQueue.Enqueue("ERROR: " + [string]$eventArgs.Data) } catch { }
        }
    })

    $process.add_Exited({
        param($sender, $eventArgs)
        $exitCode = -1
        try { $exitCode = $sender.ExitCode } catch { }
        try { $script:ExitedSignalQueue.Enqueue([int]$exitCode) } catch { }
    })

    Write-CrashLog "Start-TaskExecution: about to call Process.Start"
    try {
        $started = $process.Start()
        Write-CrashLog "Start-TaskExecution: Process.Start returned $started"
    }
    catch {
        $started = $false
        Write-CrashLog ("Start-TaskExecution: Process.Start threw: " + $_.Exception.Message)
        Append-LogLine ("Process start failed: " + $_.Exception.Message)
    }

    if (-not $started) {
        Set-ControlsBusyState -Busy $false
        Set-ActionState -Key $TaskKey -State "Error"
        [System.Windows.MessageBox]::Show("Unable to start $($task.ScriptPath)", "Rutherford Assistant") | Out-Null
        return
    }

    $script:CurrentProcess = $process
    try {
        $process.BeginOutputReadLine()
        Write-CrashLog "Start-TaskExecution: BeginOutputReadLine OK"
    } catch {
        Write-CrashLog ("BeginOutputReadLine threw: " + $_.Exception.Message)
        Append-LogLine ("BeginOutputReadLine failed: " + $_.Exception.Message)
    }
    try {
        $process.BeginErrorReadLine()
        Write-CrashLog "Start-TaskExecution: BeginErrorReadLine OK"
    }  catch {
        Write-CrashLog ("BeginErrorReadLine threw: " + $_.Exception.Message)
        Append-LogLine ("BeginErrorReadLine failed: " + $_.Exception.Message)
    }
    Write-CrashLog "Start-TaskExecution: exit clean"
}

# ============================================================================
# NEW: file-redirect-based task execution. Replaces the System.Diagnostics
# event-based approach. ZERO callbacks on background threads.
# ============================================================================

function Start-TaskExecutionFileMode {
    param([string]$TaskKey)

    Write-CrashLog "[FileMode] Start-TaskExecution: enter (key='$TaskKey', busy=$script:IsBusy)"
    if ($script:IsBusy) { Write-CrashLog "[FileMode] aborted, already busy"; return }

    if (-not $script:TasksByKey.ContainsKey($TaskKey)) {
        Write-CrashLog "[FileMode] unknown task '$TaskKey'"
        try { Append-LogLine "Unknown task: '$TaskKey'" } catch { }
        return
    }
    $task = $script:TasksByKey[$TaskKey]
    if (-not (Test-Path $task.ScriptPath)) {
        Write-CrashLog "[FileMode] script missing: $($task.ScriptPath)"
        [System.Windows.MessageBox]::Show("Missing script: $($task.ScriptPath)", "Rutherford Assistant") | Out-Null
        return
    }

    $script:LogItems.Clear()
    $script:CurrentLogLines.Clear()
    $script:LastReportPath = $null
    $reportSummaryText.Text = "Report will be generated when the task finishes."
    $openReportButton.IsEnabled = $false
    $script:CurrentTask = $task
    $script:CurrentTaskKey = $TaskKey
    $script:RunStartedAt = Get-Date
    $script:CompletionHandled = $false

    Set-ActionState -Key $TaskKey -State "Running"
    Set-ControlsBusyState -Busy $true
    Set-Status -TaskText $task.Label -StatusText "Running... keep this window open."

    Append-LogLine "Launcher root: $script:AppRoot"
    Append-LogLine "Assets root: $script:AssetsRoot"
    Append-LogLine "Starting task: $($task.Label)"
    Append-LogLine "Script: $($task.ScriptPath)"

    # Prepare output / error redirection files (in TEMP, always writable)
    $stamp = Get-Date -Format "yyyyMMdd-HHmmss-fff"
    $script:CurrentOutputFile = Join-Path $env:TEMP ("rutherford-out-$stamp.log")
    $script:CurrentErrorFile  = Join-Path $env:TEMP ("rutherford-err-$stamp.log")
    $script:OutputFilePosition = [int64]0
    $script:ErrorFilePosition  = [int64]0
    try {
        Set-Content -Path $script:CurrentOutputFile -Value "" -Encoding UTF8 -Force
        Set-Content -Path $script:CurrentErrorFile  -Value "" -Encoding UTF8 -Force
    } catch { Write-CrashLog ("Could not prepare temp files: " + $_.Exception.Message) }
    Write-CrashLog ("[FileMode] OutFile=" + $script:CurrentOutputFile)
    Write-CrashLog ("[FileMode] ErrFile=" + $script:CurrentErrorFile)

    Write-CrashLog "[FileMode] about to Start-Process"
    try {
        $argLine = "-NoProfile -ExecutionPolicy Bypass -File `"$($task.ScriptPath)`""
        Write-CrashLog ("[FileMode] argLine=" + $argLine)
        $proc = Start-Process -FilePath "powershell.exe" `
            -ArgumentList $argLine `
            -RedirectStandardOutput $script:CurrentOutputFile `
            -RedirectStandardError  $script:CurrentErrorFile `
            -WindowStyle Hidden `
            -PassThru
        Write-CrashLog ("[FileMode] Start-Process returned PID=" + $(if ($proc) { $proc.Id } else { "null" }))
    }
    catch {
        Write-CrashLog ("[FileMode] Start-Process threw: " + $_.Exception.Message)
        Append-LogLine ("Process start failed: " + $_.Exception.Message)
        Set-ControlsBusyState -Busy $false
        Set-ActionState -Key $TaskKey -State "Error"
        return
    }

    if (-not $proc) {
        Write-CrashLog "[FileMode] Start-Process returned null"
        Set-ControlsBusyState -Busy $false
        Set-ActionState -Key $TaskKey -State "Error"
        return
    }

    $script:CurrentProcess = $proc
    Write-CrashLog "[FileMode] CurrentProcess set, leaving the rest to the drain timer"
}

# Helper: read new bytes from a redirect file, decode, append as log lines.
function Drain-RedirectFile {
    param(
        [string]$Path,
        [ref]$PositionRef,
        [string]$Prefix = ""
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (-not (Test-Path -LiteralPath $Path)) { return }

    try {
        $stream = [System.IO.File]::Open(
            $Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
        )
        try {
            if ($PositionRef.Value -gt $stream.Length) { $PositionRef.Value = 0 }
            $stream.Position = $PositionRef.Value
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
            while (-not $reader.EndOfStream) {
                $line = $reader.ReadLine()
                if ($null -ne $line) {
                    $trimmed = $line.TrimEnd("`r")
                    if (-not [string]::IsNullOrWhiteSpace($trimmed)) {
                        try { Append-LogLine ($Prefix + $trimmed) } catch { }
                    }
                }
            }
            $PositionRef.Value = $stream.Position
        }
        finally { $stream.Close() }
    }
    catch {
        try { Write-CrashLog ("Drain-RedirectFile error on " + $Path + ": " + $_.Exception.Message) } catch { }
    }
}

# ----------------------------------------------------------------------------
# Wire up buttons
# ----------------------------------------------------------------------------

Build-ActionButtons

$refreshNetworkButton.Add_Click({
    try { Refresh-NetworkCards } catch { Append-LogLine ("Refresh-NetworkCards failed: " + $_.Exception.Message) }
})

$refreshAuditButton.Add_Click({
    try { Refresh-AuditPanel } catch { Append-LogLine ("Refresh-AuditPanel failed: " + $_.Exception.Message) }
})

$clearLogsButton.Add_Click({
    if ($script:IsBusy) { return }
    $script:LogItems.Clear()
    $script:CurrentLogLines.Clear()
    Set-Status -TaskText "Ready" -StatusText "Waiting for action."
})

$openReportButton.Add_Click({
    try {
        if ($script:LastReportPath -and (Test-Path $script:LastReportPath)) {
            Start-Process -FilePath $script:LastReportPath | Out-Null
        }
    }
    catch {
        Append-LogLine ("Open report failed: " + $_.Exception.Message)
    }
})

$window.Add_Closing({
    param($sender, $eventArgs)
    Write-CrashLog "Window.Closing event fired"

    if ($script:IsBusy -and $script:CurrentProcess -and -not $script:CurrentProcess.HasExited) {
        $result = [System.Windows.MessageBox]::Show(
            "A task is still running. Close anyway?",
            "Rutherford Assistant",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )

        if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
            $eventArgs.Cancel = $true
            Write-CrashLog "Window.Closing cancelled by user"
        }
        else {
            Write-CrashLog "Window.Closing confirmed by user"
        }
    }
})

$window.Add_Closed({
    Write-CrashLog "Window.Closed event fired (window destroyed)"
})

$window.Add_Deactivated({
    Write-CrashLog ("Window.Deactivated (focus lost). State=" + $window.WindowState + " IsVisible=" + $window.IsVisible)
})

$window.Add_StateChanged({
    Write-CrashLog ("Window.StateChanged -> " + $window.WindowState)
})

$window.Add_IsVisibleChanged({
    Write-CrashLog ("Window.IsVisibleChanged -> " + $window.IsVisible)
})

# ----------------------------------------------------------------------------
# Background timer for network refresh
# ----------------------------------------------------------------------------

$networkRefreshTimer = New-Object System.Windows.Threading.DispatcherTimer
$networkRefreshTimer.Interval = [TimeSpan]::FromSeconds(5)
$networkRefreshTimer.Add_Tick({
    if (-not $script:IsBusy) {
        try { Refresh-NetworkCards } catch { }
    }
})
$networkRefreshTimer.Start()

# ----------------------------------------------------------------------------
# Output drain timer - runs on UI thread, drains the queues populated by
# background OutputDataReceived/ErrorDataReceived/Exited handlers.
# All UI work happens HERE on the dispatcher thread, so no cross-thread bug.
# ----------------------------------------------------------------------------

function Handle-TaskCompletion {
    param([int]$ExitCode)

    Write-CrashLog "Handle-TaskCompletion: enter (ExitCode=$ExitCode)"
    try {
        Append-LogLine ("Process finished with exit code " + $ExitCode)

        $finalState = if ($ExitCode -eq 0) { "Done" } else { "Error" }
        $taskKey = $script:CurrentTaskKey
        $taskRef = $script:CurrentTask

        if ($taskKey) { try { Set-ActionState -Key $taskKey -State $finalState } catch { } }
        if ($taskKey) { try { Update-ScriptState -Key $taskKey -Status $finalState -ExitCode $ExitCode } catch { } }

        $reportPath = $null
        if ($taskRef) {
            try {
                $reportPath = Write-RunReport -TaskName $taskRef.Label -ScriptPath $taskRef.ScriptPath -ExitCode $ExitCode
            } catch { try { Append-LogLine ("Report error: " + $_.Exception.Message) } catch { } }
        }

        if ($taskRef) {
            if ($reportPath) {
                try { Set-Status -TaskText $taskRef.Label -StatusText ("$finalState. Report saved to $reportPath") } catch { }
            }
            else {
                try { Set-Status -TaskText $taskRef.Label -StatusText "$finalState. (No report - see logs)" } catch { }
            }
        }

        try { Refresh-NetworkCards } catch { Write-CrashLog ("Refresh-NetworkCards error: " + $_.Exception.Message) }
        if ($taskRef -and $taskRef.AuditAfterRun) {
            try { Refresh-AuditPanel } catch { Write-CrashLog ("Refresh-AuditPanel error: " + $_.Exception.Message) }
        }

        if ($script:State.lastUpdated) {
            try { $lastUpdatedText.Text = "Last update: $($script:State.lastUpdated)" } catch { }
        }
    }
    catch {
        Write-CrashLog ("Handle-TaskCompletion crashed: " + $_.Exception.Message)
        try { Append-LogLine ("Launcher post-run error: " + $_.Exception.Message) } catch { }
    }
    finally {
        $script:CurrentProcess = $null
        try { Set-ControlsBusyState -Busy $false } catch { }
        Write-CrashLog "Handle-TaskCompletion: exit"
    }
}

$script:OutputDrainTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:OutputDrainTimer.Interval = [TimeSpan]::FromMilliseconds(150)
$script:OutputDrainTimer.Add_Tick({
    try {
        # 1. Read any new output / error from the redirect files
        if ($script:CurrentOutputFile) {
            Drain-RedirectFile -Path $script:CurrentOutputFile -PositionRef ([ref]$script:OutputFilePosition) -Prefix ""
        }
        if ($script:CurrentErrorFile) {
            Drain-RedirectFile -Path $script:CurrentErrorFile -PositionRef ([ref]$script:ErrorFilePosition) -Prefix "ERROR: "
        }

        # 2. Detect process exit
        if ($script:CurrentProcess -and -not $script:CompletionHandled) {
            $hasExited = $false
            try { $hasExited = $script:CurrentProcess.HasExited } catch { }
            if ($hasExited) {
                $exitCode = -1
                try { $exitCode = $script:CurrentProcess.ExitCode } catch { }
                Write-CrashLog ("Drain timer: process has exited with code " + $exitCode)

                # Final flush of any remaining bytes
                if ($script:CurrentOutputFile) {
                    Drain-RedirectFile -Path $script:CurrentOutputFile -PositionRef ([ref]$script:OutputFilePosition) -Prefix ""
                }
                if ($script:CurrentErrorFile) {
                    Drain-RedirectFile -Path $script:CurrentErrorFile -PositionRef ([ref]$script:ErrorFilePosition) -Prefix "ERROR: "
                }

                $script:CompletionHandled = $true
                Handle-TaskCompletion -ExitCode $exitCode

                # Clean up temp files (best effort)
                try { if ($script:CurrentOutputFile -and (Test-Path -LiteralPath $script:CurrentOutputFile)) { Remove-Item -LiteralPath $script:CurrentOutputFile -Force -ErrorAction SilentlyContinue } } catch { }
                try { if ($script:CurrentErrorFile  -and (Test-Path -LiteralPath $script:CurrentErrorFile))  { Remove-Item -LiteralPath $script:CurrentErrorFile  -Force -ErrorAction SilentlyContinue } } catch { }
                $script:CurrentOutputFile = $null
                $script:CurrentErrorFile  = $null
            }
        }
    }
    catch {
        try { Write-CrashLog ("OutputDrainTimer crashed: " + $_.Exception.Message) } catch { }
    }
})
$script:OutputDrainTimer.Start()
Write-CrashLog "OutputDrainTimer started"

# ----------------------------------------------------------------------------
# Initial render
# ----------------------------------------------------------------------------

if ($script:State.lastUpdated) {
    $lastUpdatedText.Text = "Last update: $($script:State.lastUpdated)"
}
else {
    $lastUpdatedText.Text = "No previous run on this PC."
}

Refresh-NetworkCards
Refresh-AuditPanel
Set-Status -TaskText "Ready" -StatusText "Waiting for action."

$window.Add_SourceInitialized({
    Append-LogLine ("Launcher ready. Tasks: $($script:Tasks.Count). Audit checks: $($script:AuditChecks.Count).")
    Append-LogLine ("Self path: $script:LauncherPath")
    Append-LogLine ("App root : $script:AppRoot")
    Append-LogLine ("Assets   : $script:AssetsRoot")
    Append-LogLine ("Checks   : $script:ChecksRoot")
    Append-LogLine ("Net cfg  : $script:NetworkProfilesPath")

    if ($script:DiscoveryDiagnostics) {
        foreach ($diag in $script:DiscoveryDiagnostics) {
            Append-LogLine ("DIAG: " + $diag)
        }
    }

    if ($script:Tasks.Count -eq 0) {
        Append-LogLine "WARNING: no script manifest found. Verify that assets\*.manifest.json files are next to the EXE."
    }
})

[void]$window.ShowDialog()
