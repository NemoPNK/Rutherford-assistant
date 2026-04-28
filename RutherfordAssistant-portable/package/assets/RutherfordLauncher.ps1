Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase

# ----------------------------------------------------------------------------
# Path resolution
# ----------------------------------------------------------------------------

$script:LauncherPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
$script:SelfRoot = Split-Path -Parent $script:LauncherPath

if (Test-Path (Join-Path $script:SelfRoot "assets")) {
    $script:AppRoot = $script:SelfRoot
    $script:AssetsRoot = Join-Path $script:AppRoot "assets"
}
elseif ((Split-Path -Leaf $script:SelfRoot) -eq "assets") {
    $script:AssetsRoot = $script:SelfRoot
    $script:AppRoot = Split-Path -Parent $script:AssetsRoot
}
else {
    $script:AssetsRoot = $script:SelfRoot
    $script:AppRoot = Split-Path -Parent $script:AssetsRoot
}

$script:ChecksRoot         = Join-Path $script:AssetsRoot "checks"
$script:ReportsRoot        = Join-Path $script:AppRoot "reports"
$script:NetworkProfilesPath = Join-Path $script:AssetsRoot "config\network-profiles.json"
$script:StateRoot          = "C:\ProgramData\Rutherford"
$script:StateFilePath      = Join-Path $script:StateRoot "launcher-state.json"

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

    if (-not (Test-Path $script:AssetsRoot)) {
        return $tasks
    }

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
    if (-not (Test-Path $script:ChecksRoot)) { return $checks }

    Get-ChildItem -Path $script:ChecksRoot -Filter "*.check.ps1" -File -ErrorAction SilentlyContinue |
        Sort-Object Name |
        ForEach-Object {
            try {
                $check = & $_.FullName
                if ($check -is [hashtable] -and $check.ContainsKey('Test')) {
                    if (-not $check.ContainsKey('Order')) { $check.Order = 999 }
                    if (-not $check.ContainsKey('Label')) { $check.Label = $_.BaseName }
                    if (-not $check.ContainsKey('Category')) { $check.Category = "General" }
                    $check.SourcePath = $_.FullName
                    $checks += $check
                }
            }
            catch {
                # Bad check file, skip silently
            }
        }

    return @($checks | Sort-Object @{Expression = { $_.Order }}, @{Expression = { $_.Label }})
}

# ----------------------------------------------------------------------------
# Network helpers (unchanged)
# ----------------------------------------------------------------------------

function Read-NetworkProfiles {
    if (-not (Test-Path $script:NetworkProfilesPath)) {
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
              <Ellipse Width="14" Height="14" Fill="#FF5A1F" Margin="0,0,8,0" />
              <Ellipse Width="14" Height="14" Fill="#FFC83D" Margin="0,0,8,0" />
              <Ellipse Width="14" Height="14" Fill="#39D0FF" Margin="0,0,8,0" />
              <Ellipse Width="14" Height="14" Fill="#7C4DFF" />
            </StackPanel>
            <TextBlock Margin="0,14,0,0"
                       Text="Rutherford Assistant"
                       FontSize="30"
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
                    Margin="0,22,0,0"
                    Height="40"
                    FontWeight="Bold"
                    Background="#FFFFFF"
                    Foreground="#111111"
                    BorderBrush="#E5E7EB"
                    BorderThickness="1"
                    IsEnabled="False"
                    Content="Open Last Report" />

            <Button Name="RefreshNetworkButton"
                    Margin="0,10,0,0"
                    Height="40"
                    FontWeight="Bold"
                    Background="#FFFFFF"
                    Foreground="#111111"
                    BorderBrush="#E5E7EB"
                    BorderThickness="1"
                    Content="Refresh Network Status" />

            <Button Name="RefreshAuditButton"
                    Margin="0,10,0,0"
                    Height="40"
                    FontWeight="Bold"
                    Background="#FFFFFF"
                    Foreground="#111111"
                    BorderBrush="#E5E7EB"
                    BorderThickness="1"
                    Content="Refresh LaRoche Audit" />

            <Button Name="ClearLogsButton"
                    Margin="0,10,0,0"
                    Height="40"
                    FontWeight="Bold"
                    Background="#FFFFFF"
                    Foreground="#111111"
                    BorderBrush="#E5E7EB"
                    BorderThickness="1"
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

          <Border Grid.Row="2"
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
                         Text="Name, IP and mask must match Network.ps1 expectations." />
              <StackPanel Name="NetworkCardsPanel"
                          Margin="0,16,0,0" />
            </StackPanel>
          </Border>

          <Border Grid.Row="4"
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
                           Text="LaRoche Audit"
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
                         Text="Verifies that every modification expected from LaRoche.ps1 is actually applied on this machine." />
              <StackPanel Name="AuditChecksPanel"
                          Margin="0,16,0,0" />
            </StackPanel>
          </Border>

          <Border Grid.Row="6"
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
        if (-not $first) { $row.Margin = [System.Windows.Thickness]::new(0, 12, 0, 0) }

        $colMain = New-Object System.Windows.Controls.ColumnDefinition
        [void]$row.ColumnDefinitions.Add($colMain)
        $colStatus = New-Object System.Windows.Controls.ColumnDefinition
        $colStatus.Width = [System.Windows.GridLength]::new(120)
        [void]$row.ColumnDefinitions.Add($colStatus)

        $button = New-Object System.Windows.Controls.Button
        $button.Height = 50
        $button.FontWeight = "Bold"
        $button.Content = $task.Label
        if ($task.Primary) {
            $button.Background = Get-Brush "#111111"
            $button.Foreground = Get-Brush "#FFFFFF"
            $button.BorderThickness = [System.Windows.Thickness]::new(0)
        }
        else {
            $button.Background = Get-Brush "#F3F4F6"
            $button.Foreground = Get-Brush "#111111"
            $button.BorderBrush = Get-Brush "#E5E7EB"
            $button.BorderThickness = [System.Windows.Thickness]::new(1)
        }
        if ($task.Description) {
            $button.ToolTip = $task.Description
        }
        [System.Windows.Controls.Grid]::SetColumn($button, 0)
        [void]$row.Children.Add($button)

        $statusBorder = New-Object System.Windows.Controls.Border
        $statusBorder.Margin = [System.Windows.Thickness]::new(12, 0, 0, 0)
        $statusBorder.CornerRadius = [System.Windows.CornerRadius]::new(12)
        $statusBorder.Background = Get-Brush "#E5E7EB"
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

        $taskKey = $task.Key
        $button.Add_Click({
            param($sender, $e)
            try {
                Start-TaskExecution -TaskKey $taskKey
            }
            catch {
                Append-LogLine ("Launcher error: " + $_.Exception.Message)
                Set-Status -TaskText "Launcher error" -StatusText $_.Exception.Message
                Set-ControlsBusyState -Busy $false
            }
        }.GetNewClosure())

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

    if ($script:IsBusy) { return }

    if (-not $script:TasksByKey.ContainsKey($TaskKey)) {
        [System.Windows.MessageBox]::Show("Unknown task: $TaskKey", "Rutherford Assistant") | Out-Null
        return
    }

    $task = $script:TasksByKey[$TaskKey]

    if (-not (Test-Path $task.ScriptPath)) {
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

    $process.add_OutputDataReceived({
        param($sender, $eventArgs)
        if ($null -ne $eventArgs.Data) {
            try {
                $window.Dispatcher.Invoke([action]{ Append-LogLine $eventArgs.Data })
            }
            catch { }
        }
    })

    $process.add_ErrorDataReceived({
        param($sender, $eventArgs)
        if ($null -ne $eventArgs.Data) {
            try {
                $window.Dispatcher.Invoke([action]{ Append-LogLine ("ERROR: " + $eventArgs.Data) })
            }
            catch { }
        }
    })

    $process.add_Exited({
        param($sender, $eventArgs)

        try {
            $window.Dispatcher.Invoke([action]{
                try {
                    $exitCode = -1
                    try { $exitCode = $sender.ExitCode } catch { }
                    Append-LogLine ("Process finished with exit code " + $exitCode)

                    $finalState = if ($exitCode -eq 0) { "Done" } else { "Error" }
                    $taskKey = $script:CurrentTaskKey
                    $taskRef = $script:CurrentTask

                    Set-ActionState -Key $taskKey -State $finalState
                    Update-ScriptState -Key $taskKey -Status $finalState -ExitCode $exitCode

                    $reportPath = Write-RunReport -TaskName $taskRef.Label -ScriptPath $taskRef.ScriptPath -ExitCode $exitCode
                    if ($reportPath) {
                        Set-Status -TaskText $taskRef.Label -StatusText ("$finalState. Report saved to $reportPath")
                    }
                    else {
                        Set-Status -TaskText $taskRef.Label -StatusText "$finalState. (No report - see logs)"
                    }

                    Refresh-NetworkCards
                    if ($taskRef.AuditAfterRun) {
                        Refresh-AuditPanel
                    }

                    if ($script:State.lastUpdated) {
                        $lastUpdatedText.Text = "Last update: $($script:State.lastUpdated)"
                    }
                }
                catch {
                    Append-LogLine ("Launcher post-run error: " + $_.Exception.Message)
                    Set-Status -TaskText "Launcher error" -StatusText $_.Exception.Message
                }
                finally {
                    $script:CurrentProcess = $null
                    Set-ControlsBusyState -Busy $false
                }
            })
        }
        catch {
            # If the dispatcher invoke itself fails (window already closed) - swallow silently
        }
    })

    try {
        $started = $process.Start()
    }
    catch {
        $started = $false
        Append-LogLine ("Process start failed: " + $_.Exception.Message)
    }

    if (-not $started) {
        Set-ControlsBusyState -Busy $false
        Set-ActionState -Key $TaskKey -State "Error"
        [System.Windows.MessageBox]::Show("Unable to start $($task.ScriptPath)", "Rutherford Assistant") | Out-Null
        return
    }

    $script:CurrentProcess = $process
    try { $process.BeginOutputReadLine() } catch { Append-LogLine ("BeginOutputReadLine failed: " + $_.Exception.Message) }
    try { $process.BeginErrorReadLine() }  catch { Append-LogLine ("BeginErrorReadLine failed: " + $_.Exception.Message) }
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

    if ($script:IsBusy -and $script:CurrentProcess -and -not $script:CurrentProcess.HasExited) {
        $result = [System.Windows.MessageBox]::Show(
            "A task is still running. Close anyway?",
            "Rutherford Assistant",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )

        if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
            $eventArgs.Cancel = $true
        }
    }
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
    Append-LogLine "Launcher ready. Tasks discovered: $($script:Tasks.Count). Audit checks: $($script:AuditChecks.Count)."
})

[void]$window.ShowDialog()
