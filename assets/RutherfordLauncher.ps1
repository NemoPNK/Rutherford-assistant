Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase

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

$script:ReportsRoot = Join-Path $script:AppRoot "reports"
$script:NetworkProfilesPath = Join-Path $script:AssetsRoot "config\network-profiles.json"
$script:IsBusy = $false
$script:CurrentProcess = $null
$script:CurrentTask = $null
$script:LastReportPath = $null
$script:RunStartedAt = $null
$script:CurrentLogLines = New-Object System.Collections.Generic.List[string]
$script:LogItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[string]'
$script:ActionState = @{
    setup   = "Not done"
    network = "Not done"
    updates = "Not done"
}

function Test-IsAdmin {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Ensure-Elevated {
    if (Test-IsAdmin) {
        return
    }

    $argumentList = "-NoProfile -ExecutionPolicy Bypass -STA -File `"$script:LauncherPath`""
    Start-Process -FilePath "powershell.exe" -ArgumentList $argumentList -Verb RunAs | Out-Null
    exit
}

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
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        $adapter = Get-NetAdapter -Name $name -ErrorAction SilentlyContinue
        if ($adapter) {
            return $adapter
        }
    }

    return $null
}

function Convert-PrefixLengthToMask {
    param([int]$PrefixLength)

    if ($PrefixLength -lt 0 -or $PrefixLength -gt 32) {
        return ""
    }

    $bits = ("1" * $PrefixLength).PadRight(32, "0")
    $octets = for ($index = 0; $index -lt 4; $index++) {
        [Convert]::ToInt32($bits.Substring($index * 8, 8), 2)
    }

    return ($octets -join ".")
}

function Escape-Html {
    param([string]$Text)

    if ($null -eq $Text) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Get-LineSeverity {
    param([string]$Line)

    if ($Line -match "(?i)\b(ERROR|Failed|Exception|not found|can't|cant found)\b") {
        return "error"
    }

    if ($Line -match "(?i)\b(completed|configured|processed|copied|added|installed|removed|set|already|ready|success)\b") {
        return "success"
    }

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

        if (-not $currentStep) {
            continue
        }

        if (Get-LineSeverity $cleanLine -eq "error") {
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

Ensure-Elevated

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Rutherford Assistant"
        Width="1220"
        Height="860"
        MinWidth="1080"
        MinHeight="760"
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
            <TextBlock Margin="0,10,0,0"
                       Foreground="#B3B3BC"
                       FontSize="14"
                       TextWrapping="Wrap"
                       Text="Windows launcher designed for a packaged EXE workflow. Keep the interface open while scripts run." />
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
                         Text="1. Open the launcher EXE.&#x0a;2. Run Setup or Network.&#x0a;3. Watch live logs.&#x0a;4. Open the HTML report when the task finishes." />
            </StackPanel>
          </Border>
        </Grid>
      </Border>

      <Grid Grid.Row="2">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="360" />
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

            <Grid Margin="0,16,0,0">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*" />
                <ColumnDefinition Width="120" />
              </Grid.ColumnDefinitions>
              <Button Name="SetupButton"
                      Grid.Column="0"
                      Height="50"
                      FontWeight="Bold"
                      Background="#111111"
                      Foreground="#FFFFFF"
                      BorderThickness="0"
                      Content="Run Setup" />
              <Border Name="SetupStatusBorder"
                      Grid.Column="1"
                      Margin="12,0,0,0"
                      CornerRadius="12"
                      Background="#E5E7EB"
                      Padding="10,0">
                <TextBlock Name="SetupStatusText"
                           VerticalAlignment="Center"
                           HorizontalAlignment="Center"
                           FontWeight="Bold"
                           Text="Not done" />
              </Border>
            </Grid>

            <Grid Margin="0,12,0,0">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*" />
                <ColumnDefinition Width="120" />
              </Grid.ColumnDefinitions>
              <Button Name="NetworkButton"
                      Grid.Column="0"
                      Height="50"
                      FontWeight="Bold"
                      Background="#F3F4F6"
                      Foreground="#111111"
                      BorderBrush="#E5E7EB"
                      BorderThickness="1"
                      Content="Run Network" />
              <Border Name="NetworkStatusBorder"
                      Grid.Column="1"
                      Margin="12,0,0,0"
                      CornerRadius="12"
                      Background="#E5E7EB"
                      Padding="10,0">
                <TextBlock Name="NetworkStatusText"
                           VerticalAlignment="Center"
                           HorizontalAlignment="Center"
                           FontWeight="Bold"
                           Text="Not done" />
              </Border>
            </Grid>

            <Grid Margin="0,12,0,0">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*" />
                <ColumnDefinition Width="120" />
              </Grid.ColumnDefinitions>
              <Button Name="UpdatesButton"
                      Grid.Column="0"
                      Height="50"
                      FontWeight="Bold"
                      Background="#F3F4F6"
                      Foreground="#A1A1AA"
                      BorderBrush="#E5E7EB"
                      BorderThickness="1"
                      IsEnabled="False"
                      Content="Run Updates" />
              <Border Name="UpdatesStatusBorder"
                      Grid.Column="1"
                      Margin="12,0,0,0"
                      CornerRadius="12"
                      Background="#E5E7EB"
                      Padding="10,0">
                <TextBlock Name="UpdatesStatusText"
                           VerticalAlignment="Center"
                           HorizontalAlignment="Center"
                           FontWeight="Bold"
                           Text="Not done" />
              </Border>
            </Grid>

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
          </StackPanel>
        </Border>

        <Grid Grid.Column="2">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="18" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="18" />
            <RowDefinition Height="430" />
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
                  <TextBlock Text="EXE Target"
                             FontWeight="Bold" />
                  <TextBlock Margin="0,8,0,0"
                             Foreground="#B3B3BC"
                             TextWrapping="Wrap"
                             Text="The final goal is a clean Windows EXE launcher that opens this UI and keeps it visible while scripts run." />
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
              <TextBlock Text="Network Cards"
                         Foreground="#111111"
                         FontSize="20"
                         FontWeight="Bold" />
              <TextBlock Margin="0,8,0,0"
                         Foreground="#52525B"
                         Text="Name, IP and mask must match Network.ps1 expectations." />
              <StackPanel Name="NetworkCardsPanel"
                          Margin="0,16,0,0" />
            </StackPanel>
          </Border>

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
                   Text="Reports are generated as HTML with green and red status cards. Network expectations come from assets\config\network-profiles.json so the UI and Network.ps1 stay in sync." />
      </Border>
    </Grid>
  </ScrollViewer>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$setupButton = $window.FindName("SetupButton")
$networkButton = $window.FindName("NetworkButton")
$updatesButton = $window.FindName("UpdatesButton")
$openReportButton = $window.FindName("OpenReportButton")
$refreshNetworkButton = $window.FindName("RefreshNetworkButton")
$clearLogsButton = $window.FindName("ClearLogsButton")
$setupStatusBorder = $window.FindName("SetupStatusBorder")
$setupStatusText = $window.FindName("SetupStatusText")
$networkStatusBorder = $window.FindName("NetworkStatusBorder")
$networkStatusText = $window.FindName("NetworkStatusText")
$updatesStatusBorder = $window.FindName("UpdatesStatusBorder")
$updatesStatusText = $window.FindName("UpdatesStatusText")
$currentTaskText = $window.FindName("CurrentTaskText")
$currentStatusText = $window.FindName("CurrentStatusText")
$reportSummaryText = $window.FindName("ReportSummaryText")
$logsListBox = $window.FindName("LogsListBox")
$networkCardsPanel = $window.FindName("NetworkCardsPanel")

$logsListBox.ItemsSource = $script:LogItems

$taskMap = @{
    setup = @{
        Label  = "Setup Rutherford"
        Script = Join-Path $script:AssetsRoot "LaRoche.ps1"
    }
    network = @{
        Label  = "Network Rutherford"
        Script = Join-Path $script:AssetsRoot "Network.ps1"
    }
}

function Set-Status {
    param(
        [string]$TaskText,
        [string]$StatusText
    )

    $currentTaskText.Text = $TaskText
    $currentStatusText.Text = $StatusText
}

function Set-ActionState {
    param(
        [string]$TaskKey,
        [string]$State
    )

    $script:ActionState[$TaskKey] = $State

    switch ($TaskKey) {
        "setup" {
            $targetBorder = $setupStatusBorder
            $targetText = $setupStatusText
        }
        "network" {
            $targetBorder = $networkStatusBorder
            $targetText = $networkStatusText
        }
        "updates" {
            $targetBorder = $updatesStatusBorder
            $targetText = $updatesStatusText
        }
        default { return }
    }

    switch ($State) {
        "Running" {
            $targetBorder.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FEF3C7")
            $targetText.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#92400E")
        }
        "Done" {
            $targetBorder.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#DCFCE7")
            $targetText.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#166534")
        }
        "Error" {
            $targetBorder.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FEE2E2")
            $targetText.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#B91C1C")
        }
        default {
            $targetBorder.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#E5E7EB")
            $targetText.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#374151")
        }
    }

    $targetText.Text = $State
}

function Append-LogLine {
    param([string]$Line)

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return
    }

    $timestampedLine = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Line.TrimEnd()
    $script:CurrentLogLines.Add($timestampedLine) | Out-Null
    $script:LogItems.Add($timestampedLine)
    $logsListBox.ScrollIntoView($timestampedLine)
}

function Set-ControlsBusyState {
    param([bool]$Busy)

    $script:IsBusy = $Busy
    $setupButton.IsEnabled = -not $Busy
    $networkButton.IsEnabled = -not $Busy
    $refreshNetworkButton.IsEnabled = -not $Busy
    $clearLogsButton.IsEnabled = -not $Busy
    $updatesButton.IsEnabled = $false
    $openReportButton.IsEnabled = (-not $Busy) -and [bool]$script:LastReportPath
}

function Add-NetworkCardElement {
    param($Snapshot)

    $border = New-Object System.Windows.Controls.Border
    $border.CornerRadius = [System.Windows.CornerRadius]::new(18)
    $border.BorderThickness = [System.Windows.Thickness]::new(1)
    $border.Padding = [System.Windows.Thickness]::new(14)
    $border.Margin = [System.Windows.Thickness]::new(0, 0, 0, 10)
    $border.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FAFAFA")
    $border.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#E5E7EB")

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
    $nameText.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#111111")
    [System.Windows.Controls.Grid]::SetColumn($nameText, 0)
    [void]$headerGrid.Children.Add($nameText)

    $statusBorder = New-Object System.Windows.Controls.Border
    $statusBorder.CornerRadius = [System.Windows.CornerRadius]::new(10)
    $statusBorder.Padding = [System.Windows.Thickness]::new(10, 4, 10, 4)
    $statusBorder.HorizontalAlignment = "Right"
    if ($Snapshot.Status -eq "success") {
        $statusBorder.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#DCFCE7")
        $statusTextColor = "#166534"
    }
    else {
        $statusBorder.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FEE2E2")
        $statusTextColor = "#B91C1C"
    }

    $statusText = New-Object System.Windows.Controls.TextBlock
    $statusText.Text = $Snapshot.StatusText
    $statusText.FontWeight = "Bold"
    $statusText.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($statusTextColor)
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
        $text.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#52525B")
        [void]$stack.Children.Add($text)
    }

    [void]$networkCardsPanel.Children.Add($border)
}

function Refresh-NetworkCards {
    try {
        $networkCardsPanel.Children.Clear()
        foreach ($snapshot in Get-NetworkSnapshots) {
            Add-NetworkCardElement -Snapshot $snapshot
        }
    }
    catch {
        $networkCardsPanel.Children.Clear()
        $errorText = New-Object System.Windows.Controls.TextBlock
        $errorText.Text = "Unable to read network status: $($_.Exception.Message)"
        $errorText.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#B91C1C")
        [void]$networkCardsPanel.Children.Add($errorText)
    }
}

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
    param(
        [string]$TaskName,
        [string]$ScriptPath,
        [int]$ExitCode
    )

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

function Start-TaskExecution {
    param([string]$TaskKey)

    if ($script:IsBusy) {
        return
    }

    $task = $taskMap[$TaskKey]
    if (-not $task) {
        [System.Windows.MessageBox]::Show("Unknown task: $TaskKey", "Rutherford Assistant")
        return
    }

    if (-not (Test-Path $task.Script)) {
        [System.Windows.MessageBox]::Show("Missing script: $($task.Script)", "Rutherford Assistant")
        return
    }

    $script:LogItems.Clear()
    $script:CurrentLogLines.Clear()
    $script:LastReportPath = $null
    $reportSummaryText.Text = "Report will be generated when the task finishes."
    $openReportButton.IsEnabled = $false
    $script:CurrentTask = $task
    $script:RunStartedAt = Get-Date

    Set-ActionState -TaskKey $TaskKey -State "Running"
    Set-ControlsBusyState -Busy $true
    Set-Status -TaskText $task.Label -StatusText "Running... keep this window open."

    Append-LogLine "Launcher root: $script:AppRoot"
    Append-LogLine "Assets root: $script:AssetsRoot"
    Append-LogLine "Starting task: $($task.Label)"
    Append-LogLine "Script: $($task.Script)"

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = "powershell.exe"
    $startInfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$($task.Script)`""
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
            $window.Dispatcher.Invoke([action]{
                Append-LogLine $eventArgs.Data
            })
        }
    })

    $process.add_ErrorDataReceived({
        param($sender, $eventArgs)
        if ($null -ne $eventArgs.Data) {
            $window.Dispatcher.Invoke([action]{
                Append-LogLine ("ERROR: " + $eventArgs.Data)
            })
        }
    })

    $process.add_Exited({
        param($sender, $eventArgs)

        $window.Dispatcher.Invoke([action]{
            try {
                $exitCode = $sender.ExitCode
                Append-LogLine ("Process finished with exit code " + $exitCode)
                $reportPath = Write-RunReport -TaskName $script:CurrentTask.Label -ScriptPath $script:CurrentTask.Script -ExitCode $exitCode
                $finalState = if ($exitCode -eq 0) { "Done" } else { "Error" }
                Set-ActionState -TaskKey $script:CurrentTask.Label.ToLower().Split(" ")[0] -State $finalState
                Set-ControlsBusyState -Busy $false
                if ($exitCode -eq 0) {
                    Set-Status -TaskText $script:CurrentTask.Label -StatusText "Done. Report saved to $reportPath"
                }
                else {
                    Set-Status -TaskText $script:CurrentTask.Label -StatusText "Error. Report saved to $reportPath"
                }
                Refresh-NetworkCards
                $script:CurrentProcess = $null
            }
            catch {
                Set-ControlsBusyState -Busy $false
                Set-Status -TaskText "Launcher error" -StatusText $_.Exception.Message
                Append-LogLine ("ERROR: " + $_.Exception.Message)
            }
        })
    })

    $started = $process.Start()
    if (-not $started) {
        Set-ControlsBusyState -Busy $false
        Set-ActionState -TaskKey $TaskKey -State "Error"
        [System.Windows.MessageBox]::Show("Unable to start $($task.Script)", "Rutherford Assistant")
        return
    }

    $script:CurrentProcess = $process
    $process.BeginOutputReadLine()
    $process.BeginErrorReadLine()
}

$setupButton.Add_Click({ Start-TaskExecution -TaskKey "setup" })
$networkButton.Add_Click({ Start-TaskExecution -TaskKey "network" })
$refreshNetworkButton.Add_Click({ Refresh-NetworkCards })
$clearLogsButton.Add_Click({
    if ($script:IsBusy) {
        return
    }

    $script:LogItems.Clear()
    $script:CurrentLogLines.Clear()
    Set-Status -TaskText "Ready" -StatusText "Waiting for action."
})

$openReportButton.Add_Click({
    if ($script:LastReportPath -and (Test-Path $script:LastReportPath)) {
        Start-Process -FilePath $script:LastReportPath | Out-Null
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

$networkRefreshTimer = New-Object System.Windows.Threading.DispatcherTimer
$networkRefreshTimer.Interval = [TimeSpan]::FromSeconds(5)
$networkRefreshTimer.Add_Tick({ if (-not $script:IsBusy) { Refresh-NetworkCards } })
$networkRefreshTimer.Start()

Set-ActionState -TaskKey "setup" -State "Not done"
Set-ActionState -TaskKey "network" -State "Not done"
Set-ActionState -TaskKey "updates" -State "Not done"
Refresh-NetworkCards
Set-Status -TaskText "Ready" -StatusText "Waiting for action."
$window.ShowDialog() | Out-Null
