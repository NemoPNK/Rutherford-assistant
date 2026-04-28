Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase

$script:LauncherPath = $PSCommandPath
$script:AssetsRoot = Split-Path -Parent $script:LauncherPath
$script:AppRoot = Split-Path -Parent $script:AssetsRoot
$script:ReportsRoot = Join-Path $script:AppRoot "reports"
$script:IsBusy = $false
$script:CurrentProcess = $null
$script:CurrentTask = $null
$script:LastReportPath = $null
$script:CurrentLogLines = New-Object System.Collections.Generic.List[string]
$script:RunStartedAt = $null

function Test-IsAdmin {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Ensure-Elevated {
    if (Test-IsAdmin) {
        return
    }

    $argumentList = "-NoProfile -ExecutionPolicy Bypass -File `"$script:LauncherPath`""

    Start-Process -FilePath "powershell.exe" -ArgumentList $argumentList -Verb RunAs | Out-Null
    exit
}

Ensure-Elevated

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Rutherford Assistant"
        Width="1080"
        Height="760"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize"
        Background="#F4EFE5"
        FontFamily="Segoe UI">
  <Grid Margin="18">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto" />
      <RowDefinition Height="16" />
      <RowDefinition Height="*" />
      <RowDefinition Height="16" />
      <RowDefinition Height="Auto" />
    </Grid.RowDefinitions>

    <Border Grid.Row="0"
            Background="#FFF8EC"
            BorderBrush="#E7D8BE"
            BorderThickness="1"
            CornerRadius="18"
            Padding="20">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*" />
          <ColumnDefinition Width="320" />
        </Grid.ColumnDefinitions>

        <StackPanel Grid.Column="0">
          <TextBlock Text="Rutherford Assistant"
                     FontSize="30"
                     FontWeight="Bold"
                     Foreground="#1F1A14" />
          <TextBlock Margin="0,10,0,0"
                     Foreground="#6A5C4C"
                     FontSize="14"
                     TextWrapping="Wrap"
                     Text="Portable Windows launcher for USB use. Keep this window open while scripts run." />
        </StackPanel>

        <Border Grid.Column="1"
                Background="#FFFFFF"
                CornerRadius="14"
                Padding="14"
                BorderBrush="#E7D8BE"
                BorderThickness="1">
          <StackPanel>
            <TextBlock Text="Operator Flow"
                       FontWeight="Bold"
                       Foreground="#7C2D12" />
            <TextBlock Margin="0,10,0,0"
                       TextWrapping="Wrap"
                       Foreground="#6A5C4C"
                       Text="1. Click Setup or Network.&#x0a;2. Accept admin elevation once.&#x0a;3. Follow logs here.&#x0a;4. Open the final report if needed." />
          </StackPanel>
        </Border>
      </Grid>
    </Border>

    <Grid Grid.Row="2">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="330" />
        <ColumnDefinition Width="16" />
        <ColumnDefinition Width="*" />
      </Grid.ColumnDefinitions>

      <Border Grid.Column="0"
              Background="#FFFCF7"
              BorderBrush="#E7D8BE"
              BorderThickness="1"
              CornerRadius="18"
              Padding="18">
        <StackPanel>
          <TextBlock Text="Actions"
                     FontWeight="Bold"
                     Foreground="#7C2D12"
                     FontSize="18" />

          <Button Name="SetupButton"
                  Margin="0,14,0,0"
                  Height="52"
                  FontWeight="Bold"
                  Background="#9A3412"
                  Foreground="White"
                  BorderThickness="0"
                  Content="Run Setup" />

          <Button Name="NetworkButton"
                  Margin="0,12,0,0"
                  Height="52"
                  FontWeight="Bold"
                  Background="#FFF1DC"
                  Foreground="#7C2D12"
                  BorderBrush="#E7D8BE"
                  BorderThickness="1"
                  Content="Run Network" />

          <Button Name="UpdatesButton"
                  Margin="0,12,0,0"
                  Height="52"
                  FontWeight="Bold"
                  Background="#F1E9DE"
                  Foreground="#9B8A75"
                  BorderBrush="#E7D8BE"
                  BorderThickness="1"
                  IsEnabled="False"
                  Content="Run Updates (Coming Soon)" />

          <Button Name="ClearLogsButton"
                  Margin="0,24,0,0"
                  Height="40"
                  FontWeight="Bold"
                  Background="#FFFFFF"
                  Foreground="#6A5C4C"
                  BorderBrush="#E7D8BE"
                  BorderThickness="1"
                  Content="Clear Logs" />

          <Button Name="OpenReportButton"
                  Margin="0,12,0,0"
                  Height="40"
                  FontWeight="Bold"
                  Background="#FFFFFF"
                  Foreground="#6A5C4C"
                  BorderBrush="#E7D8BE"
                  BorderThickness="1"
                  IsEnabled="False"
                  Content="Open Last Report" />

          <TextBlock Margin="0,24,0,0"
                     Text="Files"
                     FontWeight="Bold"
                     Foreground="#7C2D12" />
          <TextBlock Name="FilesSummaryText"
                     Margin="0,8,0,0"
                     Foreground="#6A5C4C"
                     TextWrapping="Wrap"
                     Text="Setup: LaRoche.ps1&#x0a;Network: Network.ps1&#x0a;Future: update.ps1" />
        </StackPanel>
      </Border>

      <Border Grid.Column="2"
              Background="#FFFCF7"
              BorderBrush="#E7D8BE"
              BorderThickness="1"
              CornerRadius="18"
              Padding="18">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="16" />
            <RowDefinition Height="*" />
          </Grid.RowDefinitions>

          <Grid Grid.Row="0">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*" />
              <ColumnDefinition Width="260" />
            </Grid.ColumnDefinitions>

            <StackPanel Grid.Column="0">
              <TextBlock Text="Execution"
                         FontWeight="Bold"
                         Foreground="#7C2D12"
                         FontSize="18" />
              <TextBlock Name="CurrentTaskText"
                         Margin="0,8,0,0"
                         FontSize="22"
                         FontWeight="SemiBold"
                         Foreground="#1F1A14"
                         Text="Ready" />
              <TextBlock Name="CurrentStatusText"
                         Margin="0,6,0,0"
                         Foreground="#6A5C4C"
                         Text="Waiting for action." />
            </StackPanel>

            <Border Grid.Column="1"
                    Background="#F3EEE6"
                    CornerRadius="14"
                    Padding="14"
                    BorderBrush="#E7D8BE"
                    BorderThickness="1">
              <StackPanel>
                <TextBlock Text="Report"
                           FontWeight="Bold"
                           Foreground="#7C2D12" />
                <TextBlock Name="ReportSummaryText"
                           Margin="0,8,0,0"
                           Foreground="#6A5C4C"
                           TextWrapping="Wrap"
                           Text="No report yet." />
              </StackPanel>
            </Border>
          </Grid>

          <Border Grid.Row="2"
                  Background="#1F1F1F"
                  CornerRadius="16"
                  Padding="10">
            <TextBox Name="LogsTextBox"
                     Background="#1F1F1F"
                     Foreground="#F4EFE5"
                     BorderThickness="0"
                     FontFamily="Consolas"
                     FontSize="13"
                     VerticalScrollBarVisibility="Auto"
                     HorizontalScrollBarVisibility="Auto"
                     AcceptsReturn="True"
                     AcceptsTab="True"
                     IsReadOnly="True"
                     TextWrapping="NoWrap" />
          </Border>
        </Grid>
      </Border>
    </Grid>

    <Border Grid.Row="4"
            Margin="0,16,0,0"
            Background="#FFF8EC"
            BorderBrush="#E7D8BE"
            BorderThickness="1"
            CornerRadius="18"
            Padding="16">
      <TextBlock Foreground="#6A5C4C"
                 TextWrapping="Wrap"
                 Text="This launcher stays open until you close it. Run one action at a time. Reports are saved in the reports folder next to LaRocheLauncher.bat." />
    </Border>
  </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$setupButton = $window.FindName("SetupButton")
$networkButton = $window.FindName("NetworkButton")
$updatesButton = $window.FindName("UpdatesButton")
$clearLogsButton = $window.FindName("ClearLogsButton")
$openReportButton = $window.FindName("OpenReportButton")
$filesSummaryText = $window.FindName("FilesSummaryText")
$currentTaskText = $window.FindName("CurrentTaskText")
$currentStatusText = $window.FindName("CurrentStatusText")
$reportSummaryText = $window.FindName("ReportSummaryText")
$logsTextBox = $window.FindName("LogsTextBox")

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

function Append-LogLine {
    param([string]$Line)

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return
    }

    $timestampedLine = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Line.TrimEnd()
    $script:CurrentLogLines.Add($timestampedLine) | Out-Null

    $null = $logsTextBox.Dispatcher.Invoke([action]{
        $logsTextBox.AppendText($timestampedLine + [Environment]::NewLine)
        $logsTextBox.ScrollToEnd()
    })
}

function Set-ControlsBusyState {
    param([bool]$Busy)

    $script:IsBusy = $Busy
    $setupButton.IsEnabled = -not $Busy
    $networkButton.IsEnabled = -not $Busy
    $updatesButton.IsEnabled = $false
    $clearLogsButton.IsEnabled = -not $Busy
    $openReportButton.IsEnabled = (-not $Busy) -and [bool]$script:LastReportPath
}

function Set-Status {
    param(
        [string]$TaskText,
        [string]$StatusText
    )

    $currentTaskText.Text = $TaskText
    $currentStatusText.Text = $StatusText
}

function Get-ReportFilePath {
    param([string]$TaskName)

    if (-not (Test-Path $script:ReportsRoot)) {
        New-Item -Path $script:ReportsRoot -ItemType Directory -Force | Out-Null
    }

    $safeTask = $TaskName -replace "[^A-Za-z0-9_-]", "_"
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    return Join-Path $script:ReportsRoot "$timestamp-$safeTask-report.txt"
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
    $statusLabel = if ($ExitCode -eq 0) { "Completed" } else { "Error" }
    $lines = @(
        "Rutherford Assistant Report",
        "Generated: $($finishedAt.ToString("s"))",
        "",
        "Task: $TaskName",
        "Script: $ScriptPath",
        "Status: $statusLabel",
        "ExitCode: $ExitCode",
        "Started: $($script:RunStartedAt.ToString("s"))",
        "Finished: $($finishedAt.ToString("s"))",
        "DurationSeconds: $duration",
        "ComputerName: $env:COMPUTERNAME",
        "UserName: $env:USERNAME",
        "",
        "Logs:",
        $script:CurrentLogLines
    )

    Set-Content -Path $reportPath -Value $lines -Encoding UTF8
    $script:LastReportPath = $reportPath
    $reportSummaryText.Text = "Last report:`n$reportPath"
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

    $logsTextBox.Clear()
    $script:CurrentLogLines.Clear()
    $script:LastReportPath = $null
    $reportSummaryText.Text = "Report will be generated when the task finishes."
    $openReportButton.IsEnabled = $false
    $script:CurrentTask = $task
    $script:RunStartedAt = Get-Date

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
            Append-LogLine $eventArgs.Data
        }
    })

    $process.add_ErrorDataReceived({
        param($sender, $eventArgs)
        if ($null -ne $eventArgs.Data) {
            Append-LogLine ("ERROR: " + $eventArgs.Data)
        }
    })

    $process.add_Exited({
        param($sender, $eventArgs)
        $window.Dispatcher.Invoke([action]{
            $exitCode = $sender.ExitCode
            $statusText = if ($exitCode -eq 0) { "Completed successfully." } else { "Failed with exit code $exitCode." }
            Append-LogLine $statusText

            $reportPath = Write-RunReport -TaskName $script:CurrentTask.Label -ScriptPath $script:CurrentTask.Script -ExitCode $exitCode
            Set-ControlsBusyState -Busy $false

            if ($exitCode -eq 0) {
                Set-Status -TaskText $script:CurrentTask.Label -StatusText "Done. Report saved to $reportPath"
            }
            else {
                Set-Status -TaskText $script:CurrentTask.Label -StatusText "Error. Report saved to $reportPath"
            }

            $script:CurrentProcess = $null
        })
    })

    $started = $process.Start()
    if (-not $started) {
        [System.Windows.MessageBox]::Show("Unable to start $($task.Script)", "Rutherford Assistant")
        return
    }

    $script:CurrentProcess = $process
    Set-ControlsBusyState -Busy $true
    Set-Status -TaskText $task.Label -StatusText "Running... keep this window open."
    $process.BeginOutputReadLine()
    $process.BeginErrorReadLine()
}

$setupButton.Add_Click({ Start-TaskExecution -TaskKey "setup" })
$networkButton.Add_Click({ Start-TaskExecution -TaskKey "network" })
$clearLogsButton.Add_Click({
    if ($script:IsBusy) {
        return
    }

    $logsTextBox.Clear()
    $script:CurrentLogLines.Clear()
    Set-Status -TaskText "Ready" -StatusText "Waiting for action."
    $reportSummaryText.Text = if ($script:LastReportPath) { "Last report:`n$script:LastReportPath" } else { "No report yet." }
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

$filesSummaryText.Text = "Setup: LaRoche.ps1`nNetwork: Network.ps1`nFuture: update.ps1"
$window.ShowDialog() | Out-Null
