Write-Host "Network configuration script"

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

function Get-AdapterByNameList {
    param([string[]]$Names)

    foreach ($name in $Names) {
        if ([string]::IsNullOrWhiteSpace($name)) { continue }

        $adapter = Get-NetAdapter -Name $name -ErrorAction SilentlyContinue
        if ($adapter) { return $adapter }
    }

    return $null
}

function Rename-AdapterIfNeeded {
    param(
        [string[]]$PossibleCurrentNames,
        [string]$NewName
    )

    $adapter = Get-AdapterByNameList -Names (@($NewName) + $PossibleCurrentNames)

    if (-not $adapter) {
        Write-Host "Adapter not found. Expected: $($PossibleCurrentNames -join ', ') or $NewName"
        return $null
    }

    if ($adapter.Name -eq $NewName) {
        Write-Host "Already named: $NewName"
        return $adapter
    }

    $existingTarget = Get-NetAdapter -Name $NewName -ErrorAction SilentlyContinue
    if ($existingTarget) {
        throw "Name already used: $NewName"
    }

    Rename-NetAdapter -Name $adapter.Name -NewName $NewName -ErrorAction Stop
    Start-Sleep -Seconds 2

    return Get-NetAdapter -Name $NewName -ErrorAction Stop
}

function Set-StaticIPv4 {
    param(
        [string]$InterfaceName,
        [string]$IPv4,
        [string]$SubnetMask,
        [string]$DefaultGateway = ""
    )

    if ([string]::IsNullOrWhiteSpace($IPv4) -or [string]::IsNullOrWhiteSpace($SubnetMask)) {
        Write-Host "No IP config for $InterfaceName"
        return
    }

    $existing = Get-NetIPAddress -InterfaceAlias $InterfaceName -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.IPAddress -eq $IPv4 }

    if ($existing) {
        Write-Host "IP already configured on $InterfaceName"
        return
    }

    Write-Host "Setting IP on $InterfaceName"

    if ([string]::IsNullOrWhiteSpace($DefaultGateway)) {
        $args = @(
            "interface", "ipv4", "set", "address",
            "name=$InterfaceName", "static", $IPv4, $SubnetMask, "none"
        )
    }
    else {
        $args = @(
            "interface", "ipv4", "set", "address",
            "name=$InterfaceName", "static", $IPv4, $SubnetMask, $DefaultGateway
        )
    }

    $result = & netsh @args 2>&1

    if ($LASTEXITCODE -ne 0) {
        throw "netsh error: $result"
    }

    Write-Host "Configured: $InterfaceName -> $IPv4 / $SubnetMask"
}

function Set-NetworkAdapterProfile {
    param(
        [string[]]$PossibleCurrentNames,
        [string]$NewName,
        [string]$IPv4 = "",
        [string]$SubnetMask = "",
        [string]$DefaultGateway = ""
    )

    Write-Step "Configuring $NewName"

    $adapter = Rename-AdapterIfNeeded `
        -PossibleCurrentNames $PossibleCurrentNames `
        -NewName $NewName

    if (-not $adapter) { return }

    if ($IPv4 -eq "DHCP") {
        Write-Host "Setting DHCP on $NewName"
        & netsh interface ipv4 set address name="$NewName" source=dhcp | Out-Null
        return
    }

    Set-StaticIPv4 `
        -InterfaceName $NewName `
        -IPv4 $IPv4 `
        -SubnetMask $SubnetMask `
        -DefaultGateway $DefaultGateway
}

if (-not (Test-IsAdmin)) {
    Write-Host "Run as Administrator"
    exit 1
}

# ===== CONFIG =====

$networkProfiles = @(
    @{
        PossibleCurrentNames = @("Ethernet", "IntelliTrax")
        NewName              = "IntelliTrax"
        IPv4                 = "172.16.1.15"
        SubnetMask           = "255.255.0.0"
    },
    @{
        PossibleCurrentNames = @("Ethernet 2", "Network")
        NewName              = "Network"
        IPv4                 = "DHCP"
    },
    @{
        PossibleCurrentNames = @("Ethernet 3", "Ethernet 4", "PUPI")
        NewName              = "PUPI"
        IPv4                 = "10.10.6.15"
        SubnetMask           = "255.255.255.0"
    }
)

# ===== DEBUG =====

Write-Step "Adapters detected"
Get-NetAdapter | Format-Table -AutoSize Name, InterfaceDescription, Status, MacAddress, LinkSpeed

Write-Step "Starting config"

foreach ($profile in $networkProfiles) {
    try {
        Set-NetworkAdapterProfile @profile
    }
    catch {
        Write-Host "ERROR while configuring $($profile.NewName): $($_.Exception.Message)"
    }
}

Write-Step "Final result"
Get-NetAdapter | Format-Table -AutoSize Name, InterfaceDescription, Status, MacAddress, LinkSpeed

Write-Host "`n=== IP CONFIG ==="
ipconfig /all