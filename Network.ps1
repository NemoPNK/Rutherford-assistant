Write-Host "Network configuration script"
Write-Host "Please be patient, this can take some time while Windows applies network changes."

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

function Wait-AdapterReady {
    param(
        [string]$Name,
        [int]$TimeoutSeconds = 20
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)

    do {
        $adapter = Get-NetAdapter -Name $Name -ErrorAction SilentlyContinue
        if ($adapter) {
            return $adapter
        }

        Start-Sleep -Seconds 1
    } while ((Get-Date) -lt $deadline)

    throw "Adapter $Name was not ready after $TimeoutSeconds seconds."
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
    Start-Sleep -Seconds 3

    Start-Sleep -Seconds 2
    return Wait-AdapterReady -Name $NewName -TimeoutSeconds 10
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

    Wait-AdapterReady -Name $InterfaceName -TimeoutSeconds 20 | Out-Null

    $currentIPv4 = Get-NetIPAddress -InterfaceAlias $InterfaceName -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.IPAddress -notlike "169.254.*" }

    $existing = $currentIPv4 | Where-Object { $_.IPAddress -eq $IPv4 }

    if ($existing) {
        Write-Host "IP already configured on $InterfaceName"
        return
    }

    Write-Host "Setting IP on $InterfaceName"

    # Force adapter into static mode before applying the address.
    # This avoids the first-run bug where Windows keeps DHCP active right after rename.
    Write-Host "Applying static IPv4 configuration, please wait..."
    & netsh interface ipv4 set address name="$InterfaceName" source=static addr=$IPv4 mask=$SubnetMask gateway=none | Out-Null
    Start-Sleep -Seconds 2

    if (-not [string]::IsNullOrWhiteSpace($DefaultGateway)) {
        $result = & netsh interface ipv4 set address name="$InterfaceName" static $IPv4 $SubnetMask $DefaultGateway 2>&1

        if ($LASTEXITCODE -ne 0) {
            throw "netsh error: $result"
        }
    }

    $finalIPv4 = Get-NetIPAddress -InterfaceAlias $InterfaceName -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.IPAddress -eq $IPv4 }

    if (-not $finalIPv4) {
        throw "IPv4 was not applied on $InterfaceName. Expected $IPv4"
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
        Write-Host "$NewName stays in DHCP automatic mode."
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
        PossibleCurrentNames = @("Ethernet 3", "PUPI")
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