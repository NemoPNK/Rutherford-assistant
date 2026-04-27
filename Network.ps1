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

function Convert-MaskToPrefixLength {
    param([string]$SubnetMask)

    if ([string]::IsNullOrWhiteSpace($SubnetMask)) { return $null }

    $octets = $SubnetMask.Split('.')
    if ($octets.Count -ne 4) {
        throw "Invalid subnet mask format: $SubnetMask"
    }

    $prefixLength = 0
    foreach ($octet in $octets) {
        switch ([int]$octet) {
            255 { $prefixLength += 8 }
            254 { $prefixLength += 7 }
            252 { $prefixLength += 6 }
            248 { $prefixLength += 5 }
            240 { $prefixLength += 4 }
            224 { $prefixLength += 3 }
            192 { $prefixLength += 2 }
            128 { $prefixLength += 1 }
            0   { $prefixLength += 0 }
            default { throw "Invalid subnet mask value: $SubnetMask" }
        }
    }

    return $prefixLength
}

function Set-NetworkAdapterProfile {
    param(
        [string]$CurrentName,
        [string]$NewName,
        [string]$IPv4 = "",
        [string]$SubnetMask = "",
        [string]$DefaultGateway = ""
    )

    Write-Step "Configuring adapter $CurrentName"

    $adapter = Get-NetAdapter -Name $CurrentName -ErrorAction SilentlyContinue
    if (-not $adapter) {
        Write-Host "Adapter not found: $CurrentName"
        return
    }

    if ($CurrentName -ne $NewName) {
        Rename-NetAdapter -Name $CurrentName -NewName $NewName -ErrorAction Stop
        Start-Sleep -Seconds 2
    }

    if (-not [string]::IsNullOrWhiteSpace($IPv4) -and -not [string]::IsNullOrWhiteSpace($SubnetMask)) {
        $prefixLength = Convert-MaskToPrefixLength -SubnetMask $SubnetMask

        Get-NetIPAddress -InterfaceAlias $NewName -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Remove-NetIPAddress -Confirm:$false -ErrorAction SilentlyContinue

        Get-NetRoute -InterfaceAlias $NewName -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Remove-NetRoute -Confirm:$false -ErrorAction SilentlyContinue

        $ipParams = @{
            InterfaceAlias = $NewName
            IPAddress      = $IPv4
            PrefixLength   = $prefixLength
            AddressFamily  = "IPv4"
        }

        if (-not [string]::IsNullOrWhiteSpace($DefaultGateway)) {
            $ipParams.DefaultGateway = $DefaultGateway
        }

        New-NetIPAddress @ipParams -ErrorAction Stop | Out-Null

        Write-Host "Configured: $NewName | IP: $IPv4 | Mask: $SubnetMask"
    }
    else {
        Write-Host "Renamed only: $NewName"
    }
}

if (-not (Test-IsAdmin)) {
    Write-Host "This script must be run as Administrator."
    exit 1
}

$networkProfiles = @(
    @{
        CurrentName = "Ethernet"
        NewName     = "IntelliTrax"
        IPv4        = "172.16.1.15"
        SubnetMask  = "255.255.0.0"
    },
    @{
        CurrentName = "Ethernet 2"
        NewName     = "Network"
    },
    @{
        CurrentName = "Ethernet 3"
        NewName     = "PUPI"
        IPv4        = "10.10.6.15"
        SubnetMask  = "255.255.255.0"
    }
)

Write-Step "Starting network configuration"

foreach ($profile in $networkProfiles) {
    Set-NetworkAdapterProfile @profile
}

Write-Step "Network adapters summary"
Get-NetAdapter | Format-Table -AutoSize Name, Status, MacAddress, LinkSpeed

Write-Host "`nIPv4 configuration:"
Get-NetIPConfiguration | Format-List InterfaceAlias, IPv4Address, IPv4DefaultGateway, DNSServer