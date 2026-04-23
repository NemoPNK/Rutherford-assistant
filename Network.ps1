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
    param([Parameter(Mandatory=$true)][string]$SubnetMask)

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
        [Parameter(Mandatory=$true)][string]$CurrentName,
        [Parameter(Mandatory=$true)][string]$NewName,
        [Parameter(Mandatory=$true)][string]$IPv4,
        [Parameter(Mandatory=$true)][string]$SubnetMask,
        [string]$DefaultGateway = ""
    )

    Write-Step "Configuring adapter $CurrentName"

    $adapter = Get-NetAdapter -Name $CurrentName -ErrorAction SilentlyContinue
    if (-not $adapter) {
        throw "Network adapter not found: $CurrentName"
    }

    if ($CurrentName -ne $NewName) {
        Rename-NetAdapter -Name $CurrentName -NewName $NewName -ErrorAction Stop
        Start-Sleep -Seconds 2
    }

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

    Write-Host "Adapter configured: $NewName | IP: $IPv4 | Mask: $SubnetMask"
}

if (-not (Test-IsAdmin)) {
    Write-Host "This script must be run as Administrator."
    exit 1
}

# =====================================================================
# TEMPLATE A REMPLIR
# Remplace les valeurs ci-dessous quand tu auras les vraies infos.
# =====================================================================

$networkProfiles = @(
    @{
        CurrentName    = "Ethernet"
        NewName        = "PORT-1"
        IPv4           = "192.168.1.10"
        SubnetMask     = "255.255.255.0"
        DefaultGateway = ""
    },
    @{
        CurrentName    = "Ethernet 2"
        NewName        = "PORT-2"
        IPv4           = "192.168.1.11"
        SubnetMask     = "255.255.255.0"
        DefaultGateway = ""
    }
)

Write-Step "Starting network configuration"

foreach ($profile in $networkProfiles) {
    Set-NetworkAdapterProfile `
        -CurrentName $profile.CurrentName `
        -NewName $profile.NewName `
        -IPv4 $profile.IPv4 `
        -SubnetMask $profile.SubnetMask `
        -DefaultGateway $profile.DefaultGateway
}

Write-Step "Summary"
Write-Host "Network adapters configured successfully."
Write-Host "Check adapter names with: Get-NetAdapter"

Write-Step "Network adapters summary"

Get-NetAdapter | Format-Table -AutoSize Name, InterfaceDescription, Status, MacAddress, LinkSpeed

Write-Host ""
Write-Host "IPv4 configuration:"
Get-NetIPConfiguration | Format-List InterfaceAlias, InterfaceDescription, IPv4Address, IPv4DefaultGateway, DNSServer

Write-Host ""
Write-Host "Detailed adapter info:"
Get-NetAdapter | ForEach-Object {
    Write-Host "----------------------------------------"
    Get-NetIPAddress -InterfaceAlias $_.Name -AddressFamily IPv4 -ErrorAction SilentlyContinue | Format-Table -AutoSize InterfaceAlias, IPAddress, PrefixLength
}