@{
    Id       = "startup-cleanup"
    Label    = "Startup apps cleaned"
    Category = "Cleanup"
    Order    = 97
    Test     = {
        $patterns = @("Teams", "Spotify", "OneDrive", "Copilot")
        $registryPaths = @(
            "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run",
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"
        )

        $remaining = @()

        foreach ($regPath in $registryPaths) {
            if (-not (Test-Path $regPath)) { continue }
            $values = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
            if (-not $values) { continue }
            foreach ($prop in $values.PSObject.Properties) {
                if ($prop.Name -in @('PSPath','PSParentPath','PSChildName','PSDrive','PSProvider')) { continue }
                foreach ($pat in $patterns) {
                    if ($prop.Name -like "*$pat*" -or [string]$prop.Value -like "*$pat*") {
                        $remaining += "$pat ($($prop.Name))"
                        break
                    }
                }
            }
        }

        if ($remaining.Count -eq 0) {
            return @{ Status = "ok"; Detail = "No tracked startup entries remaining" }
        }

        return @{ Status = "partial"; Detail = "Still present: $($remaining -join ', ')" }
    }
}
