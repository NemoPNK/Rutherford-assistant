@{
    Id       = "standby"
    Label    = "Standby and hibernate disabled"
    Category = "Power"
    Order    = 70
    Test     = {
        try {
            $output = & powercfg /q SCHEME_CURRENT SUB_SLEEP STANDBYIDLE 2>&1
            $acLine = ($output | Where-Object { $_ -match "Current AC Power Setting Index" } | Select-Object -First 1)
            $dcLine = ($output | Where-Object { $_ -match "Current DC Power Setting Index" } | Select-Object -First 1)

            $acIsZero = $acLine -and ($acLine -match "0x00000000")
            $dcIsZero = $dcLine -and ($dcLine -match "0x00000000")

            $hibernateOn = $false
            try {
                $hib = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Power" -Name HibernateEnabled -ErrorAction Stop).HibernateEnabled
                if ($hib -ne 0) { $hibernateOn = $true }
            } catch { }

            $issues = @()
            if (-not $acIsZero) { $issues += "AC standby not zero" }
            if (-not $dcIsZero) { $issues += "DC standby not zero" }
            if ($hibernateOn) { $issues += "hibernate still enabled" }

            if ($issues.Count -eq 0) {
                return @{ Status = "ok"; Detail = "Standby AC/DC = never, hibernate disabled" }
            }

            return @{ Status = "partial"; Detail = ($issues -join " ; ") }
        }
        catch {
            return @{ Status = "unknown"; Detail = "Unable to read power config: $($_.Exception.Message)" }
        }
    }
}
