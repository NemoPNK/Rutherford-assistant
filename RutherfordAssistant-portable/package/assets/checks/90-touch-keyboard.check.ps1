@{
    Id       = "touch-keyboard"
    Label    = "Touch keyboard configured"
    Category = "Input"
    Order    = 90
    Test     = {
        $checks = @(
            @{ Path = "HKCU:\Software\Microsoft\TabletTip\1.7";                                       Name = "EnableDesktopModeAutoInvoke";  Want = 1 }
            @{ Path = "HKCU:\Software\Microsoft\TabletTip\1.7";                                       Name = "TipbandDesiredVisibility";     Want = 1 }
            @{ Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced";            Name = "ShowTouchKeyboardButton";      Want = 1 }
        )

        $okCount = 0
        $missing = @()

        foreach ($entry in $checks) {
            try {
                $value = (Get-ItemProperty -Path $entry.Path -Name $entry.Name -ErrorAction Stop).$($entry.Name)
                if ($value -eq $entry.Want) { $okCount++ }
                else { $missing += "$($entry.Name)=$value" }
            }
            catch {
                $missing += $entry.Name
            }
        }

        if ($missing.Count -eq 0) {
            return @{ Status = "ok"; Detail = "$okCount/$($checks.Count) keys configured" }
        }

        if ($okCount -eq 0) {
            return @{ Status = "missing"; Detail = "Touch keyboard not configured" }
        }

        return @{ Status = "partial"; Detail = "$okCount/$($checks.Count) ; missing: $($missing -join ', ')" }
    }
}
