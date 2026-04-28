@{
    Id       = "wallpaper"
    Label    = "Wallpaper deployed"
    Category = "Appearance"
    Order    = 10
    Test     = {
        $deployedPath = "C:\ProgramData\Rutherford\wallpaper.jpg"
        if (-not (Test-Path $deployedPath)) {
            return @{ Status = "missing"; Detail = "Wallpaper not present in C:\ProgramData\Rutherford" }
        }

        $registryPath = "HKCU:\Control Panel\Desktop"
        try {
            $current = (Get-ItemProperty -Path $registryPath -Name Wallpaper -ErrorAction Stop).Wallpaper
        }
        catch {
            return @{ Status = "partial"; Detail = "File deployed but registry not set" }
        }

        if ($current -eq $deployedPath) {
            return @{ Status = "ok"; Detail = "Wallpaper file deployed and registry points to it" }
        }

        return @{ Status = "partial"; Detail = "File deployed but registry points elsewhere: $current" }
    }
}
