@{
    Id       = "onedrive"
    Label    = "OneDrive removed"
    Category = "Cleanup"
    Order    = 20
    Test     = {
        $candidates = @(
            "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe",
            "$env:PROGRAMFILES\Microsoft OneDrive\OneDrive.exe",
            "${env:PROGRAMFILES(X86)}\Microsoft OneDrive\OneDrive.exe"
        )

        $foundPaths = $candidates | Where-Object { $_ -and (Test-Path $_) }

        $running = $null
        try { $running = Get-Process -Name OneDrive -ErrorAction SilentlyContinue } catch { }

        if (-not $foundPaths -and -not $running) {
            return @{ Status = "ok"; Detail = "OneDrive executable not found and no process running" }
        }

        if ($running) {
            return @{ Status = "missing"; Detail = "OneDrive process is running" }
        }

        return @{ Status = "partial"; Detail = "OneDrive binary still present: $($foundPaths -join ', ')" }
    }
}
