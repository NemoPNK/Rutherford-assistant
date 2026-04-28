@{
    Id       = "appx-cleanup"
    Label    = "Unwanted Appx removed"
    Category = "Cleanup"
    Order    = 95
    Test     = {
        $unwanted = @(
            "Microsoft.MicrosoftSolitaireCollection",
            "Microsoft.XboxGamingOverlay",
            "Microsoft.GamingApp",
            "Microsoft.BingNews",
            "Microsoft.Weather",
            "Microsoft.WindowsMaps",
            "Microsoft.GetHelp",
            "Microsoft.Getstarted",
            "Microsoft.WindowsFeedbackHub",
            "SpotifyAB.SpotifyMusic",
            "Clipchamp.Clipchamp",
            "Microsoft.ZuneMusic",
            "Microsoft.ZuneVideo",
            "MSTeams",
            "Microsoft.MicrosoftOfficeHub",
            "Microsoft.Todos"
        )

        try {
            $stillPresent = @()
            foreach ($name in $unwanted) {
                $pkg = Get-AppxPackage -Name $name -AllUsers -ErrorAction SilentlyContinue
                if ($pkg) { $stillPresent += $name }
            }

            if ($stillPresent.Count -eq 0) {
                return @{ Status = "ok"; Detail = "All $($unwanted.Count) tracked Appx removed" }
            }

            $removedCount = $unwanted.Count - $stillPresent.Count
            if ($removedCount -eq 0) {
                return @{ Status = "missing"; Detail = "No tracked Appx removed yet" }
            }

            return @{ Status = "partial"; Detail = "$removedCount/$($unwanted.Count) removed ; still present: $($stillPresent -join ', ')" }
        }
        catch {
            return @{ Status = "unknown"; Detail = "Unable to query Appx: $($_.Exception.Message)" }
        }
    }
}
