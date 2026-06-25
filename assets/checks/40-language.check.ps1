@{
    Id       = "language"
    Label    = "System language en-US"
    Category = "Locale"
    Order    = 40
    Test     = {
        try {
            $systemLocale = (Get-WinSystemLocale).Name
            $uiOverride   = (Get-WinUILanguageOverride)
            $langList     = Get-WinUserLanguageList
            $userList     = $langList | Select-Object -ExpandProperty LanguageTag
            $inputTips    = @($langList | ForEach-Object { $_.InputMethodTips })

            $issues = @()
            if ($systemLocale -ne "en-US") { $issues += "system locale = $systemLocale" }
            if ($uiOverride -and $uiOverride.Name -ne "en-US") { $issues += "UI override = $($uiOverride.Name)" }
            if ($userList -notcontains "en-US") { $issues += "en-US not in user list" }
            if ($userList -contains "fr-FR") { $issues += "fr-FR still present in user list" }
            if ($inputTips -notcontains "0409:00000409") { $issues += "US keyboard (0409:00000409) not active" }
            if ($inputTips | Where-Object { $_ -like "*040c*" }) { $issues += "French keyboard layout still present" }

            if ($issues.Count -eq 0) {
                return @{ Status = "ok"; Detail = "System locale en-US, US keyboard, French removed" }
            }

            return @{ Status = "partial"; Detail = ($issues -join " ; ") }
        }
        catch {
            return @{ Status = "unknown"; Detail = "Unable to read locale: $($_.Exception.Message)" }
        }
    }
}
