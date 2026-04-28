@{
    Id       = "consumer-features"
    Label    = "Windows consumer features blocked"
    Category = "Policy"
    Order    = 80
    Test     = {
        $expected = @(
            @{ Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"; Name = "DisableConsumerFeatures";          Want = 1 }
            @{ Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"; Name = "DisableWindowsConsumerFeatures";   Want = 1 }
            @{ Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"; Name = "DisableSoftLanding";               Want = 1 }
            @{ Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"; Name = "DisableWindowsSpotlightFeatures";  Want = 1 }
            @{ Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo"; Name = "DisabledByGroupPolicy";          Want = 1 }
            @{ Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot"; Name = "TurnOffWindowsCopilot";          Want = 1 }
        )

        $missing = @()
        $okCount = 0

        foreach ($entry in $expected) {
            try {
                $value = (Get-ItemProperty -Path $entry.Path -Name $entry.Name -ErrorAction Stop).$($entry.Name)
                if ($value -eq $entry.Want) {
                    $okCount++
                }
                else {
                    $missing += "$($entry.Name)=$value"
                }
            }
            catch {
                $missing += $entry.Name
            }
        }

        if ($missing.Count -eq 0) {
            return @{ Status = "ok"; Detail = "$okCount/$($expected.Count) policies set" }
        }

        if ($okCount -eq 0) {
            return @{ Status = "missing"; Detail = "No consumer-feature policy set" }
        }

        return @{ Status = "partial"; Detail = "$okCount/$($expected.Count) set ; missing: $($missing -join ', ')" }
    }
}
