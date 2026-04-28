@{
    Id       = "start-policy"
    Label    = "Start menu Recommended hidden"
    Category = "Policy"
    Order    = 50
    Test     = {
        $path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer"
        if (-not (Test-Path $path)) {
            return @{ Status = "missing"; Detail = "Explorer policy key not present" }
        }

        try {
            $value = (Get-ItemProperty -Path $path -Name "HideRecommendedSection" -ErrorAction Stop).HideRecommendedSection
        }
        catch {
            return @{ Status = "missing"; Detail = "HideRecommendedSection not set" }
        }

        if ($value -eq 1) {
            return @{ Status = "ok"; Detail = "HideRecommendedSection = 1" }
        }

        return @{ Status = "missing"; Detail = "HideRecommendedSection = $value (expected 1)" }
    }
}
