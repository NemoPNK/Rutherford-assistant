@{
    Id       = "start-policy"
    Label    = "Start menu Recommended hidden + pins emptied"
    Category = "Policy"
    Order    = 50
    Test     = {
        try {
            $build = [int](Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name CurrentBuildNumber -ErrorAction Stop).CurrentBuildNumber
        }
        catch {
            $build = 0
        }

        $path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer"
        $issues = @()

        # HideRecommendedSection: only honored on build 22621+ (Win11 22H2), per Microsoft Policy CSP doc.
        $hide = $null
        try { $hide = (Get-ItemProperty -Path $path -Name "HideRecommendedSection" -ErrorAction Stop).HideRecommendedSection } catch { }

        if ($build -ge 22621) {
            if ($hide -ne 1) { $issues += "HideRecommendedSection=$hide (expected 1)" }
        }
        else {
            $issues += "build $build < 22621: HideRecommendedSection not supported by this Windows"
        }

        # ConfigureStartPins: supported from build 22000 (Win11 21H2); value must be the JSON itself.
        $pins = $null
        try { $pins = (Get-ItemProperty -Path $path -Name "ConfigureStartPins" -ErrorAction Stop).ConfigureStartPins } catch { }

        if ($build -ge 22000) {
            if ($pins -ne '{"pinnedList":[]}') {
                $detail = if ($null -eq $pins) { "not set" } else { "unexpected value" }
                $issues += "ConfigureStartPins $detail (expected empty pinnedList JSON)"
            }
        }

        if ($issues.Count -eq 0) {
            return @{ Status = "ok"; Detail = "Recommended hidden and pinned section emptied (build $build)" }
        }

        if ($build -lt 22000 -and $build -gt 0) {
            return @{ Status = "unknown"; Detail = "Start policies unsupported on build $build : $($issues -join ' ; ')" }
        }

        return @{ Status = "partial"; Detail = ($issues -join " ; ") }
    }
}
