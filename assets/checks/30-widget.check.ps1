@{
    Id       = "widget"
    Label    = "Widget / News policy disabled"
    Category = "Policy"
    Order    = 30
    Test     = {
        $path = "HKLM:\Software\Policies\Microsoft\Dsh"
        if (-not (Test-Path $path)) {
            return @{ Status = "missing"; Detail = "Policy key Dsh not present" }
        }

        try {
            $value = (Get-ItemProperty -Path $path -Name "AllowNewsAndInterests" -ErrorAction Stop).AllowNewsAndInterests
        }
        catch {
            return @{ Status = "missing"; Detail = "AllowNewsAndInterests value not set" }
        }

        if ($value -eq 0) {
            return @{ Status = "ok"; Detail = "AllowNewsAndInterests = 0" }
        }

        return @{ Status = "missing"; Detail = "AllowNewsAndInterests = $value (expected 0)" }
    }
}
