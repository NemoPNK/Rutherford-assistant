@{
    Id       = "ops-folder"
    Label    = "OPS folder copied to C:\"
    Category = "Filesystem"
    Order    = 60
    Test     = {
        $target = "C:\OPS"
        if (-not (Test-Path $target)) {
            return @{ Status = "missing"; Detail = "C:\OPS does not exist" }
        }

        $items = Get-ChildItem -Path $target -ErrorAction SilentlyContinue
        if (-not $items -or $items.Count -eq 0) {
            return @{ Status = "partial"; Detail = "C:\OPS exists but is empty" }
        }

        return @{ Status = "ok"; Detail = "C:\OPS contains $($items.Count) item(s)" }
    }
}
