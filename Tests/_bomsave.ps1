param([string]$Path)
$c = [System.IO.File]::ReadAllText($Path)
[System.IO.File]::WriteAllText($Path, $c, [System.Text.UTF8Encoding]::new($true))
Write-Host "BOM saved: $Path"
