# =============================================================================
# Tool_FolderTree.ps1 - フォルダツリー表示 (ftree)
# 指定フォルダの階層構造を見やすく表示する
# =============================================================================

function Invoke-FolderTree {
    param(
        [string]$Path = ".",
        [string[]]$ExcludeDirs = @(),
        [int]$MaxDepth = 5
    )

    # 相対パスの場合、スクリプトのベースパスからの相対として解決を試みる
    if (-not [System.IO.Path]::IsPathRooted($Path)) {
        $candidate = Join-Path $script:BasePath $Path
        if (Test-Path $candidate) {
            $Path = (Resolve-Path $candidate).Path
        } else {
            $Path = (Resolve-Path $Path -ErrorAction SilentlyContinue).Path
        }
    }
    if (-not $Path -or -not (Test-Path $Path)) {
        Write-Host "パスが見つかりません: $Path" -ForegroundColor Red
        return
    }

    Write-Host ""
    Write-Host $Path -ForegroundColor Cyan

    function Show-Tree {
        param(
            [string]$Dir,
            [string]$Prefix,
            [int]$Depth
        )

        if ($Depth -ge $MaxDepth) {
            Write-Host "$Prefix  ..." -ForegroundColor DarkGray
            return
        }

        $items = Get-ChildItem -Path $Dir -Force -ErrorAction SilentlyContinue |
            Where-Object {
                if ($_.PSIsContainer) {
                    $_.Name -notin $ExcludeDirs
                } else {
                    $true
                }
            } |
            Sort-Object { -not $_.PSIsContainer }, Name

        for ($i = 0; $i -lt $items.Count; $i++) {
            $item = $items[$i]
            $isLast = ($i -eq $items.Count - 1)
            $connector = if ($isLast) { "+-- " } else { "|-- " }
            $nextPrefix = if ($isLast) { "$Prefix    " } else { "$Prefix|   " }

            if ($item.PSIsContainer) {
                Write-Host "$Prefix$connector$($item.Name)\" -ForegroundColor Yellow
                Show-Tree -Dir $item.FullName -Prefix $nextPrefix -Depth ($Depth + 1)
            } else {
                $size = ""
                if ($item.Length -ge 1MB) {
                    $size = " ({0:N1} MB)" -f ($item.Length / 1MB)
                } elseif ($item.Length -ge 1KB) {
                    $size = " ({0:N1} KB)" -f ($item.Length / 1KB)
                }
                Write-Host "$Prefix$connector$($item.Name)$size"
            }
        }
    }

    Show-Tree -Dir $Path -Prefix "" -Depth 0
    Write-Host ""
}
