# =============================================================================
# Tool_Search.ps1 - エクスポート済みデータ検索 (search)
# Data/ フォルダ内のCSV/JSONファイルからキーワード検索する
# =============================================================================

function Invoke-DataSearch {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Keyword,
        [string]$DataPath
    )

    if (-not (Test-Path $DataPath)) {
        Write-Host "データフォルダが見つかりません: $DataPath" -ForegroundColor Red
        return
    }

    $files = Get-ChildItem -Path $DataPath -Include "*.json", "*.csv" -Recurse
    if ($files.Count -eq 0) {
        Write-Host "データファイルがありません" -ForegroundColor Yellow
        return
    }

    $totalHits = 0

    foreach ($file in $files) {
        $hits = @()

        switch ($file.Extension.ToLower()) {
            ".json" {
                $data = Get-Content $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
                if ($data -is [System.Array]) {
                    foreach ($record in $data) {
                        $text = ($record.PSObject.Properties | ForEach-Object { $_.Value }) -join " "
                        if ($text -match [regex]::Escape($Keyword)) {
                            $hits += $record
                        }
                    }
                }
            }
            ".csv" {
                $data = Import-Csv $file.FullName -Encoding UTF8
                foreach ($record in $data) {
                    $text = ($record.PSObject.Properties | ForEach-Object { $_.Value }) -join " "
                    if ($text -match [regex]::Escape($Keyword)) {
                        $hits += $record
                    }
                }
            }
        }

        if ($hits.Count -gt 0) {
            Write-Host ""
            Write-Host "--- $($file.Name) ($($hits.Count) 件ヒット) ---" -ForegroundColor Cyan
            $hits | Format-Table -AutoSize | Out-String | Write-Host
            $totalHits += $hits.Count
        }
    }

    if ($totalHits -eq 0) {
        Write-Host "「$Keyword」に一致するデータはありませんでした" -ForegroundColor Yellow
    } else {
        Write-Host "合計 $totalHits 件ヒット" -ForegroundColor Green
    }
}
