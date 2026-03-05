# =============================================================================
# ExcelSync.ps1 - Excel読み書きモジュール
# シート＝テーブル、1行目＝ヘッダー
# =============================================================================

$script:excelApp = $null

function Open-ExcelApp {
    if (-not $script:excelApp) {
        $script:excelApp = New-Object -ComObject Excel.Application
        $script:excelApp.Visible = $false
        $script:excelApp.DisplayAlerts = $false
    }
    return $script:excelApp
}

function Close-ExcelApp {
    if ($script:excelApp) {
        try {
            $script:excelApp.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:excelApp) | Out-Null
        } catch {}
        $script:excelApp = $null
    }
}

function Read-ExcelTables {
    param([string]$ExcelPath)
    if (-not (Test-Path $ExcelPath)) { return $null }

    $app = Open-ExcelApp
    $wb = $app.Workbooks.Open($ExcelPath, 0, $true)  # ReadOnly
    $tables = [ordered]@{}

    for ($s = 1; $s -le $wb.Sheets.Count; $s++) {
        $ws = $wb.Sheets.Item($s)
        $sheetName = $ws.Name
        $usedRange = $ws.UsedRange
        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count

        if ($rowCount -lt 1 -or $colCount -lt 1) { continue }

        # ヘッダー読み取り
        $headers = @()
        for ($c = 1; $c -le $colCount; $c++) {
            $val = $ws.Cells.Item(1, $c).Text
            if (-not $val) { break }
            $headers += $val
        }
        $colCount = $headers.Count
        if ($colCount -eq 0) { continue }

        # データ読み取り
        $records = @()
        for ($r = 2; $r -le $rowCount; $r++) {
            # 空行スキップ
            $firstCell = $ws.Cells.Item($r, 1).Text
            if (-not $firstCell) { continue }

            $obj = New-Object PSObject
            for ($c = 1; $c -le $colCount; $c++) {
                $obj | Add-Member -NotePropertyName $headers[$c-1] -NotePropertyValue $ws.Cells.Item($r, $c).Text
            }
            $records += $obj
        }
        $tables[$sheetName] = $records
    }

    $wb.Close($false)
    return $tables
}

function Save-ExcelRecord {
    param(
        [string]$ExcelPath,
        [string]$SheetName,
        [int]$RecordIndex,
        [PSObject]$Record
    )

    $app = Open-ExcelApp
    $wb = $app.Workbooks.Open($ExcelPath)
    $ws = $wb.Sheets.Item($SheetName)

    $row = $RecordIndex + 2  # 1行目ヘッダー、0-indexed → 2-indexed
    $col = 1
    foreach ($prop in $Record.PSObject.Properties) {
        $ws.Cells.Item($row, $col) = $prop.Value
        $col++
    }

    $wb.Save()
    $wb.Close()
}

function Add-ExcelRecord {
    param(
        [string]$ExcelPath,
        [string]$SheetName,
        [PSObject]$Record
    )

    $app = Open-ExcelApp
    $wb = $app.Workbooks.Open($ExcelPath)
    $ws = $wb.Sheets.Item($SheetName)

    $lastRow = $ws.UsedRange.Rows.Count + 1
    $col = 1
    foreach ($prop in $Record.PSObject.Properties) {
        $ws.Cells.Item($lastRow, $col) = $prop.Value
        $col++
    }

    $wb.Save()
    $wb.Close()
}

function Remove-ExcelRecord {
    param(
        [string]$ExcelPath,
        [string]$SheetName,
        [int]$RecordIndex
    )

    $app = Open-ExcelApp
    $wb = $app.Workbooks.Open($ExcelPath)
    $ws = $wb.Sheets.Item($SheetName)

    $row = $RecordIndex + 2
    [void]$ws.Rows.Item($row).Delete()

    $wb.Save()
    $wb.Close()
}
