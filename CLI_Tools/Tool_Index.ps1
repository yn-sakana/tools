# =============================================================================
# Tool_Index.ps1 - 対話型テーブル検索 (index)
# 矢印キーで選択、Enter で決定、文字入力で絞り込み、Esc で戻る
# Excel 読み書き対応
# =============================================================================

# --- ExcelSync 読み込み ---
$_excelSyncPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "ExcelSync.ps1"
if (Test-Path $_excelSyncPath) { . $_excelSyncPath }

# --- 表示ユーティリティ ---

function Get-DisplayWidth {
    param([string]$Text)
    $w = 0
    foreach ($c in $Text.ToCharArray()) {
        if ([int]$c -gt 0xFF) { $w += 2 } else { $w += 1 }
    }
    return $w
}

function Fit-ToWidth {
    param([string]$Text, [int]$Width)
    $w = 0
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $Text.ToCharArray()) {
        $cw = if ([int]$c -gt 0xFF) { 2 } else { 1 }
        if ($w + $cw -gt $Width) { break }
        [void]$sb.Append($c)
        $w += $cw
    }
    if ($w -lt $Width) {
        [void]$sb.Append(' ', ($Width - $w))
    }
    return $sb.ToString()
}

# --- セレクターUI ---

function Show-Selector {
    param(
        [string]$Title,
        [string[]]$Items,
        [string[]]$DisplayItems
    )

    if (-not $DisplayItems) { $DisplayItems = $Items }
    if ($Items.Count -eq 0) { return $null }

    $script:_sel_selected = 0
    $script:_sel_filter = ""
    $script:_sel_filtered = @(0..($Items.Count - 1))
    $script:_sel_maxDisplay = 15
    $script:_sel_startTop = [Console]::CursorTop
    $script:_sel_lastLineCount = 0
    $colWidth = [Console]::BufferWidth - 1

    function Render {
        [Console]::CursorVisible = $false
        [Console]::SetCursorPosition(0, $script:_sel_startTop)

        $lines = @()

        if ($script:_sel_filter) {
            $lines += "$Title (filter: $($script:_sel_filter))"
        } else {
            $lines += $Title
        }

        $maxShow = [Math]::Min($script:_sel_filtered.Count, $script:_sel_maxDisplay)
        for ($i = 0; $i -lt $maxShow; $i++) {
            $idx = $script:_sel_filtered[$i]
            if ($i -eq $script:_sel_selected) {
                $lines += "  > $($DisplayItems[$idx])"
            } else {
                $lines += "    $($DisplayItems[$idx])"
            }
        }
        if ($script:_sel_filtered.Count -gt $script:_sel_maxDisplay) {
            $lines += "    ... (他 $($script:_sel_filtered.Count - $script:_sel_maxDisplay) 件)"
        }
        if ($script:_sel_filtered.Count -eq 0) {
            $lines += "    (該当なし)"
        }

        $lines += ""
        $lines += "  [↑↓] 選択  [Enter] 決定  [文字] 絞り込み  [Esc] 戻る"

        while ($lines.Count -lt $script:_sel_lastLineCount) {
            $lines += ""
        }

        foreach ($line in $lines) {
            Write-Host (Fit-ToWidth $line $colWidth)
        }

        $script:_sel_startTop = [Console]::CursorTop - $lines.Count
        $script:_sel_lastLineCount = $lines.Count
        [Console]::CursorVisible = $true
    }

    function ClearSelector {
        [Console]::SetCursorPosition(0, $script:_sel_startTop)
        $blank = " " * $colWidth
        for ($i = 0; $i -lt $script:_sel_lastLineCount; $i++) {
            Write-Host $blank
        }
        [Console]::SetCursorPosition(0, $script:_sel_startTop)
    }

    function ApplyFilter {
        if ($script:_sel_filter) {
            $escaped = [regex]::Escape($script:_sel_filter)
            $script:_sel_filtered = @(
                for ($i = 0; $i -lt $Items.Count; $i++) {
                    if ($DisplayItems[$i] -match $escaped) { $i }
                }
            )
        } else {
            $script:_sel_filtered = @(0..($Items.Count - 1))
        }
        $script:_sel_selected = 0
    }

    Render

    while ($true) {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)

            switch ($key.Key) {
                "UpArrow" {
                    if ($script:_sel_selected -gt 0) { $script:_sel_selected-- }
                    Render
                }
                "DownArrow" {
                    $maxIdx = [Math]::Min($script:_sel_filtered.Count, $script:_sel_maxDisplay) - 1
                    if ($script:_sel_selected -lt $maxIdx) { $script:_sel_selected++ }
                    Render
                }
                "Enter" {
                    if ($script:_sel_filtered.Count -gt 0) {
                        ClearSelector
                        return $Items[$script:_sel_filtered[$script:_sel_selected]]
                    }
                }
                "Escape" {
                    ClearSelector
                    return $null
                }
                "Backspace" {
                    if ($script:_sel_filter.Length -gt 0) {
                        $script:_sel_filter = $script:_sel_filter.Substring(0, $script:_sel_filter.Length - 1)
                        ApplyFilter
                        Render
                    }
                }
                default {
                    if ($key.KeyChar -and $key.KeyChar -ne "`0") {
                        $script:_sel_filter += $key.KeyChar
                        ApplyFilter
                        Render
                    }
                }
            }
        } else {
            Start-Sleep -Milliseconds 30
        }
    }
}

# --- レコード表示・操作 ---

function Get-OpenableFields {
    param($Record)
    $pattern = '^(path|file|folder|dir|directory|url|link|href|uri)$'
    $fields = @()
    foreach ($prop in $Record.PSObject.Properties) {
        if ($prop.Name -match $pattern -and $prop.Value) {
            $fields += $prop
        }
    }
    return $fields
}

function Show-Record {
    param($Record)
    Write-Host ""
    Write-Host "  ===========================" -ForegroundColor DarkGray
    foreach ($prop in $Record.PSObject.Properties) {
        Write-Host ("  {0,-15} : {1}" -f $prop.Name, $prop.Value)
    }
    Write-Host "  ===========================" -ForegroundColor DarkGray
    Write-Host ""
}

function Edit-Record {
    param($Record)

    # 編集用のコピーを作成
    $values = [ordered]@{}
    foreach ($prop in $Record.PSObject.Properties) {
        $values[$prop.Name] = [string]$prop.Value
    }
    $fields = @($values.Keys)

    $selected = 0
    $startTop = [Console]::CursorTop
    $lastLineCount = 0
    $colWidth = [Console]::BufferWidth - 1

    function RenderEditor {
        [Console]::CursorVisible = $false
        [Console]::SetCursorPosition(0, $startTop)

        $lines = @()
        $lines += "  --- 編集 [↑↓] 選択  [Enter] 値を変更  [Esc] 終了 ---"
        for ($i = 0; $i -lt $fields.Count; $i++) {
            $name = $fields[$i]
            $val = $values[$name]
            if ($i -eq $selected) {
                $lines += "  > {0,-15} : {1}" -f $name, $val
            } else {
                $lines += "    {0,-15} : {1}" -f $name, $val
            }
        }
        $lines += ""

        while ($lines.Count -lt $lastLineCount) { $lines += "" }

        foreach ($line in $lines) {
            Write-Host (Fit-ToWidth $line $colWidth)
        }

        $script:_edit_startTop = [Console]::CursorTop - $lines.Count
        $script:_edit_lastLineCount = $lines.Count
        [Console]::CursorVisible = $true
    }

    # 初回クロージャ用に変数をscriptスコープに
    $script:_edit_startTop = $startTop
    $script:_edit_lastLineCount = $lastLineCount

    # 描画ループ内で参照更新
    RenderEditor

    while ($true) {
        $startTop = $script:_edit_startTop
        $lastLineCount = $script:_edit_lastLineCount

        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)

            switch ($key.Key) {
                "UpArrow" {
                    if ($selected -gt 0) { $selected-- }
                    RenderEditor
                }
                "DownArrow" {
                    if ($selected -lt $fields.Count - 1) { $selected++ }
                    RenderEditor
                }
                "Enter" {
                    $name = $fields[$selected]
                    # 入力行を表示してRead-Host
                    [Console]::SetCursorPosition(0, $script:_edit_startTop + $script:_edit_lastLineCount)
                    Write-Host "  $name = " -NoNewline -ForegroundColor Cyan
                    $newVal = Read-Host
                    if ($newVal -ne "") {
                        $values[$name] = $newVal
                    }
                    # 入力行をクリアして再描画
                    [Console]::SetCursorPosition(0, $script:_edit_startTop + $script:_edit_lastLineCount)
                    Write-Host (Fit-ToWidth "" $colWidth)
                    RenderEditor
                }
                "Escape" {
                    # エディタ表示をクリア
                    [Console]::SetCursorPosition(0, $script:_edit_startTop)
                    $blank = " " * $colWidth
                    for ($i = 0; $i -lt $script:_edit_lastLineCount; $i++) {
                        Write-Host $blank
                    }
                    [Console]::SetCursorPosition(0, $script:_edit_startTop)

                    # 変更があるか確認
                    $hasChanges = $false
                    foreach ($name in $fields) {
                        $orig = [string]$Record.$name
                        if ($values[$name] -ne $orig) { $hasChanges = $true; break }
                    }

                    if (-not $hasChanges) {
                        Write-Host "  変更なし" -ForegroundColor DarkGray
                        return $null
                    }

                    # 差分表示
                    Write-Host "  --- 変更内容 ---" -ForegroundColor Cyan
                    foreach ($name in $fields) {
                        $orig = [string]$Record.$name
                        $newV = $values[$name]
                        if ($newV -ne $orig) {
                            Write-Host ("  {0,-15} : {1}" -f $name, $orig) -ForegroundColor DarkGray
                            Write-Host ("  {0,-15} → {1}" -f "", $newV) -ForegroundColor Green
                        } else {
                            Write-Host ("  {0,-15} : {1}" -f $name, $orig)
                        }
                    }

                    Write-Host ""
                    Write-Host "  保存しますか? [Y/N]: " -NoNewline -ForegroundColor Yellow
                    $confirm = Read-Host
                    if ($confirm -ne "Y" -and $confirm -ne "y") {
                        Write-Host "  キャンセル" -ForegroundColor DarkGray
                        return $null
                    }

                    $edited = New-Object PSObject
                    foreach ($name in $fields) {
                        $edited | Add-Member -NotePropertyName $name -NotePropertyValue $values[$name]
                    }
                    return $edited
                }
            }
        } else {
            Start-Sleep -Milliseconds 30
        }
    }
}

function New-EmptyRecord {
    param([string[]]$Fields)
    Write-Host "  --- 新規追加 ---" -ForegroundColor Cyan
    $rec = New-Object PSObject
    foreach ($f in $Fields) {
        Write-Host "  ${f}: " -NoNewline -ForegroundColor DarkGray
        $val = Read-Host
        $rec | Add-Member -NotePropertyName $f -NotePropertyValue $val
    }
    return $rec
}

function Show-RecordActions {
    param($Record, $Source, [int]$RecordIndex, [ref]$Records)

    $openable = @(Get-OpenableFields -Record $Record)
    $isExcel = ($Source -and $Source.Type -eq "excel")

    # アクション一覧を組み立て
    $hints = @()
    if ($openable.Count -eq 1) { $hints += "[O] $($openable[0].Name) を開く" }
    elseif ($openable.Count -gt 1) { $hints += "[O] 開く" }
    if ($isExcel) { $hints += "[E] 編集  [A] 追加  [D] 削除" }
    $hints += "[Enter] 続ける  [Esc] 戻る"

    Write-Host "  $($hints -join '  ')" -ForegroundColor DarkGray

    while ($true) {
        if ([Console]::KeyAvailable) {
            $k = [Console]::ReadKey($true)
            if ($k.Key -eq "Escape") { return "escape" }
            if ($k.Key -eq "Enter") { return "continue" }

            # 開く
            if ($k.Key -eq "O" -and $openable.Count -gt 0) {
                if ($openable.Count -eq 1) {
                    Write-Host "  => $($openable[0].Value)" -ForegroundColor Green
                    Start-Process $openable[0].Value
                } else {
                    $actionItems = @("--- 戻る ---")
                    $actionDisplay = @("(戻る)")
                    foreach ($f in $openable) {
                        $actionItems += $f.Name
                        $actionDisplay += "$($f.Name) : $($f.Value)"
                    }
                    $choice = Show-Selector -Title "=== 開く ===" -Items $actionItems -DisplayItems $actionDisplay
                    if ($choice -and $choice -ne "--- 戻る ---") {
                        $target = ($openable | Where-Object { $_.Name -eq $choice }).Value
                        Write-Host "  => $target" -ForegroundColor Green
                        Start-Process $target
                    }
                }
                return "continue"
            }

            # 編集
            if ($k.Key -eq "E" -and $isExcel) {
                $edited = Edit-Record $Record
                if ($edited) {
                    try {
                        Save-ExcelRecord -ExcelPath $Source.File -SheetName $Source.Sheet -RecordIndex $RecordIndex -Record $edited
                        $Records.Value[$RecordIndex] = $edited
                        Write-Host "  保存しました" -ForegroundColor Green
                    } catch {
                        Write-Host "  保存エラー: $_" -ForegroundColor Red
                    }
                    return "reload"
                }
                return "continue"
            }

            # 追加
            if ($k.Key -eq "A" -and $isExcel) {
                $fields = @($Record.PSObject.Properties | ForEach-Object { $_.Name })
                $newRec = New-EmptyRecord $fields
                try {
                    Add-ExcelRecord -ExcelPath $Source.File -SheetName $Source.Sheet -Record $newRec
                    $Records.Value += $newRec
                    Write-Host "  追加しました" -ForegroundColor Green
                } catch {
                    Write-Host "  追加エラー: $_" -ForegroundColor Red
                }
                return "reload"
            }

            # 削除
            if ($k.Key -eq "D" -and $isExcel) {
                Write-Host "  削除しますか? [Y/N]: " -NoNewline -ForegroundColor Yellow
                $confirm = Read-Host
                if ($confirm -eq "Y" -or $confirm -eq "y") {
                    try {
                        Remove-ExcelRecord -ExcelPath $Source.File -SheetName $Source.Sheet -RecordIndex $RecordIndex
                        $list = [System.Collections.ArrayList]@($Records.Value)
                        $list.RemoveAt($RecordIndex)
                        $Records.Value = @($list)
                        Write-Host "  削除しました" -ForegroundColor Green
                    } catch {
                        Write-Host "  削除エラー: $_" -ForegroundColor Red
                    }
                }
                return "reload"
            }
        } else {
            Start-Sleep -Milliseconds 30
        }
    }
}

# --- メイン ---

function Invoke-IndexSearch {
    param(
        [string]$TableName,
        [string]$Key,
        [string]$DataPath
    )

    if (-not (Test-Path $DataPath)) {
        Write-Host "データフォルダが見つかりません: $DataPath" -ForegroundColor Red
        return
    }

    # --- データ読み込み（Excel優先、JSON/CSVフォールバック）---
    $allTables = [ordered]@{}
    $tableSources = @{}

    # Excel
    $xlsFiles = @(Get-ChildItem -Path $DataPath -Filter "*.xlsx" -ErrorAction SilentlyContinue)
    foreach ($xf in $xlsFiles) {
        try {
            $tables = Read-ExcelTables -ExcelPath $xf.FullName
            if ($tables) {
                foreach ($k in $tables.Keys) {
                    $allTables[$k] = $tables[$k]
                    $tableSources[$k] = @{ Type = "excel"; File = $xf.FullName; Sheet = $k }
                }
            }
        } catch {
            Write-Host "Excel読み込みエラー ($($xf.Name)): $_" -ForegroundColor Yellow
        }
    }

    # JSON/CSV（Excel と同名はスキップ）
    $files = @(Get-ChildItem -Path $DataPath -Recurse | Where-Object { $_.Extension -match '^\.(json|csv)$' })
    foreach ($f in $files) {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
        if ($allTables.Contains($name)) { continue }
        $records = @()
        switch ($f.Extension.ToLower()) {
            ".json" {
                $d = Get-Content $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
                if ($d -is [System.Array]) { $records = $d } else { $records = @($d) }
            }
            ".csv" { $records = @(Import-Csv $f.FullName -Encoding UTF8) }
        }
        $allTables[$name] = $records
        $tableSources[$name] = @{ Type = $f.Extension.TrimStart(".").ToLower(); File = $f.FullName }
    }

    if ($allTables.Count -eq 0) {
        Write-Host "データがありません" -ForegroundColor Yellow
        Close-ExcelApp
        return
    }

    # --- テーブル選択 ---
    if (-not $TableName) {
        $tNames = @($allTables.Keys)
        $tDisplay = @($tNames | ForEach-Object {
            $src = if ($tableSources[$_].Type -eq "excel") { "[xlsx]" } else { "[" + $tableSources[$_].Type + "]" }
            "$_ ($($allTables[$_].Count) 件) $src"
        })

        Write-Host ""
        $TableName = Show-Selector -Title "=== テーブルを選択 ===" -Items $tNames -DisplayItems $tDisplay
        if (-not $TableName) {
            Close-ExcelApp
            return
        }
    }

    if (-not $allTables.Contains($TableName)) {
        Write-Host "テーブルが見つかりません: $TableName" -ForegroundColor Red
        Close-ExcelApp
        return
    }

    # --- レコード選択ループ ---
    $source = $tableSources[$TableName]

    while ($true) {
        $records = $allTables[$TableName]
        if ($records.Count -eq 0) {
            Write-Host "データがありません" -ForegroundColor Yellow
            break
        }

        $props = @($records[0].PSObject.Properties | Select-Object -First 3 | ForEach-Object { $_.Name })
        $recordDisplay = @($records | ForEach-Object {
            $r = $_
            ($props | ForEach-Object { $r.$_ }) -join " | "
        })
        $recordIndices = @(0..($records.Count - 1) | ForEach-Object { [string]$_ })

        Write-Host ""
        $srcLabel = if ($source.Type -eq "excel") { " [xlsx]" } else { "" }
        $selectedIdx = Show-Selector -Title "=== $TableName ($($records.Count) 件)$srcLabel ===" -Items $recordIndices -DisplayItems $recordDisplay
        if ($selectedIdx -eq $null) { break }

        $recIdx = [int]$selectedIdx
        $record = $records[$recIdx]
        Show-Record -Record $record

        $recordsRef = [ref]$allTables[$TableName]
        $action = Show-RecordActions -Record $record -Source $source -RecordIndex $recIdx -Records $recordsRef
        $allTables[$TableName] = $recordsRef.Value

        if ($action -eq "escape") { break }
        # "reload" と "continue" はループ継続（reload は表示を再構築）
    }

    Close-ExcelApp
}
