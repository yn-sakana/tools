# =============================================================================
# Tool_Index.ps1 - 対話型テーブル検索 (index)
# 矢印キーで選択、Enter で決定、文字入力で絞り込み、Esc で戻る
# =============================================================================

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

function Show-RecordActions {
    param($Record)
    $openable = @(Get-OpenableFields -Record $Record)
    if ($openable.Count -eq 0) {
        Write-Host "  [Enter] 続けて検索  [Esc] 戻る" -ForegroundColor DarkGray
        while ($true) {
            if ([Console]::KeyAvailable) {
                $k = [Console]::ReadKey($true)
                if ($k.Key -eq "Escape") { return "escape" }
                if ($k.Key -eq "Enter") { return "continue" }
            } else {
                Start-Sleep -Milliseconds 30
            }
        }
    }

    if ($openable.Count -eq 1) {
        Write-Host "  [O] $($openable[0].Name) を開く  [Enter] 続けて検索  [Esc] 戻る" -ForegroundColor DarkGray
        while ($true) {
            if ([Console]::KeyAvailable) {
                $k = [Console]::ReadKey($true)
                if ($k.Key -eq "Escape") { return "escape" }
                if ($k.Key -eq "Enter") { return "continue" }
                if ($k.Key -eq "O") {
                    Write-Host "  => $($openable[0].Value)" -ForegroundColor Green
                    Start-Process $openable[0].Value
                    return "continue"
                }
            } else {
                Start-Sleep -Milliseconds 30
            }
        }
    }

    $actionItems = @("--- 戻る ---")
    $actionDisplay = @("(検索に戻る)")
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
    return "continue"
}

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

    $files = @(Get-ChildItem -Path $DataPath -Recurse | Where-Object { $_.Extension -match '^\.(json|csv)$' })
    if ($files.Count -eq 0) {
        Write-Host "データファイルがありません" -ForegroundColor Yellow
        return
    }

    # --- テーブル選択 ---
    if (-not $TableName) {
        $tableNames = @()
        $tableDisplay = @()
        foreach ($f in $files) {
            $name = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
            $tableNames += $name
            $count = 0
            switch ($f.Extension.ToLower()) {
                ".json" {
                    $d = Get-Content $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
                    if ($d -is [System.Array]) { $count = $d.Count }
                }
                ".csv" { $count = @(Import-Csv $f.FullName -Encoding UTF8).Count }
            }
            $tableDisplay += "$name ($count 件)"
        }

        Write-Host ""
        $TableName = Show-Selector -Title "=== テーブルを選択 ===" -Items $tableNames -DisplayItems $tableDisplay
        if (-not $TableName) { return }
    }

    # --- データ読み込み ---
    $file = Get-ChildItem -Path $DataPath -Recurse | Where-Object { $_.BaseName -eq $TableName -and $_.Extension -match '^\.(json|csv)$' } | Select-Object -First 1
    if (-not $file) {
        Write-Host "テーブルが見つかりません: $TableName" -ForegroundColor Red
        return
    }

    $records = @()
    switch ($file.Extension.ToLower()) {
        ".json" {
            $d = Get-Content $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($d -is [System.Array]) { $records = $d } else { $records = @($d) }
        }
        ".csv" { $records = @(Import-Csv $file.FullName -Encoding UTF8) }
    }

    if ($records.Count -eq 0) {
        Write-Host "データがありません" -ForegroundColor Yellow
        return
    }

    # --- レコード選択ループ ---
    $props = @($records[0].PSObject.Properties | Select-Object -First 3 | ForEach-Object { $_.Name })
    $recordDisplay = @($records | ForEach-Object {
        $r = $_
        ($props | ForEach-Object { $r.$_ }) -join " | "
    })
    $recordIndices = @(0..($records.Count - 1) | ForEach-Object { [string]$_ })

    while ($true) {
        Write-Host ""
        $selectedIdx = Show-Selector -Title "=== $TableName ($($records.Count) 件) ===" -Items $recordIndices -DisplayItems $recordDisplay
        if ($selectedIdx -eq $null) { break }
        $record = $records[[int]$selectedIdx]
        Show-Record -Record $record
        $action = Show-RecordActions -Record $record
        if ($action -eq "escape") { return }
    }
}
