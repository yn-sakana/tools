# =============================================================================
# IndexGUI.ps1 - index GUI版 (Windows Forms)
# Excel連携 + CRUD + テーブル検索・閲覧・開く
# =============================================================================

param([switch]$DryRun)

$ErrorActionPreference = "Stop"
$script:BasePath = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type @"
using System.Runtime.InteropServices;
public class DpiHelper {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
}
"@
[DpiHelper]::SetProcessDPIAware()

# --- ExcelSync 読み込み ---
. (Join-Path $script:BasePath "CLI_Tools\ExcelSync.ps1")

# --- config ---
$script:Config = Get-Content (Join-Path $script:BasePath "config.json") -Raw -Encoding UTF8 | ConvertFrom-Json
$script:openablePattern = '^(path|file|folder|dir|directory|url|link|href|uri)$'

# --- データ読み込み ---
$script:allTables = [ordered]@{}
$script:tableSource = @{}  # テーブル名 → @{ Type="excel"|"json"|"csv"; File=path; Sheet=name }
$script:tableFields = @{}  # テーブル名 → @("field1", "field2", ...)
$dataPath = $script:Config.dataPath

function Load-AllData {
    $script:allTables = [ordered]@{}
    $script:tableSource = @{}
    $script:tableFields = @{}

    if (-not (Test-Path $dataPath)) { return }

    # Excel ファイル読み込み（シート＝テーブル）
    $xlsFiles = @(Get-ChildItem -Path $dataPath -Filter "*.xlsx" -ErrorAction SilentlyContinue)
    foreach ($xf in $xlsFiles) {
        try {
            $tables = Read-ExcelTables -ExcelPath $xf.FullName
            if ($tables) {
                foreach ($key in $tables.Keys) {
                    $script:allTables[$key] = $tables[$key]
                    $script:tableSource[$key] = @{ Type = "excel"; File = $xf.FullName; Sheet = $key }
                    if ($tables[$key].Count -gt 0) {
                        $script:tableFields[$key] = @($tables[$key][0].PSObject.Properties | ForEach-Object { $_.Name })
                    }
                }
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Excel読み込みエラー: $($xf.Name)`n$_", "エラー", "OK", "Warning")
        }
    }

    # JSON/CSV 読み込み（Excelに同名テーブルがあればスキップ）
    $dataFiles = @(Get-ChildItem -Path $dataPath -Recurse | Where-Object { $_.Extension -match '^\.(json|csv)$' })
    foreach ($f in $dataFiles) {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
        if ($script:allTables.Contains($name)) { continue }
        $records = @()
        switch ($f.Extension.ToLower()) {
            ".json" {
                $d = Get-Content $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
                if ($d -is [System.Array]) { $records = $d } else { $records = @($d) }
            }
            ".csv" { $records = @(Import-Csv $f.FullName -Encoding UTF8) }
        }
        $script:allTables[$name] = $records
        $script:tableSource[$name] = @{ Type = $f.Extension.TrimStart(".").ToLower(); File = $f.FullName }
        if ($records.Count -gt 0) {
            $script:tableFields[$name] = @($records[0].PSObject.Properties | ForEach-Object { $_.Name })
        }
    }
}

Load-AllData
$script:tableNames = @($script:allTables.Keys)
$script:filteredIndices = @()

# =============================================================================
# フォーム
# =============================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "index"
$form.Size = New-Object System.Drawing.Size(750, 650)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Yu Gothic UI", 10)
$script:M = 8  # 基準マージン
$form.Padding = New-Object System.Windows.Forms.Padding($script:M)

# --- 上部: テーブル選択 + フィルタ ---
$pnlTop = New-Object System.Windows.Forms.TableLayoutPanel
$pnlTop.Dock = "Top"
$pnlTop.Height = 72
$pnlTop.ColumnCount = 2
$pnlTop.RowCount = 2
$pnlTop.Padding = New-Object System.Windows.Forms.Padding(0, 0, 0, $script:M)
$pnlTop.Margin = New-Object System.Windows.Forms.Padding(0)
[void]$pnlTop.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("AutoSize")))
[void]$pnlTop.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100)))
[void]$pnlTop.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 34)))
[void]$pnlTop.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 34)))

$lblTable = New-Object System.Windows.Forms.Label
$lblTable.Text = "テーブル:"
$lblTable.AutoSize = $true
$lblTable.Anchor = "Left"
$lblTable.Margin = New-Object System.Windows.Forms.Padding(0, 6, $script:M, 0)
$pnlTop.Controls.Add($lblTable, 0, 0)

$cmbTable = New-Object System.Windows.Forms.ComboBox
$cmbTable.DropDownStyle = "DropDownList"
$cmbTable.Dock = "Fill"
$cmbTable.Margin = New-Object System.Windows.Forms.Padding(0, 3, 0, 3)
$pnlTop.Controls.Add($cmbTable, 1, 0)

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = "絞り込み:"
$lblFilter.AutoSize = $true
$lblFilter.Anchor = "Left"
$lblFilter.Margin = New-Object System.Windows.Forms.Padding(0, 6, $script:M, 0)
$pnlTop.Controls.Add($lblFilter, 0, 1)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Dock = "Fill"
$txtFilter.Margin = New-Object System.Windows.Forms.Padding(0, 3, 0, 3)
$pnlTop.Controls.Add($txtFilter, 1, 1)

# --- 中央: レコードリスト ---
$listRecords = New-Object System.Windows.Forms.ListBox
$listRecords.Dock = "Fill"
$listRecords.Font = New-Object System.Drawing.Font("Consolas", 10)
$listRecords.IntegralHeight = $false

# --- 下部: 詳細エリア ---
$pnlDetail = New-Object System.Windows.Forms.Panel
$pnlDetail.Dock = "Bottom"
$pnlDetail.Height = 240
$pnlDetail.Padding = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)

$splitterDetail = New-Object System.Windows.Forms.Splitter
$splitterDetail.Dock = "Bottom"
$splitterDetail.Height = 4

# 操作ボタン行（開く + CRUD）
$pnlButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlButtons.Dock = "Top"
$pnlButtons.Height = 34
$pnlButtons.FlowDirection = "LeftToRight"
$pnlButtons.Padding = New-Object System.Windows.Forms.Padding(0, 1, 0, 1)

# フィールド編集エリア（スクロール可能）
$pnlFields = New-Object System.Windows.Forms.Panel
$pnlFields.Dock = "Fill"
$pnlFields.AutoScroll = $true
$pnlFields.Padding = New-Object System.Windows.Forms.Padding(0, $script:M, 0, 0)

$pnlDetail.Controls.Add($pnlFields)
$pnlDetail.Controls.Add($pnlButtons)

# 件数ラベル
$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Dock = "Bottom"
$lblCount.Height = 24
$lblCount.TextAlign = "BottomRight"
$lblCount.ForeColor = [System.Drawing.Color]::DimGray
$lblCount.Padding = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)

# フォーム組み立て
$form.Controls.Add($listRecords)
$form.Controls.Add($splitterDetail)
$form.Controls.Add($pnlDetail)
$form.Controls.Add($lblCount)
$form.Controls.Add($pnlTop)

# =============================================================================
# フィールドエディタ管理
# =============================================================================
$script:fieldEditors = @{}  # フィールド名 → TextBox

function Build-FieldEditors {
    param([string]$TableName)
    $pnlFields.SuspendLayout()
    $pnlFields.Controls.Clear()
    $script:fieldEditors = @{}

    $fields = $script:tableFields[$TableName]
    if (-not $fields) { $pnlFields.ResumeLayout(); return }

    $isEditable = ($script:tableSource[$TableName].Type -eq "excel")
    $m = $script:M
    # ラベル幅: 最長フィールド名に合わせ（最小80、最大160）
    $g = $pnlFields.CreateGraphics()
    $maxLblChars = 16  # この文字数を超えたら省略
    $lblW = 80
    foreach ($fn in $fields) {
        $w = [int][math]::Ceiling($g.MeasureString($fn, $pnlFields.Font).Width) + 8
        if ($w -gt $lblW) { $lblW = $w }
    }
    if ($lblW -gt 160) { $lblW = 160 }
    $g.Dispose()
    $txtX = $lblW + $m
    $y = $m
    $tip = New-Object System.Windows.Forms.ToolTip
    foreach ($fieldName in $fields) {
        # 長すぎるフィールド名は前...後で省略
        $displayName = $fieldName
        if ($fieldName.Length -gt $maxLblChars) {
            $head = $fieldName.Substring(0, [math]::Floor($maxLblChars / 2))
            $tail = $fieldName.Substring($fieldName.Length - [math]::Floor($maxLblChars / 2))
            $displayName = "$head...$tail"
        }
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $displayName
        $lbl.Location = New-Object System.Drawing.Point($m, ($y + 3))
        $lbl.Size = New-Object System.Drawing.Size($lblW, 22)
        $lbl.TextAlign = "MiddleRight"
        if ($displayName -ne $fieldName) { $tip.SetToolTip($lbl, $fieldName) }
        $pnlFields.Controls.Add($lbl)

        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Location = New-Object System.Drawing.Point($txtX, $y)
        $txt.Size = New-Object System.Drawing.Size(($pnlFields.ClientSize.Width - $txtX - $m), 22)
        $txt.Anchor = "Top,Left,Right"
        $txt.ReadOnly = (-not $isEditable)
        if (-not $isEditable) {
            $txt.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
        }
        $pnlFields.Controls.Add($txt)

        $script:fieldEditors[$fieldName] = $txt
        $y += 28
    }
    $pnlFields.ResumeLayout()
}

function Fill-FieldEditors {
    param([PSObject]$Record)
    if (-not $Record) {
        foreach ($txt in $script:fieldEditors.Values) { $txt.Text = "" }
        return
    }
    foreach ($prop in $Record.PSObject.Properties) {
        if ($script:fieldEditors.ContainsKey($prop.Name)) {
            $script:fieldEditors[$prop.Name].Text = [string]$prop.Value
        }
    }
}

function Get-EditedRecord {
    param([string]$TableName)
    $fields = $script:tableFields[$TableName]
    if (-not $fields) { return $null }
    $obj = New-Object PSObject
    foreach ($fieldName in $fields) {
        $val = ""
        if ($script:fieldEditors.ContainsKey($fieldName)) {
            $val = $script:fieldEditors[$fieldName].Text
        }
        $obj | Add-Member -NotePropertyName $fieldName -NotePropertyValue $val
    }
    return $obj
}

# =============================================================================
# ロジック
# =============================================================================
function Update-RecordList {
    $tableName = $cmbTable.SelectedItem
    if (-not $tableName) {
        $listRecords.Items.Clear()
        $pnlFields.Controls.Clear()
        $pnlButtons.Controls.Clear()
        $lblCount.Text = ""
        return
    }

    $records = $script:allTables[$tableName]
    $props = @()
    if ($records.Count -gt 0) {
        $props = @($records[0].PSObject.Properties | Select-Object -First 3 | ForEach-Object { $_.Name })
    }

    $listRecords.BeginUpdate()
    $listRecords.Items.Clear()
    $filter = $txtFilter.Text.Trim()
    $script:filteredIndices = @()

    for ($i = 0; $i -lt $records.Count; $i++) {
        $rec = $records[$i]
        $label = ($props | ForEach-Object { [string]$rec.$_ }) -join " | "
        if ($filter) {
            $allText = ($rec.PSObject.Properties | ForEach-Object { [string]$_.Value }) -join " "
            if ($allText -notmatch [regex]::Escape($filter)) { continue }
        }
        $script:filteredIndices += $i
        [void]$listRecords.Items.Add($label)
    }
    $listRecords.EndUpdate()

    $total = $records.Count
    $shown = $script:filteredIndices.Count
    if ($filter) {
        $lblCount.Text = "$shown / $total 件"
    } else {
        $lblCount.Text = "$total 件"
    }

    if ($listRecords.Items.Count -gt 0) { $listRecords.SelectedIndex = 0 }
}

function Update-Detail {
    $tableName = $cmbTable.SelectedItem
    $idx = $listRecords.SelectedIndex

    $pnlButtons.Controls.Clear()

    if (-not $tableName -or $idx -lt 0 -or $idx -ge $script:filteredIndices.Count) {
        Fill-FieldEditors $null
        return
    }

    $recIdx = $script:filteredIndices[$idx]
    $rec = $script:allTables[$tableName][$recIdx]
    Fill-FieldEditors $rec

    $src = $script:tableSource[$tableName]
    $isExcel = ($src.Type -eq "excel")

    # 開くボタン（openableフィールドがあれば表示、複数ならメニュー選択）
    $openable = @($rec.PSObject.Properties | Where-Object { $_.Name -match $script:openablePattern -and $_.Value })
    if ($openable.Count -eq 1) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = "$($openable[0].Name) を開く"
        $btn.AutoSize = $true
        $btn.Padding = New-Object System.Windows.Forms.Padding(6, 0, 6, 0)
        $btn.Height = 28
        $btn.Tag = [string]$openable[0].Value
        $btn.TabStop = $false
        $btn.Add_Click({ Start-Process $this.Tag })
        $pnlButtons.Controls.Add($btn)
    } elseif ($openable.Count -gt 1) {
        $btnOpen = New-Object System.Windows.Forms.Button
        $btnOpen.Text = "開く..."
        $btnOpen.AutoSize = $true
        $btnOpen.Padding = New-Object System.Windows.Forms.Padding(6, 0, 6, 0)
        $btnOpen.Height = 28
        $btnOpen.TabStop = $false
        $openMenu = New-Object System.Windows.Forms.ContextMenuStrip
        foreach ($field in $openable) {
            $item = New-Object System.Windows.Forms.ToolStripMenuItem
            $item.Text = "$($field.Name): $($field.Value)"
            $item.Tag = [string]$field.Value
            $item.Add_Click({ Start-Process $this.Tag })
            [void]$openMenu.Items.Add($item)
        }
        $btnOpen.Add_Click({
            $openMenu.Show($this, (New-Object System.Drawing.Point(0, $this.Height)))
        }.GetNewClosure())
        $pnlButtons.Controls.Add($btnOpen)
    }

    # CRUD ボタン（Excel テーブルのみ）
    if ($isExcel) {
        $sep = New-Object System.Windows.Forms.Label
        $sep.Text = "|"
        $sep.AutoSize = $true
        $sep.Padding = New-Object System.Windows.Forms.Padding(4, 6, 4, 0)
        $sep.ForeColor = [System.Drawing.Color]::LightGray
        $pnlButtons.Controls.Add($sep)

        $btnSave = New-Object System.Windows.Forms.Button
        $btnSave.Text = "保存"
        $btnSave.AutoSize = $true
        $btnSave.Height = 28
        $btnSave.TabStop = $false
        $btnSave.Add_Click({ Save-CurrentRecord })
        $pnlButtons.Controls.Add($btnSave)

        $btnAdd = New-Object System.Windows.Forms.Button
        $btnAdd.Text = "新規追加"
        $btnAdd.AutoSize = $true
        $btnAdd.Height = 28
        $btnAdd.TabStop = $false
        $btnAdd.Add_Click({ Add-NewRecord })
        $pnlButtons.Controls.Add($btnAdd)

        $btnDel = New-Object System.Windows.Forms.Button
        $btnDel.Text = "削除"
        $btnDel.AutoSize = $true
        $btnDel.Height = 28
        $btnDel.TabStop = $false
        $btnDel.ForeColor = [System.Drawing.Color]::DarkRed
        $btnDel.Add_Click({ Remove-CurrentRecord })
        $pnlButtons.Controls.Add($btnDel)
    }
}

# =============================================================================
# CRUD 操作
# =============================================================================
function Save-CurrentRecord {
    $tableName = $cmbTable.SelectedItem
    $idx = $listRecords.SelectedIndex
    if (-not $tableName -or $idx -lt 0) { return }

    $src = $script:tableSource[$tableName]
    if ($src.Type -ne "excel") { return }

    $recIdx = $script:filteredIndices[$idx]
    $edited = Get-EditedRecord $tableName

    try {
        Save-ExcelRecord -ExcelPath $src.File -SheetName $src.Sheet -RecordIndex $recIdx -Record $edited
        # メモリ上のデータも更新
        $script:allTables[$tableName][$recIdx] = $edited
        # リスト表示を更新（選択位置を維持）
        $savedIdx = $idx
        Update-RecordList
        if ($savedIdx -lt $listRecords.Items.Count) { $listRecords.SelectedIndex = $savedIdx }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("保存エラー:`n$_", "エラー", "OK", "Error")
    }
}

function Add-NewRecord {
    $tableName = $cmbTable.SelectedItem
    if (-not $tableName) { return }

    $src = $script:tableSource[$tableName]
    if ($src.Type -ne "excel") { return }

    $newRec = Get-EditedRecord $tableName

    try {
        Add-ExcelRecord -ExcelPath $src.File -SheetName $src.Sheet -Record $newRec
        # メモリ上にも追加
        $script:allTables[$tableName] += $newRec
        Update-RecordList
        # 最後の項目を選択
        if ($listRecords.Items.Count -gt 0) {
            $listRecords.SelectedIndex = $listRecords.Items.Count - 1
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("追加エラー:`n$_", "エラー", "OK", "Error")
    }
}

function Remove-CurrentRecord {
    $tableName = $cmbTable.SelectedItem
    $idx = $listRecords.SelectedIndex
    if (-not $tableName -or $idx -lt 0) { return }

    $src = $script:tableSource[$tableName]
    if ($src.Type -ne "excel") { return }

    $recIdx = $script:filteredIndices[$idx]
    $rec = $script:allTables[$tableName][$recIdx]
    $preview = @($rec.PSObject.Properties | Select-Object -First 2 | ForEach-Object { "$($_.Name): $($_.Value)" }) -join ", "

    $result = [System.Windows.Forms.MessageBox]::Show(
        "削除しますか?`n$preview", "確認", "YesNo", "Question"
    )
    if ($result -ne "Yes") { return }

    try {
        Remove-ExcelRecord -ExcelPath $src.File -SheetName $src.Sheet -RecordIndex $recIdx
        # メモリ上からも削除
        $list = [System.Collections.ArrayList]@($script:allTables[$tableName])
        $list.RemoveAt($recIdx)
        $script:allTables[$tableName] = @($list)
        Update-RecordList
    } catch {
        [System.Windows.Forms.MessageBox]::Show("削除エラー:`n$_", "エラー", "OK", "Error")
    }
}

# =============================================================================
# イベント
# =============================================================================
$cmbTable.Add_SelectedIndexChanged({
    $tableName = $cmbTable.SelectedItem
    $txtFilter.Clear()
    Build-FieldEditors $tableName
    Update-RecordList
})

$listRecords.Add_SelectedIndexChanged({ Update-Detail })

# Enter / ダブルクリック: 1つなら即開く、複数ならフィールド選択メニュー
$script:openAction = {
    $tableName = $cmbTable.SelectedItem
    $idx = $listRecords.SelectedIndex
    if (-not $tableName -or $idx -lt 0 -or $idx -ge $script:filteredIndices.Count) { return }
    $recIdx = $script:filteredIndices[$idx]
    $rec = $script:allTables[$tableName][$recIdx]
    $openable = @($rec.PSObject.Properties | Where-Object { $_.Name -match $script:openablePattern -and $_.Value })
    if ($openable.Count -eq 1) {
        Start-Process $openable[0].Value
    } elseif ($openable.Count -gt 1) {
        $menu = New-Object System.Windows.Forms.ContextMenuStrip
        foreach ($field in $openable) {
            $item = New-Object System.Windows.Forms.ToolStripMenuItem
            $item.Text = "$($field.Name): $($field.Value)"
            $item.Tag = [string]$field.Value
            $item.Add_Click({ Start-Process $this.Tag })
            [void]$menu.Items.Add($item)
        }
        $pt = $listRecords.PointToScreen(
            (New-Object System.Drawing.Point(0, ($listRecords.GetItemRectangle($idx).Bottom)))
        )
        $menu.Show($pt)
    }
}

$listRecords.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        & $script:openAction
        $_.Handled = $true
    }
})

$listRecords.Add_DoubleClick({ & $script:openAction })

$txtFilter.Add_TextChanged({ Update-RecordList })

$txtFilter.Add_KeyDown({
    if ($_.KeyCode -eq "Enter" -or $_.KeyCode -eq "Down") {
        $listRecords.Focus()
        if ($listRecords.Items.Count -gt 0 -and $listRecords.SelectedIndex -lt 0) {
            $listRecords.SelectedIndex = 0
        }
        $_.Handled = $true
    }
})

$form.Add_Shown({ $listRecords.Focus() })

# フォーム終了時に Excel COM を解放
$form.Add_FormClosed({ Close-ExcelApp })

# =============================================================================
# 初期表示
# =============================================================================
foreach ($name in $script:tableNames) {
    [void]$cmbTable.Items.Add($name)
}

if ($cmbTable.Items.Count -gt 0) {
    $cmbTable.SelectedIndex = 0
}

[void]$form.ShowDialog()
$form.Dispose()
