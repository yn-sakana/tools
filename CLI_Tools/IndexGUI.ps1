# =============================================================================
# IndexGUI.ps1 - index GUI版 (Windows Forms)
# プルダウンでテーブル選択 → リストでレコード選択 → 詳細 + 開くボタン
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

# --- config ---
$script:Config = Get-Content (Join-Path $script:BasePath "config.json") -Raw -Encoding UTF8 | ConvertFrom-Json
$script:openablePattern = '^(path|file|folder|dir|directory|url|link|href|uri)$'

# --- データ読み込み ---
$script:allTables = [ordered]@{}
$dataPath = $script:Config.dataPath
if (Test-Path $dataPath) {
    $dataFiles = @(Get-ChildItem -Path $dataPath -Recurse | Where-Object { $_.Extension -match '^\.(json|csv)$' })
    foreach ($f in $dataFiles) {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
        $records = @()
        switch ($f.Extension.ToLower()) {
            ".json" {
                $d = Get-Content $f.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
                if ($d -is [System.Array]) { $records = $d } else { $records = @($d) }
            }
            ".csv" { $records = @(Import-Csv $f.FullName -Encoding UTF8) }
        }
        $script:allTables[$name] = $records
    }
}
$script:tableNames = @($script:allTables.Keys)
$script:filteredIndices = @()

# =============================================================================
# フォーム
# =============================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "index"
$form.Size = New-Object System.Drawing.Size(700, 600)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Yu Gothic UI", 10)
$form.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)

# --- 上部: テーブル選択 + フィルタ ---
$pnlTop = New-Object System.Windows.Forms.TableLayoutPanel
$pnlTop.Dock = "Top"
$pnlTop.Height = 68
$pnlTop.ColumnCount = 2
$pnlTop.RowCount = 2
$pnlTop.Padding = New-Object System.Windows.Forms.Padding(6, 6, 6, 2)
[void]$pnlTop.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("AutoSize")))
[void]$pnlTop.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100)))
[void]$pnlTop.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 50)))
[void]$pnlTop.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 50)))

$lblTable = New-Object System.Windows.Forms.Label
$lblTable.Text = "テーブル:"
$lblTable.AutoSize = $true
$lblTable.Anchor = "Left"
$pnlTop.Controls.Add($lblTable, 0, 0)

$cmbTable = New-Object System.Windows.Forms.ComboBox
$cmbTable.DropDownStyle = "DropDownList"
$cmbTable.Dock = "Fill"
$pnlTop.Controls.Add($cmbTable, 1, 0)

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = "絞り込み:"
$lblFilter.AutoSize = $true
$lblFilter.Anchor = "Left"
$pnlTop.Controls.Add($lblFilter, 0, 1)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Dock = "Fill"
$pnlTop.Controls.Add($txtFilter, 1, 1)

# --- 中央: レコードリスト ---
$listRecords = New-Object System.Windows.Forms.ListBox
$listRecords.Dock = "Fill"
$listRecords.Font = New-Object System.Drawing.Font("Consolas", 10)
$listRecords.IntegralHeight = $false

# --- 下部: ボタン + 詳細 ---
$pnlDetail = New-Object System.Windows.Forms.Panel
$pnlDetail.Dock = "Bottom"
$pnlDetail.Height = 200

$splitterDetail = New-Object System.Windows.Forms.Splitter
$splitterDetail.Dock = "Bottom"
$splitterDetail.Height = 5

$pnlButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlButtons.Dock = "Top"
$pnlButtons.Height = 38
$pnlButtons.FlowDirection = "LeftToRight"
$pnlButtons.Padding = New-Object System.Windows.Forms.Padding(4, 4, 0, 0)

$txtDetail = New-Object System.Windows.Forms.TextBox
$txtDetail.Dock = "Fill"
$txtDetail.Multiline = $true
$txtDetail.ReadOnly = $true
$txtDetail.ScrollBars = "Vertical"
$txtDetail.Font = New-Object System.Drawing.Font("Consolas", 10)
$txtDetail.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
$txtDetail.TabStop = $false

$pnlDetail.Controls.Add($txtDetail)
$pnlDetail.Controls.Add($pnlButtons)

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Dock = "Bottom"
$lblCount.Height = 22
$lblCount.TextAlign = "BottomRight"
$lblCount.ForeColor = [System.Drawing.Color]::DimGray
$lblCount.Padding = New-Object System.Windows.Forms.Padding(0, 0, 4, 0)

# フォーム組み立て（Dock順: Fill → Bottom → Top）
$form.Controls.Add($listRecords)
$form.Controls.Add($splitterDetail)
$form.Controls.Add($pnlDetail)
$form.Controls.Add($lblCount)
$form.Controls.Add($pnlTop)

# =============================================================================
# ロジック
# =============================================================================
function Update-RecordList {
    $tableName = $cmbTable.SelectedItem
    if (-not $tableName) {
        $listRecords.Items.Clear()
        $txtDetail.Clear()
        $pnlButtons.Controls.Clear()
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
        $txtDetail.Clear()
        return
    }

    $recIdx = $script:filteredIndices[$idx]
    $rec = $script:allTables[$tableName][$recIdx]

    # 詳細テキスト
    $lines = @()
    foreach ($prop in $rec.PSObject.Properties) {
        $lines += "{0,-15} : {1}" -f $prop.Name, $prop.Value
    }
    $txtDetail.Text = $lines -join "`r`n"

    # 開くボタン（全ての openable フィールドにボタン生成）
    foreach ($prop in $rec.PSObject.Properties) {
        if ($prop.Name -match $script:openablePattern -and $prop.Value) {
            $btn = New-Object System.Windows.Forms.Button
            $btn.Text = "$($prop.Name) を開く"
            $btn.AutoSize = $true
            $btn.Padding = New-Object System.Windows.Forms.Padding(6, 0, 6, 0)
            $btn.Height = 28
            $btn.Tag = [string]$prop.Value
            $btn.TabStop = $false
            $btn.Add_Click({ Start-Process $this.Tag })
            $pnlButtons.Controls.Add($btn)
        }
    }
}

# =============================================================================
# イベント
# =============================================================================
$cmbTable.Add_SelectedIndexChanged({
    $txtFilter.Clear()
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

# ボタン間のTab移動後、Escでリストに戻る
$pnlButtons.Add_KeyDown({
    if ($_.KeyCode -eq "Escape") {
        $listRecords.Focus()
        $_.Handled = $true
    }
})

$form.Add_Shown({ $listRecords.Focus() })

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
