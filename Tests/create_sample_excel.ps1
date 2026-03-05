$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Add()

# --- contacts ---
$ws = $wb.Sheets.Item(1)
$ws.Name = "contacts"
$headers = @("name", "department", "email", "phone")
for ($c = 0; $c -lt $headers.Count; $c++) { $ws.Cells.Item(1, $c+1) = $headers[$c] }
$data = @(
    @("田中太郎", "営業部", "tanaka@example.com", "03-1234-5678"),
    @("鈴木花子", "開発部", "suzuki@example.com", "03-2345-6789"),
    @("佐藤一郎", "総務部", "sato@example.com", "03-3456-7890")
)
for ($r = 0; $r -lt $data.Count; $r++) {
    for ($c = 0; $c -lt $data[$r].Count; $c++) { $ws.Cells.Item($r+2, $c+1) = $data[$r][$c] }
}

# --- bookmarks ---
$ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
$ws2.Name = "bookmarks"
$headers = @("name", "category", "url")
for ($c = 0; $c -lt $headers.Count; $c++) { $ws2.Cells.Item(1, $c+1) = $headers[$c] }
$data = @(
    @("Google", "検索", "https://www.google.com"),
    @("GitHub", "開発", "https://github.com"),
    @("Stack Overflow", "開発", "https://stackoverflow.com")
)
for ($r = 0; $r -lt $data.Count; $r++) {
    for ($c = 0; $c -lt $data[$r].Count; $c++) { $ws2.Cells.Item($r+2, $c+1) = $data[$r][$c] }
}

# --- folders ---
$ws3 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
$ws3.Name = "folders"
$headers = @("name", "purpose", "folder")
for ($c = 0; $c -lt $headers.Count; $c++) { $ws3.Cells.Item(1, $c+1) = $headers[$c] }
$data = @(
    @("ダウンロード", "ブラウザDL先", "C:\Users\ynisi\Downloads"),
    @("ツール本体", "tools プロジェクト", "C:\workspace\dev\tools")
)
for ($r = 0; $r -lt $data.Count; $r++) {
    for ($c = 0; $c -lt $data[$r].Count; $c++) { $ws3.Cells.Item($r+2, $c+1) = $data[$r][$c] }
}

# --- products ---
$ws4 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
$ws4.Name = "products"
$headers = @("code", "name", "price", "category")
for ($c = 0; $c -lt $headers.Count; $c++) { $ws4.Cells.Item(1, $c+1) = $headers[$c] }
$data = @(
    @("A001", "ウィジェットA", 1200, "部品"),
    @("A002", "ウィジェットB", 2400, "部品"),
    @("B001", "ガジェットX", 5800, "完成品")
)
for ($r = 0; $r -lt $data.Count; $r++) {
    for ($c = 0; $c -lt $data[$r].Count; $c++) { $ws4.Cells.Item($r+2, $c+1) = $data[$r][$c] }
}

$savePath = "C:\workspace\dev\tools\Data\master.xlsx"
$wb.SaveAs($savePath)
$wb.Close()
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Created: $savePath"
