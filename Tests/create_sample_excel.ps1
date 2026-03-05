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

# --- employees (大量データ: 200件, 10フィールド) ---
$ws5 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
$ws5.Name = "employees"
$eHeaders = @("employee_id", "full_name", "department_name", "position_title", "email_address", "phone_number", "hire_date", "monthly_salary", "document_folder_path", "intranet_profile_url")
for ($c = 0; $c -lt $eHeaders.Count; $c++) { $ws5.Cells.Item(1, $c+1) = $eHeaders[$c] }

$depts = @("営業部", "開発部", "総務部", "企画部", "経理部", "人事部", "広報部", "法務部")
$positions = @("部長", "課長", "係長", "主任", "一般")
$lastNames = @("田中", "鈴木", "佐藤", "高橋", "渡辺", "伊藤", "山本", "中村", "小林", "加藤")
$firstNames = @("太郎", "花子", "一郎", "美咲", "健太", "由美", "大輔", "恵子", "翔太", "あかり")

for ($r = 0; $r -lt 200; $r++) {
    $id = "E{0:D4}" -f ($r + 1)
    $ln = $lastNames[$r % $lastNames.Count]
    $fn = $firstNames[[math]::Floor($r / $lastNames.Count) % $firstNames.Count]
    $fullName = "$ln$fn"
    $dept = $depts[$r % $depts.Count]
    $pos = $positions[$r % $positions.Count]
    $email = "emp{0:D4}@example.com" -f ($r + 1)
    $phone = "03-{0:D4}-{1:D4}" -f (1000 + $r), (5000 + $r)
    $hireDate = (Get-Date "2015-04-01").AddDays($r * 7).ToString("yyyy-MM-dd")
    $salary = 250000 + ($r * 5000)
    $path = "C:\Users\ynisi\Documents\employees\$id"
    $url = "https://intranet.example.com/profile/$id"

    $ws5.Cells.Item($r + 2, 1) = $id
    $ws5.Cells.Item($r + 2, 2) = $fullName
    $ws5.Cells.Item($r + 2, 3) = $dept
    $ws5.Cells.Item($r + 2, 4) = $pos
    $ws5.Cells.Item($r + 2, 5) = $email
    $ws5.Cells.Item($r + 2, 6) = $phone
    $ws5.Cells.Item($r + 2, 7) = $hireDate
    $ws5.Cells.Item($r + 2, 8) = $salary
    $ws5.Cells.Item($r + 2, 9) = $path
    $ws5.Cells.Item($r + 2, 10) = $url
}
Write-Host "employees: 200 records created"

$savePath = "C:\workspace\dev\tools\Data\master.xlsx"
$wb.SaveAs($savePath)
$wb.Close()
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Created: $savePath"
