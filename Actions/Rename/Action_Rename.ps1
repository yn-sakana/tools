# =============================================================================
# Action_Rename.ps1 - テンプレートベースのリネーム
# 同名の .json ファイルから設定を読み込む
#
# テンプレート変数:
#   {basename}  - 元のファイル名（拡張子なし）
#   {ext}       - 拡張子（ドット付き）
#   {original}  - 元のファイル名（拡張子付き）
#   {date}      - 日付（dateFormatに従う）
# =============================================================================

param(
    [Parameter(Mandatory = $true)]
    [string]$TargetPath
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $TargetPath)) {
    Write-Host "[Rename] ファイルが見つかりません: $TargetPath" -ForegroundColor Red
    exit 1
}

# 同名jsonから設定読み込み
$configPath = [System.IO.Path]::ChangeExtension($MyInvocation.MyCommand.Path, ".json")
if (-not (Test-Path $configPath)) {
    Write-Host "[Rename] 設定ファイルが見つかりません: $configPath" -ForegroundColor Red
    exit 1
}
$config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json

$dir = [System.IO.Path]::GetDirectoryName($TargetPath)
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
$ext = [System.IO.Path]::GetExtension($TargetPath)
$dateStr = Get-Date -Format $config.dateFormat

# テンプレート展開
$newName = $config.template
$newName = $newName -replace "\{basename\}", $baseName
$newName = $newName -replace "\{ext\}", $ext
$newName = $newName -replace "\{original\}", ([System.IO.Path]::GetFileName($TargetPath))
$newName = $newName -replace "\{date\}", $dateStr

$newPath = Join-Path $dir $newName

# 同名がある場合は連番
if (Test-Path $newPath) {
    $newBase = [System.IO.Path]::GetFileNameWithoutExtension($newName)
    $newExt = [System.IO.Path]::GetExtension($newName)
    $counter = 2
    while (Test-Path $newPath) {
        $newName = "{0}-{1}{2}" -f $newBase, $counter, $newExt
        $newPath = Join-Path $dir $newName
        $counter++
    }
}

Rename-Item -Path $TargetPath -NewName $newName
Write-Host "[Rename] $([System.IO.Path]::GetFileName($TargetPath)) -> $newName" -ForegroundColor Green
