# =============================================================================
# Action_Move.ps1 - ルールベースのファイル移動
# 同名の .json ファイルから設定を読み込む
# ファイル名が pattern にマッチしたら destination へ移動
# 同名ファイルがあれば連番を付ける
# =============================================================================

param(
    [Parameter(Mandatory = $true)]
    [string]$TargetPath
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $TargetPath)) {
    Write-Host "[Move] ファイルが見つかりません: $TargetPath" -ForegroundColor Red
    exit 1
}

# 同名jsonから設定読み込み
$configPath = [System.IO.Path]::ChangeExtension($MyInvocation.MyCommand.Path, ".json")
if (-not (Test-Path $configPath)) {
    Write-Host "[Move] 設定ファイルが見つかりません: $configPath" -ForegroundColor Red
    exit 1
}
$config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json

$fileName = [System.IO.Path]::GetFileName($TargetPath)
$matched = $false

foreach ($rule in $config.rules) {
    if ($fileName -match $rule.pattern) {
        $matched = $true
        $destDir = $rule.destination

        if (-not (Test-Path $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }

        $destPath = Join-Path $destDir $fileName

        # 同名ファイルがあれば連番
        if (Test-Path $destPath) {
            $base = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
            $ext = [System.IO.Path]::GetExtension($fileName)
            $counter = 2
            while (Test-Path $destPath) {
                $destPath = Join-Path $destDir ("{0}-{1}{2}" -f $base, $counter, $ext)
                $counter++
            }
        }

        Move-Item -Path $TargetPath -Destination $destPath -Force
        Write-Host "[Move] $fileName -> $destPath" -ForegroundColor Green
        break
    }
}

if (-not $matched) {
    Write-Host "[Move] マッチするルールがありません: $fileName" -ForegroundColor Yellow
}
