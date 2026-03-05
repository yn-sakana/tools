# =============================================================================
# Main.ps1 - 基幹スクリプト
# 起動すると右クリックメニュー登録・フォルダ監視・CLI対話ループが全て有効になる
# 終了すると全てクリーンアップされる
# =============================================================================

param(
    [switch]$DryRun  # テスト用：レジストリ変更やファイル移動を実行しない
)

$ErrorActionPreference = "Stop"
$script:BasePath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:Config = $null
$script:Watchers = @()
$script:RegisteredMenuKeys = @()

# --- 共通関数 ---------------------------------------------------------------

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"
    Write-Host $line
    $logDir = Join-Path $script:BasePath "Logs"
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $logFile = Join-Path $logDir ("log_{0}.txt" -f (Get-Date -Format "yyyyMMdd"))
    Add-Content -Path $logFile -Value $line -Encoding UTF8
}

function Load-Config {
    $configPath = Join-Path $script:BasePath "config.json"
    if (-not (Test-Path $configPath)) {
        Write-Log "config.json が見つかりません: $configPath" "ERROR"
        exit 1
    }
    $script:Config = Get-Content $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
    Write-Log "config.json を読み込みました"
}

# --- Win11 従来メニュー強制 --------------------------------------------------

$script:ClassicMenuRegKey = "Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InProcServer32"

function Enable-ClassicContextMenu {
    if ($DryRun) {
        Write-Log "[DryRun] 従来コンテキストメニュー設定をスキップ" "INFO"
        return
    }

    $hkcu = [Microsoft.Win32.Registry]::CurrentUser
    $existing = $hkcu.OpenSubKey($script:ClassicMenuRegKey)
    if ($existing) {
        $existing.Close()
        Write-Log "従来コンテキストメニュー: 既に設定済み"
    } else {
        Set-RegistryKey -KeyPath $script:ClassicMenuRegKey -Value ""
        Write-Log "従来コンテキストメニューを有効化（反映にはエクスプローラー再起動が必要）"
        Write-Host ""
        Write-Host "  従来コンテキストメニューを初回設定しました。" -ForegroundColor Yellow
        Write-Host "  反映するにはタスクマネージャーからエクスプローラーを再起動してください。" -ForegroundColor Yellow
        Write-Host ""
    }
}

# --- 右クリックメニュー登録/解除 ---------------------------------------------

function Set-RegistryKey {
    param([string]$KeyPath, [string]$Value)
    $hkcu = [Microsoft.Win32.Registry]::CurrentUser
    $key = $hkcu.CreateSubKey($KeyPath)
    if ($Value) { $key.SetValue("", $Value) }
    $key.Close()
}

function Remove-RegistryKey {
    param([string]$KeyPath)
    $hkcu = [Microsoft.Win32.Registry]::CurrentUser
    try { $hkcu.DeleteSubKeyTree($KeyPath, $false) } catch {}
}

function Register-ContextMenus {
    if ($DryRun) {
        Write-Log "[DryRun] 右クリックメニュー登録をスキップ" "INFO"
        return
    }

    $menuCount = $script:Config.contextMenu.Count
    for ($i = 0; $i -lt $menuCount; $i++) {
        $menu = $script:Config.contextMenu[$i]
        $keyName = "MyAuto_{0:D2}_{1}" -f $i, ([System.IO.Path]::GetFileNameWithoutExtension($menu.action))
        $actionScript = Join-Path $script:BasePath "Actions\$($menu.action).ps1"
        $cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$actionScript`" `"%1`""

        $isFirst = ($i -eq 0)
        $isLast = ($i -eq $menuCount - 1)

        foreach ($target in @("file", "folder")) {
            if ($menu.target -ne $target -and $menu.target -ne "both") { continue }
            $root = if ($target -eq "file") { "Software\Classes\*\shell" } else { "Software\Classes\Directory\shell" }
            $regKey = "$root\$keyName"
            Set-RegistryKey -KeyPath $regKey -Value $menu.label
            Set-RegistryKey -KeyPath "$regKey\command" -Value $cmd
            $hkcu = [Microsoft.Win32.Registry]::CurrentUser
            $key = $hkcu.OpenSubKey($regKey, $true)
            if ($isFirst) { $key.SetValue("SeparatorBefore", "") }
            if ($isLast) { $key.SetValue("SeparatorAfter", "") }
            $key.Close()
            $script:RegisteredMenuKeys += $regKey
        }

        Write-Log "右クリックメニュー登録: $($menu.label) -> $($menu.action)"
    }
}

function Unregister-ContextMenus {
    if ($DryRun) {
        Write-Log "[DryRun] 右クリックメニュー解除をスキップ" "INFO"
        return
    }

    foreach ($regKey in $script:RegisteredMenuKeys) {
        Remove-RegistryKey -KeyPath $regKey
        Write-Log "右クリックメニュー解除: $regKey"
    }
    $script:RegisteredMenuKeys = @()
}

# --- フォルダ監視 -------------------------------------------------------------

function Start-FolderWatchers {
    foreach ($watch in $script:Config.watchFolders) {
        if (-not (Test-Path $watch.path)) {
            Write-Log "監視対象フォルダが存在しません: $($watch.path)" "WARN"
            continue
        }

        $watcher = New-Object System.IO.FileSystemWatcher
        $watcher.Path = $watch.path
        $watcher.Filter = $watch.filter
        $watcher.EnableRaisingEvents = $true

        $actionScript = Join-Path $script:BasePath "Actions\$($watch.action).ps1"

        $msgData = @{
            ActionScript = $actionScript
            DryRun = [bool]$DryRun
            LastEvents = @{}
        }

        $handler = {
            $path = $Event.SourceEventArgs.FullPath
            $now = [DateTime]::Now
            $data = $Event.MessageData

            if ($data.LastEvents[$path] -and ($now - $data.LastEvents[$path]).TotalMilliseconds -lt 500) {
                return
            }
            $data.LastEvents[$path] = $now

            Start-Sleep -Milliseconds 500

            Write-Host ""
            Write-Host "[監視] ファイル検知: $path" -ForegroundColor Magenta
            if (-not $data.DryRun) {
                try {
                    & $data.ActionScript $path
                } catch {
                    Write-Host "[監視] エラー: $_" -ForegroundColor Red
                }
            } else {
                Write-Host "[DryRun] 実行スキップ: $($data.ActionScript) $path"
            }
            Write-Host "tools>" -NoNewline -ForegroundColor Green
        }

        $job1 = Register-ObjectEvent -InputObject $watcher -EventName "Created" -Action $handler -MessageData $msgData
        $job2 = Register-ObjectEvent -InputObject $watcher -EventName "Renamed" -Action $handler -MessageData $msgData

        $script:Watchers += @{ Watcher = $watcher; Jobs = @($job1, $job2) }
        Write-Log "フォルダ監視開始: $($watch.path) (filter: $($watch.filter))"
    }
}

function Stop-FolderWatchers {
    foreach ($w in $script:Watchers) {
        $w.Watcher.EnableRaisingEvents = $false
        $w.Watcher.Dispose()
        foreach ($job in $w.Jobs) {
            if ($job) {
                Unregister-Event -SourceIdentifier $job.Name -ErrorAction SilentlyContinue
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            }
        }
    }
    $script:Watchers = @()
    Write-Log "フォルダ監視を全て停止しました"
}

# --- CLIツール読み込み --------------------------------------------------------

function Load-CLITools {
    $toolsDir = Join-Path $script:BasePath "CLI_Tools"
    if (Test-Path $toolsDir) {
        $toolFiles = Get-ChildItem -Path $toolsDir -Filter "Tool_*.ps1"
        foreach ($toolFile in $toolFiles) {
            . $toolFile.FullName
            Write-Log "CLIツール読み込み: $($toolFile.Name)"
        }
    }
}

# --- CLI対話ループ ------------------------------------------------------------

function Show-Help {
    Write-Host ""
    Write-Host "=== コマンド一覧 ===" -ForegroundColor Cyan
    Write-Host "  ftree [path]       フォルダツリーを表示（省略時はカレント）"
    Write-Host "  search <keyword>   エクスポート済みデータを検索"
    Write-Host "  index [table] [key] テーブル検索（引数なしで一覧）"
    Write-Host "  status             監視状況・登録メニューの確認"
    Write-Host "  reload             config.json を再読み込み"
    Write-Host "  help               このヘルプを表示"
    Write-Host "  quit               終了（クリーンアップ実行）"
    Write-Host ""
}

function Show-Status {
    Write-Host ""
    Write-Host "=== ステータス ===" -ForegroundColor Cyan
    Write-Host "  基幹スクリプト: 動作中"
    Write-Host "  DryRunモード: $DryRun"
    Write-Host "  監視中のフォルダ: $($script:Watchers.Count) 件"
    foreach ($w in $script:Watchers) {
        Write-Host "    - $($w.Watcher.Path)"
    }
    Write-Host "  登録済み右クリックメニュー: $($script:RegisteredMenuKeys.Count) 件"
    foreach ($key in $script:RegisteredMenuKeys) {
        Write-Host "    - $key"
    }
    Write-Host ""
}

function Read-HostNonBlocking {
    Write-Host "tools>" -NoNewline -ForegroundColor Green
    $line = ""
    while ($true) {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq "Enter") {
                Write-Host ""
                return $line
            } elseif ($key.Key -eq "Backspace") {
                if ($line.Length -gt 0) {
                    $line = $line.Substring(0, $line.Length - 1)
                    Write-Host "`b `b" -NoNewline
                }
            } else {
                $line += $key.KeyChar
                Write-Host $key.KeyChar -NoNewline
            }
        } else {
            Start-Sleep -Milliseconds 50
        }
    }
}

function Start-CLILoop {
    Show-Help

    while ($true) {
        $userInput = Read-HostNonBlocking

        if ([string]::IsNullOrWhiteSpace($userInput)) { continue }

        $parts = $userInput -split "\s+", 2
        $cmd = $parts[0].ToLower()
        $arg = if ($parts.Count -gt 1) { $parts[1] } else { "" }

        switch ($cmd) {
            "ftree" {
                $targetPath = if ($arg) { $arg } else { Get-Location }
                Invoke-FolderTree -Path $targetPath -ExcludeDirs $script:Config.folderTree.excludeDirs
            }
            "search" {
                if (-not $arg) {
                    Write-Host "使い方: search <キーワード>" -ForegroundColor Yellow
                } else {
                    Invoke-DataSearch -Keyword $arg -DataPath $script:Config.dataPath
                }
            }
            "index" {
                $indexParts = $arg -split "\s+", 2
                $tableName = if ($indexParts[0]) { $indexParts[0] } else { "" }
                $indexKey = if ($indexParts.Count -gt 1) { $indexParts[1] } else { "" }
                Invoke-IndexSearch -TableName $tableName -Key $indexKey -DataPath $script:Config.dataPath
            }
            "status" {
                Show-Status
            }
            "reload" {
                Load-Config
                Write-Host "設定を再読み込みしました" -ForegroundColor Cyan
            }
            "help" {
                Show-Help
            }
            "quit" {
                return
            }
            default {
                Write-Host "不明なコマンド: $cmd （help で一覧表示）" -ForegroundColor Yellow
            }
        }
    }
}

# --- メイン処理 ---------------------------------------------------------------

try {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  tools - 起動中..." -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    Load-Config

    # CLIツールをスクリプトスコープで読み込み
    $toolsDir = Join-Path $script:BasePath "CLI_Tools"
    if (Test-Path $toolsDir) {
        $toolFiles = Get-ChildItem -Path $toolsDir -Filter "Tool_*.ps1"
        foreach ($toolFile in $toolFiles) {
            . $toolFile.FullName
            Write-Log "CLIツール読み込み: $($toolFile.Name)"
        }
    }

    Enable-ClassicContextMenu
    Register-ContextMenus
    Start-FolderWatchers

    Write-Log "=== tools 起動完了 ==="

    Start-CLILoop
}
finally {
    Write-Host ""
    Write-Log "=== クリーンアップ開始 ==="
    Stop-FolderWatchers
    Unregister-ContextMenus
    Write-Log "=== tools 終了 ==="
    Write-Host ""
}
