# セットアップ手順

## 1. PowerShell プロファイルにエイリアスを追加

PowerShellを開いて以下を実行:

```powershell
# プロファイルが無ければ作成
if (!(Test-Path $PROFILE)) { New-Item -Path $PROFILE -Force }

# メモ帳で開く
notepad $PROFILE
```

以下の行を追加して保存:

```powershell
function tools { powershell -ExecutionPolicy Bypass -File "C:\workspace\dev\tools\Main.ps1" @args }
```

PowerShellを再起動すれば `tools` で起動できる。
DryRunモードは `tools -DryRun` で起動。

## 2. パスを変えた場合

`config.json` 内のパスを自分の環境に合わせて編集する。

## 3. ExecutionPolicy について

会社PCで `$PROFILE` の読み込み自体がブロックされている場合は、
PowerShellを以下で起動する:

```
powershell -ExecutionPolicy Bypass
```

この状態なら `$PROFILE` も読み込まれ、`tools` コマンドが使える。
