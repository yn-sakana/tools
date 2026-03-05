# tools - オレオレRPA

会社PCの制約（管理者権限なし、外部ツール不可）の中で動く、Windows標準機能だけの自動化環境。

PowerShellを起動するだけで、右クリックメニュー・フォルダ監視・CLIコマンドが全部使える。終了すれば全部消える。


## クイックスタート

### 1. プロファイルにエイリアス登録

```powershell
if (!(Test-Path $PROFILE)) { New-Item -Path $PROFILE -Force }
notepad $PROFILE
```

以下を追加して保存:

```powershell
function tools { powershell -ExecutionPolicy Bypass -File "C:\workspace\dev\tools\Main.ps1" @args }
```

### 2. 起動

```
tools          # CLI対話モード（右クリック・監視・全コマンド）
tools -DryRun  # DryRunモード
```


## CLI コマンド

| コマンド | 説明 |
|---|---|
| `ftree [path]` | フォルダツリー表示 |
| `index [table]` | 対話型テーブル検索（矢印キー選択・フィルタ・開く） |
| `index gui` | テーブル検索 GUI版（Windows Forms） |
| `status` | 監視状況・登録メニューの確認 |
| `reload` | config.json 再読み込み |
| `quit` | 終了（クリーンアップ実行） |


## 技術スタック

| 技術 | 役割 |
|---|---|
| **PowerShell** | 基幹スクリプト、OS操作、CLI、GUI |
| **VBA** | Excel / Outlook のOffice操作 |
| **Power Automate** | Teams / SharePoint のクラウド連携 |
| **CSV / JSON** | VBA↔PowerShell のデータ受け渡し |

**原則:** OS側→PowerShell、Office内部→VBA、クラウド→Power Automate。VBAとPowerShellは直接繋がず、CSV/JSONファイルでデータ連携。


## ディレクトリ構成

```
tools/
├─ Main.ps1                 # 基幹スクリプト（対話モード / 単発コマンド）
├─ config.json              # 全体設定
│
├─ Actions/                 # 処理アドイン（ps1 + json ペア）
│   ├─ Rename/              #   タイムスタンプ付きリネーム
│   └─ Move/                #   ルールベースのファイル仕分け
│
├─ CLI_Tools/               # コマンドの実体
│   ├─ Tool_FolderTree.ps1  #   ftree
│   ├─ Tool_Index.ps1       #   index CLI版（対話型セレクターUI）
│   └─ IndexGUI.ps1         #   index GUI版（Windows Forms）
│
├─ Data/                    # テーブルデータ（JSON/CSV）
│   ├─ contacts.json        #   連絡先
│   ├─ products.json        #   商品マスタ
│   ├─ bookmarks.json       #   ブックマーク（URL）
│   ├─ folders.json         #   フォルダショートカット
│   ├─ mails.json           #   メールインデックス
│   └─ logs.csv             #   操作ログ
│
├─ VBA/                     # VBAモジュール（配布用）
│   ├─ OutlookMailSort.bas  #   メール振り分け・添付保存
│   ├─ ExcelExporter.bas    #   テーブル→CSV/JSON出力
│   └─ ExcelOperationLog.bas#   Excel操作ログ
│
├─ Tests/                   # テスト用フォルダ・ファイル
└─ Logs/                    # 実行ログ
```


## 設計思想

1. **アドイン化** — Actions/ にps1+jsonペアを置くだけで機能追加
2. **設定の一元管理** — 環境依存データはconfig.jsonに外出し
3. **起動で全部入り、終了でクリーン** — レジストリ残骸なし
4. **CLI/GUI並行** — 同じDataフォルダを参照、用途で使い分け


## indexコマンドの仕組み

Data/ フォルダのJSON/CSVを自動認識してテーブル一覧を表示。レコードに以下のフィールド名があれば「開く」アクションが使える:

`path` `file` `folder` `dir` `url` `link` `href` `uri`

- 1つ → `[O]` で即開く
- 複数 → セレクターで選択して開く


## パスを変えた場合

`config.json` 内のパスを自分の環境に合わせて編集。

## ExecutionPolicy

会社PCで `$PROFILE` がブロックされている場合:

```
powershell -ExecutionPolicy Bypass
```
