# tools 利用手順

## セットアップ
1. `tools` フォルダを任意の場所に配置
2. `config.json` のパスを自分の環境に合わせて編集
3. PowerShellで `Main.ps1` を実行

## 使い方

### 起動
```
powershell -ExecutionPolicy Bypass -File Main.ps1
```

### CLIコマンド
- `ftree [パス]` - フォルダツリーを表示
- `search <キーワード>` - エクスポート済みデータを検索
- `status` - 動作状況を確認
- `reload` - 設定を再読み込み
- `help` - ヘルプ表示
- `quit` - 終了

### 右クリックメニュー
Main.ps1 動作中のみ有効。ファイルを右クリックすると登録した処理が表示される。

### フォルダ監視
Main.ps1 動作中、config.json で指定したフォルダを監視。
ファイルが追加されると自動で仕分けを実行。

## VBAアドイン

### OutlookMailSort
1. Outlook の Alt+F11 でVBAエディタを開く
2. `VBA\OutlookMailSort.bas` をインポート
3. `SortMailsByContact` を実行

### ExcelExporter
1. 対象のExcelブックで Alt+F11
2. `VBA\ExcelExporter.bas` をインポート
3. ThisWorkbook に BeforeSave イベントを追加（コード内のコメント参照）

### ExcelOperationLog
1. 監視したいExcelブックで Alt+F11
2. `VBA\ExcelOperationLog.bas` をインポート
3. ThisWorkbook にイベントハンドラを追加（コード内のコメント参照）

## 終了
CLIで `quit` と入力。右クリックメニューが自動解除され、監視が停止する。
