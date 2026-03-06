# shinsa

交付金審査用のローカル作業ツールです。

## 入口

```powershell
.\shinsa.bat
```

起動時に自動で以下を実行します。

- OneDrive 同期元から clone を更新
- index を再生成
- 対話モードに入る

## 使い方

起動後は `shinsa>` で短い単語コマンドを使います。

- `gui`
  - 検索・閲覧・審査 GUI を開く
- `mail`
  - Outlook 取込 + clone 更新 + index 更新
- `sync`
  - clone 更新 + index 更新
- `index`
  - index を再生成
- `writeback`
  - レビュー内容を OneDrive 側台帳へ反映
- `status`
  - 状態表示
- `config`
  - 設定ファイルの場所を表示
- `quit`
  - 終了

## 設定

最初に [config.local.json](C:\workspace\dev\tools\shinsa\config\config.local.json) を環境に合わせて編集します。

主に直す項目

- `paths.onedriveLedgerRoot`
- `paths.onedriveCaseRoot`
- `paths.mailSourceRoot`

## 補足

- SharePoint は直接読まず、OneDrive 同期済みローカルパスを使います
- Outlook 取込はローカル Outlook を使います
- GUI から `Sync` `Mail` `Writeback` も実行できます
