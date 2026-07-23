# はじめに

新規プロジェクトは `xlflow new Book.xlsm`、既存ブックは `xlflow init Existing.xlsm` で開始します。

```powershell
xlflow doctor --json
xlflow pull --json
xlflow lint --json
xlflow macros --json
xlflow run Main.Run --diagnostic --json
```

`src/` が編集対象、`build/` が保存済みブックです。Excel 側を編集したら `pull`、ソースを編集したら `push`、セッションで未保存の変更を残したら `save --session` を使います。

詳しくは [既存ブックのチュートリアル](./existing-workbook) と [状態モデル](../concepts/workbook-session-source) を参照してください。
