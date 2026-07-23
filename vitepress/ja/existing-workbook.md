# 既存ブックを Git 管理する

```powershell
xlflow doctor --json
xlflow init C:\path\to\Existing.xlsm
xlflow pull --json
xlflow lint --json
xlflow push --json
xlflow run Main.Run --diagnostic --json
xlflow inspect workbook --json
git diff -- src xlflow.toml
```

`init` は元ファイルを直接変更せず `build/` にコピーします。VBE で編集した場合は `pull`、VS Code で編集した場合は `push` を行います。どちらが新しいか不明な場合は `status --json` とバックアップで確認してから上書きしてください。

全手順は [英語版チュートリアル](../tutorials/existing-workbook) にあります。
