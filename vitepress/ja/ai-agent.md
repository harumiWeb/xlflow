# AI エージェントで開発する

エージェントには、まず `status --json` と `doctor --json` を実行し、`src/` だけを編集し、`fmt`・`lint`・`analyze` の後に Excel を操作するルールを与えます。

```bash
xlflow session start --json
xlflow push --fast --session --no-save --json
xlflow test --session --json
xlflow run Main.Run --diagnostic --session --json
xlflow inspect workbook --session --json
xlflow export-image --session --json
xlflow save --session --json
xlflow session stop --json
```

JSON と終了コードを解析し、ダイアログを待ち続けないようにします。失敗時は `error.code` を確認して修正し、`workbook_recovery_required` なら通常ループを止めて recovery 手順に移ります。

再現可能なプロンプトと検証例は [英語版チュートリアル](../tutorials/ai-agent) を参照してください。
