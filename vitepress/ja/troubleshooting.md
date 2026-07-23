# トラブルシューティング

まず次を保存します。

```powershell
xlflow version --verbose --json
xlflow doctor --json
xlflow status --json
```

| 症状                          | 確認・対処                                                 |
| ----------------------------- | ---------------------------------------------------------- |
| コマンドが見つからない        | PATH、Windows bridge、VS Code の `xlflow.path` を確認      |
| Excel/VBIDE に接続できない    | Trust Center 設定後に `doctor` を再実行                    |
| ソースとブックが不一致        | `status` で新しい側を確認し、必要なら `pull` または `push` |
| マクロが見つからない          | `macros --json` の `qualified_name` を使う                 |
| ダイアログやタイムアウト      | `run --diagnostic`、セッション状態、UserForm を確認        |
| WSL から Excel を操作できない | `/mnt/c` 配下へ移動し、Windows 実行ファイルを設定          |

詳細な原因・診断・検証手順は [英語版](../help/troubleshooting) にあります。
