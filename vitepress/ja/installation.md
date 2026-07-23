# インストール

Windows 版の `xlflow` CLI と `.NET` Excel bridge をインストールし、Excel の Trust Center で **VBA プロジェクト オブジェクト モデルへのアクセスを信頼する** を有効にします。

```powershell
xlflow version
xlflow doctor --json
```

WSL から使う場合は Linux フロントエンドも入れ、プロジェクトを `/mnt/c` など Windows マウント配下に置きます。VS Code 拡張の設定で CLI が見つからない場合は `xlflow.path` に Windows 実行ファイルの絶対パスを指定します。

失敗時は [トラブルシューティング](./troubleshooting) と [英語の詳細手順](../installation) を参照してください。
