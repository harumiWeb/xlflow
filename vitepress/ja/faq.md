# FAQ

## 元のブックは変更されますか？

`init` は `build/` にコピーを作ります。設定した workbook path を明示的に変更しない限り、元ファイルは対象になりません。

## Excel は必須ですか？

`fmt`、`lint`、`analyze` などソース中心の操作には不要です。push、run、test、session、COM 検査には Windows Excel が必要です。

## `pull` と `push` はどちらを使いますか？

Excel で編集したら `pull`、VS Code やエージェントで編集したら `push` です。不明な場合は `status` とバックアップを先に使います。

## 問題を報告するには？

バージョン、doctor/status の JSON、正確なコマンド、終了コード、ブック形式、セッション状態を添えて、機密情報を除いた最小再現例を報告してください。
