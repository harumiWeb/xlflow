# VS Code 拡張

拡張は xlflow CLI と VBA Language Server を VS Code に統合します。補完、`CreateObject` の型推論、Hover、定義移動、リアルタイム診断、CodeLens 実行、Testing UI、pull/push/session 操作を利用できます。

1. CLI と bridge をインストールする。
2. `xlflow.toml` のあるフォルダーを開く。
3. `xlflow: Check Environment` を実行する。
4. 補完や診断が表示されない場合は Language Server を再起動し、`xlflow.path` と `xlflow.lsp.logFile` を確認する。

## UserForm Designer のプロパティ編集

設定した forms root（標準は `src/forms`）直下の `specs` にある `.yaml`、
`.yml`、`.json` の FormSpec を開き、**Reopen Editor With... → xlflow
UserForm Designer** を選びます。背景はフォーム、コントロールはその対象、
Page タブは Page のプロパティを選択します。Property Grid は接続中 LSP の
メタデータから、対応する scalar フィールドを表示します。

文字列・数値は Enter またはフォーカスを外して確定し、Escape で取り消します。
boolean／enum は選択時に確定します。一確定は一つのソース編集となり、VS Code の
undo/redo を利用できます。不正な入力はエラーとともに残り、ソースは変更しません。
外部のソース変更で古い入力が失効した場合は、更新後の文書に対して再入力します。

未設定・明示的 null・空文字・false・0 は別の状態です。観測値や描画の補完値を
authoring 値として書き戻しません。null は対応する optional フィールドだけに設定でき、
削除を意味しません。未設定に戻すにはテキストエディターでキーを削除します。
legacy フィールドと `build.*` は別々に編集します。`form.name` の変更はファイル名、
コード sidecar、VBA 参照を改名しません。Page の geometry、observed、コレクション、
picture、任意 property bag は Property Grid の編集対象外です。

プロパティ編集には `userFormPropertyEdit` capability が必要です。古い geometry 対応
サーバーでは移動・リサイズを維持し、preview のみのサーバーでは読み取り専用です。
ソースが曖昧な場合は有効なプレビューを保ったままプロパティ編集を無効にします。
FormSpec のソース編集に Excel は不要で、ブックへ自動適用はしません。

[英語の VS Code 詳細](../vscode/) には設定一覧と機能別の制約を掲載しています。
