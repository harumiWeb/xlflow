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

## UserForm Designer の Toolbox と構造編集

Toolbox には Pointer と、Label、TextBox、ComboBox、ListBox、CommandButton、
CheckBox、OptionButton、ToggleButton、SpinButton、ScrollBar、Image、Frame、
MultiPage、TabStrip の14種類を表示します。Page は一覧に含めません。
種類を選んでキャンバスをクリックすると、フォーム直下にコントロールを追加します。
ドラッグ追加やFrame／Page内への配置は行いません。描画と同じ種類別の既定サイズを使い、
フォーム内に収まるよう位置を補正します。フォームが既定サイズより小さい場合は追加を拒否します。

追加したコントロールには`control-<UUID>` IDと、既存コントロール名およびフォーム名に
大文字・小文字を区別せず重複しない最小の正整数suffixを使った名前を割り当てます。名前を変更してもIDは変わりません。
MultiPage はPageを持たず`selectedIndex: -1`で作成し、TabStrip は`tabs: []`と
`selectedIndex: -1`で作成します。追加に成功すると新しいコントロールを選択し、Toolboxを
Pointerに戻します。配置中のEscapeは追加を取り消します。

選択中のコントロールはツールバー、またはキャンバスにフォーカスがあるときのDeleteで削除できます。
子孫がないコントロールはそのまま削除します。子孫がある場合だけ対象名と子孫数を表示して確認し、
対象と子孫を一つのundo/redo可能な編集で削除します。入力欄の文字Deleteは横取りせず、Pageの直接削除は無効です。
追加・削除には`userFormStructuralEdit` capabilityが必要です。未対応サーバーでも、既存の
geometry／property編集は各capabilityが示す範囲で利用できます。ソース編集は自動保存せず、
code-behindも変更しません。

Frameへの配置・親変更は[Issue #920](https://github.com/harumiWeb/xlflow/issues/920)、
PageやTabの編集は[Issue #921](https://github.com/harumiWeb/xlflow/issues/921)の範囲です。

[英語の VS Code 詳細](../vscode/) には設定一覧と機能別の制約を掲載しています。
