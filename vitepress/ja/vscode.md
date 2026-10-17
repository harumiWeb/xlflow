# VS Code 拡張

拡張は xlflow CLI と VBA Language Server を VS Code に統合します。補完、`CreateObject` の型推論、Hover、定義移動、リアルタイム診断、CodeLens 実行、Testing UI、pull/push/session 操作を利用できます。

1. CLI と bridge をインストールする。
2. `xlflow.toml` のあるフォルダーを開く。
3. `xlflow: Check Environment` を実行する。
4. 補完や診断が表示されない場合は Language Server を再起動し、`xlflow.path` と `xlflow.lsp.logFile` を確認する。

## UserForm Designer のワークスペース

YAML／YML／JSONのソースエディターを右クリックし、「xlflow: Designerで開く」から
Designerを開けます。対象は設定されたforms rootの`specs`直下にあるFormSpecです。
開くだけでは未保存のソースを保存しません。

Canvas上部にはUndo／Redoボタンもあります。VS Code標準のショートカットと同じ履歴を使い、
操作できない場合や保存・競合・ダイアログ表示中は無効になります。

Canvas 上部のコンパクトな操作バーで、切取・コピー・貼付・削除を操作できます。
整列・サイズは同じ親の2個以上、均等配置は3個以上の選択で表示します。
重なり順はメニューにまとめ、複製とプロパティへの移動は「その他の操作」から使えます。
Grid／Snap／Zoom は画面下部にあります。

右ペインは「プロパティ」「構造」「Tab順」の3タブです。「構造」では重なった、
不可視のコントロールも名前で選択できます。非アクティブPage配下の選択では、
ソースの selectedIndex を変更せず該当Pageを表示します。矢印キーで移動・展開し、
Enter／Spaceで選択します。ツリーのドラッグによる親変更は行いません。
幅900px未満では操作バーのボタンからToolboxまたは右ペインを開きます。

「ソースを隣に開く」はDesignerとテキストエディターを並べて表示します。
Designerの変更は下書きとして保持し、ヘッダーの保存ボタンまたはDesigner上の
Ctrl+S／Cmd+Sでソースへ反映します。ビューやタブの操作だけでは反映しません。

## Designerからイベントコードを開く

標準コントロールやフォーム背景をダブルクリックすると、既定イベントのコードを開きます。
TextBoxは`Change`、CommandButton・ListBox・UserForm本体は`Click`、
MultiPageは`Change`が対象です。選択要素の操作メニューや構造ツリーの右クリックから
「イベントを開く…」を選ぶと、`Initialize`など別のイベントも選択できます。
手続きがあれば本文へ移動し、なければVBEと同じ宣言で空の手続きを追加します。

Designerまたは対応するYAML／YML／JSONに未保存変更があれば、保存して続行するか
キャンセルするかを選びます。生成したコードは未保存のエディターで開き、通常の
コード編集のUndoで取り消せます。Excelの起動やフォームのビルドは行いません。

編集先は`xlflow.toml`の`[userform].code_source`に従います。
`sidecar`では`<forms-root>/code/<FormName>.bas`を編集し、なければ作成します。
`frm`では既存の対応する`.frm`のコード部分を編集します。フォーム名や対応ファイルが
不一致・重複している場合は、修正するまでコードを生成しません。
コード側の未保存変更は保持し、既存の手続き宣言を自動修正しません。
要素の改名・削除でも、既存イベントコードは自動変更しません。

Page・Tabの項目は所有するMultiPage／TabStripのイベントを開きます。
Imageは確認したVBEの挙動に合わせて`BeforeDragOver`を開き、非表示の
`Click`／`DblClick`は候補に出しません。独自ActiveXは今回の対象外です。
この機能には対応する新しいxlflow LSPが必要です。

`form build`は`frm`設定でもsidecarコードを使用します。frmコードの反映には
既存のfrm方式のpushを使うか、明示的にsidecar方式へ移行してください。

## サイドバーからフォームをビルド

フォーム行のビルドボタンまたは右クリックで個別に、フォームビュー上部の
「全フォームをビルド」でまとめて、FormSpecをExcelブックに反映できます。
実行前に対象ブックとフォームを確認し、既存フォームの上書きを承認します。
対象のDesigner・ソース・コードビハインドに未保存変更があれば、
「保存してビルド」またはキャンセルを選びます。無関係なファイルは保存しません。

Specがないフォームはビルド対象外です。重複するSpecは解消してから実行します。
全フォームは一つのExcelセッションで順番に反映し、全件成功後にブックを一度保存します。
既存セッションは開いたままにします。途中失敗時は後続を止め、成功・失敗・未実行を案内します。
成功済みの変更を自動で戻すことはせず、未保存変更のあるセッションを保持します。

## UserForm Designer のプロパティ編集

標準の `src/forms/specs` にある `.yaml`、`.yml`、`.json` は Designer で開きます。
別の forms root を設定している場合は、ソースエディターを右クリックして
**xlflow: Designerで開く** を選びます。Designer からテキストエディターへ戻すには
**Reopen Editor With...** を使います。
背景はフォーム、コントロールはその対象、Page タブは Page を選択します。
Property Grid は接続中 LSP のメタデータから名前／値の2列を表示します。
プロパティ名は言語設定にかかわらず英語、カテゴリ・説明・操作は表示言語に従います。
文字列・数値は Enter またはフォーカスを外して確定し、Escape で取り消します。
boolean はチェックボックス、enum は ComboBox で編集します。

Font ダイアログでは名前、ポイント単位のサイズ、装飾、数値の charset を指定し、
現在値・代表的な候補・自由入力を利用できます。Color は RGB、system、OLE を区別し、
system 色の意味や未知の OLE 値を保持します。ダイアログと Canvas に一時表示し、
フォントの代替表示や system 色は Excel と異なる場合があります。対応項目は Go/LSP が
決定し、custom ActiveX や個々の Tab レコードに型付きEditorは追加しません。
Page は Font／ForeColor／BackColor／BorderStyle、CommandButton は TextAlign の対象外です。
TextProps のフォントサイズは 0.05 pt 単位です。空のフォント名や既存 StdFont への
非ASCII名の更新は拒否し、未編集の不透明なフォント情報は保持します。

Image はネイティブのファイル選択で BMP／JPEG を指定できます。プロジェクト内の画像は
相対パスを利用し、外部画像は OK 時だけ設定した forms root の
`assets/<full-lowercase-SHA256>.bmp|jpg` へ取り込みます。同一内容は再利用し、
Undo 後も Redo のため画像を保持します。Cancel は画像を作りません。
明示的な削除だけが `picture: { remove: true }` になります。PictureAlignment と
PictureSizeMode も編集できます。Position は単一コントロールの Left／Top／Width／Height、
フォームの Width／Height が対象で、変更した軸だけを一括確定します。Page・複数選択は対象外です。

開くだけ・入力中・Cancel では Designer の下書きや履歴を変更しません。OK は一つの
undo/redo 可能な編集です。ソースへの反映は引き続き Designer の明示的な保存で行います。
ソース変更や LSP 再起動時はダイアログ入力を残し、一時プレビューを解除して OK を無効化します。
Cancel して開き直してください。確定失敗時も入力とエラーを保持します。
legacy property bag は確定した項目だけを型付き表現へ移行し、他の項目を保持します。

Designer に型・`null` の選択欄はなく、各プロパティに具体的な値を入力します。
既存の未設定・明示的 `null` は、その項目の具体値を確定するまで保持します。
開く・フォーカスするだけでは既定値を書き込まず、空文字・false・0 は有効な値です。
任意の scalar を許す項目は既存の具体的な型を維持し、未設定・`null` なら文字列で入力します。
入力内容から別の型を推測しません。FormSpec 自体の `null` 対応は変わりません。
`null` の設定や、キーを削除して未設定に戻す操作はテキストエディターで行います。
Page の geometry、observed、任意 property bag の直接編集は対象外です。
`form.name` の変更はファイル、コード sidecar、VBA 参照を改名しません。

scalar 編集には `userFormPropertyEdit`、型付きEditorには version 1 の対応が必要です。
旧サーバーでは通知された既存の scalar／geometry 機能を維持します。ソースが曖昧な場合は
有効なプレビューを保ったまま編集を無効にします。FormSpec 編集に Excel は不要で、
ブックへ自動適用はしません。

## UserForm Designer の Toolbox と構造編集

Toolbox には Pointer と、Label、TextBox、ComboBox、ListBox、CommandButton、
CheckBox、OptionButton、ToggleButton、SpinButton、ScrollBar、Image、Frame、
MultiPage、TabStrip の14種類を表示します。Page は一覧に含めません。
種類を選んでキャンバスをクリックすると、フォーム直下または対象 Frame 内にコントロールを追加します。
ドラッグ追加は行いません。描画と同じ種類別の既定サイズを使い、
配置先に収まるよう位置を補正します。配置先が既定サイズより小さい場合は追加を拒否します。

追加したコントロールには`control-<UUID>` IDと、既存コントロール名およびフォーム名に
大文字・小文字を区別せず重複しない最小の正整数suffixを使った名前を割り当てます。名前を変更してもIDは変わりません。
MultiPage はPage1／Page2と`selectedIndex: 0`、TabStrip はTab1／Tab2と
`selectedIndex: 0`で作成します。既存の明示的な空コレクションは保持します。追加に成功すると新しいコントロールを選択し、Toolboxを
Pointerに戻します。配置中のEscapeは追加を取り消します。

選択中のコントロールはツールバー、またはキャンバスにフォーカスがあるときのDelete／Backspaceで削除できます。
「構造」で選択している場合もBackspaceを使えます。文字入力・IME変換・キーの長押し中は削除しません。
子孫がないコントロールはそのまま削除します。子孫がある場合だけ対象名と子孫数を表示して確認し、
対象と子孫を一つのundo/redo可能な編集で削除します。入力欄の文字Deleteは横取りせず、Pageの直接削除は無効です。
追加・削除には`userFormStructuralEdit` capabilityが必要です。未対応サーバーでも、既存の
geometry／property編集は各capabilityが示す範囲で利用できます。ソース編集は自動保存せず、
code-behindも変更しません。

コントロールを Frame 内またはフォーム背景へ通常ドラッグすると、親を変更します。
配置先の強調表示とプレビューで確定位置を確認できます。座標は新しい親の座標系へ変換し、
グリッドと境界補正を適用します。配置先より大きい場合はドロップを拒否します。
Frame の移動では子も一緒に動きます。リサイズとキーボード移動は現在の親を維持します。
キャンバス上のパンくずから祖先 Frame をクリックして選択できます。

Frame 配置には追加で `userFormParentEdit`、ドラッグによる親変更には親編集と geometry の
両 capability が必要です。旧サーバーではルート配置と現在の親内の移動を維持します。
Page を直接の親とする関係の変更や Page／Tab の編集は
[Issue #921](https://github.com/harumiWeb/xlflow/issues/921)の範囲です。

[英語の VS Code 詳細](../vscode/) には設定一覧と機能別の制約を掲載しています。
