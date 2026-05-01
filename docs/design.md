# xlflow の高速化設計

現状コードを見る限り、特に重そうなのは以下です。

- Go 側は毎回 PowerShell を起動して `.ps1` を実行している
- `push` は毎回 Excel を新規起動し、ブックを開き、既存 VBA コンポーネントを全 export バックアップし、非ドキュメントモジュールを全削除して全 import している
- `run` も毎回 Excel を新規起動し、ブックを開き、実行用 harness module を VBProject に注入してから `Excel.Run` している
- 最後にブック/Excel を閉じているため、次回も同じ初期化コストが発生する

`push` は現在、バックアップディレクトリと一時 import ディレクトリを作り、Excel COM でブックを開いたあと、全コンポーネントをバックアップし、非ドキュメントモジュールを削除してから全 import しています。これは安全ですが、差分が小さくても毎回フルリビルドに近い動きになります。

`run` も毎回 Excel COM を起動してブックを開き、VBIDE に harness module を追加して、実行後に削除しています。これも安定性は高いですが、反復実行ではかなり重いです。

## 一番効く改善は「daemon / session mode」です

最も効果が大きいのは、`xlflow daemon` または `xlflow session` の導入です。

現在はおそらく毎回こうです。

```txt
xlflow run
  -> powershell起動
  -> Excel.Application 起動
  -> workbook open
  -> harness inject
  -> macro run
  -> workbook close
  -> Excel quit
```

これを次のようにします。

```txt
xlflow session start
  -> Excel.Application 起動
  -> workbook open
  -> keep alive

xlflow run
  -> 既存Excel/既存workbookに対して macro run

xlflow push
  -> 既存Excel/既存workbookに対して source sync

xlflow session stop
  -> workbook close
  -> Excel quit
```

コマンド例はこうです。

```bash
xlflow session start
xlflow push
xlflow run Main.Run
xlflow run Main.Run
xlflow push
xlflow run Main.Run
xlflow session stop
```

これができると、1回ごとの Excel 起動・ブック open・close を省けます。
体感では、現在 30〜60秒かかる処理が、ケースによっては数秒〜十数秒まで落ちる可能性があります。

特に AIエージェントが `push -> run -> 修正 -> push -> run` を何度も繰り返す用途では、これはかなり効きます。

## 次に効くのは `push` の差分更新です

現状の `push` はかなり安全側です。
毎回バックアップして、既存コンポーネントを export して、削除して、再 import しています。

これは初期実装としては正しいですが、開発ループでは重いです。

改善するなら、`push` にモードを分けるとよいです。

```bash
xlflow push
xlflow push --fast
xlflow push --full
xlflow push --no-backup
xlflow push --changed-only
```

おすすめは、デフォルトは安全なままにして、開発時だけ高速モードを使えるようにすることです。

### `push --changed-only`

ファイルの hash を `.xlflow/state.json` に保存して、変更された `.bas` / `.cls` / `.frm` だけ反映します。

```json
{
  "components": {
    "Main.bas": {
      "hash": "abc123",
      "component": "Main",
      "type": "standard"
    }
  }
}
```

変更がないファイルは import しない。
変更があるファイルだけ remove/import する。

これだけでかなり速くなります。

### `push --no-backup`

開発中は毎回バックアップ不要な場面も多いです。

```bash
xlflow push --no-backup
```

または、

```bash
xlflow push --backup=auto
xlflow push --backup=always
xlflow push --backup=never
```

がよさそうです。

個人的にはこれが良いです。

```bash
xlflow push --fast
```

`--fast` の中身は、

```txt
- changed-only
- no full export backup
- no full component rebuild
```

にする。

ただし事故防止のため、`--fast` は `.xlflow/backups` に頼らない代わりに、Git 管理前提であることを明記した方がよいです。

## `run` は harness 注入を毎回やめると速くなります

`run.ps1` は毎回 VBProject に一時 module を追加し、そこに harness code を入れて、実行後に削除しています。

これは柔軟ですが、VBIDE 操作は遅いです。
しかも「VBA プロジェクト オブジェクト モデルへのアクセス」も必要になります。

改善案は2つあります。

## 案A: 永続 runner module を入れる

初回だけ `XlflowRunner.bas` を注入して、以後はそれを使い回します。

```bash
xlflow runner install
xlflow run Main.Run
xlflow runner remove
```

内部的には毎回 temporary module を作らず、

```vb
Application.Run "XlflowRunner.RunMacro", "Main.Run", argsJson
```

のようにする。

これにより、

```txt
毎回 add module
毎回 AddFromString
毎回 remove module
```

を省けます。

## 案B: 引数なし・単純実行なら harness なしで直接 `Excel.Run`

引数なしの macro であれば、そもそも harness module が不要な場合があります。

```powershell
$excel.Run($MacroName)
```

で済むなら、`--direct` を用意できます。

```bash
xlflow run Main.Run --direct
```

ただし、現在の harness はエラー捕捉や実行時間測定、trace 連携の役割を持っているはずなので、`--direct` は高速だが診断は弱いモードになります。

```bash
xlflow run Main.Run
xlflow run Main.Run --fast
xlflow run Main.Run --direct
```

という設計が良いです。

## PowerShell 起動コストも削れるが、優先度は少し下

現状 Go 側では、コマンドごとに `powershell -NoProfile -ExecutionPolicy Bypass -File ...` を起動しています。

これ自体も数百 ms 〜 数秒のコストになります。

ただ、1分かかる原因としては、PowerShell より Excel COM / workbook open / VBIDE 操作 / save の方が大きい可能性が高いです。

とはいえ、daemon 化するなら PowerShell を毎回起動しない構成にできます。

選択肢は3つです。

```txt
A. Go process が PowerShell を毎回起動する
   現状。単純だが遅い。

B. PowerShell daemon を常駐させる
   実装しやすい。Excel COM を保持できる。

C. Go から直接 COM を叩く
   高速化余地はあるが実装難度が上がる。
```

現実的には、まず **B: PowerShell daemon** がよいと思います。
今の `.ps1` 資産を活かしながら、Excel Application と Workbook を保持できます。

## 保存を減らすのも効きます

`push` は最後に `$workbook.Save()` しています。
`run` も `--save` 指定時は保存します。

Excel の保存はブックサイズやネットワークドライブ次第で重いです。
なので開発ループでは保存タイミングを制御できるとよいです。

```bash
xlflow push --no-save
xlflow save
```

ただし、`push --no-save` はブックを閉じたら消える変更になります。
daemon/session mode とセットなら有効です。

```bash
xlflow session start
xlflow push --no-save
xlflow run Main.Run
xlflow push --no-save
xlflow run Main.Run
xlflow save
xlflow session stop
```

これはかなり速い開発ループになります。

## おすすめ設計

最終的にはこういう階層がよいと思います。

### 安全モード

```bash
xlflow push
xlflow run Main.Run
```

特徴：

```txt
- 毎回ブックを開く
- 毎回バックアップ
- 毎回保存
- 診断重視
- CI/リリース前向け
```

### 高速開発モード

```bash
xlflow dev
```

または、

```bash
xlflow session start
xlflow push --fast
xlflow run Main.Run --fast
xlflow session stop
```

特徴：

```txt
- Excel/Workbookを保持
- 変更ファイルだけ反映
- バックアップを省略または軽量化
- harnessを使い回す
- 保存を必要時だけ行う
```

### CI/検証モード

```bash
xlflow check
xlflow push --full
xlflow run Main.Run --timeout 60
```

特徴：

```txt
- 再現性重視
- クリーンな Excel 起動
- フルバックアップ
- フル同期
```

## 実装優先度

私ならこの順でやります。

```txt
1. push --no-backup
2. push --changed-only
3. run --direct / run --fast
4. persistent runner module
5. session/daemon mode
6. push --no-save + explicit xlflow save
```

ただし、効果が一番大きいのは `session/daemon mode` です。
実装が軽い順なら `push --no-backup` と `run --direct` が先です。

## 現実的な落としどころ

短期ではこれが良いです。

```bash
xlflow push --fast
xlflow run Main.Run --fast
```

`push --fast` は、

```txt
- バックアップ省略
- 変更ファイルのみ import
- 必要なコンポーネントだけ remove/import
```

`run --fast` は、

```txt
- 引数なしなら Excel.Run を直接呼ぶ
- trace なし
- harness 注入なし、または永続 runner を利用
```

中期でこれです。

```bash
xlflow session start
xlflow push --fast
xlflow run Main.Run --fast
xlflow session stop
```

この方向に進めば、「1回1分かかるけど待てるようにする」から、
**「そもそも反復開発ループを速くする」** に進化できます。

特に xlflow は AIエージェント向けの開発ハーネスなので、`push -> run -> 修正` のループが速いことはかなり重要です。
体験価値としては、keepalive 改善と同じか、それ以上に重要だと思います。
