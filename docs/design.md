<!-- 設計メモ -->

# xlflow-native テスト UX 案

## 背景

- 現状の `xlflow test` は、名前規約に一致する引数なし `Sub` を見つけて順に実行する軽量な仕組みになっている。
- scaffold 時に `tests` フォルダは作られるが、`push` のソース列挙対象ではなく、実質的に dead folder になっている。
- Rubberduck 互換をそのまま追うと、private テスト実行、COM 依存の Assert/Fakes、注釈駆動の複雑なライフサイクルなど、xlflow の headless / JSON / session 中心の設計と噛み合わない部分がある。

## 目標

- Rubberduck 互換そのものではなく、xlflow により適合したテスト開発体験を提供する。
- headless 実行、JSON 出力、UI 応答注入、session workflow と自然に統合されたテスト UX を作る。
- AI agent / CI / unattended 実行でも同じ作法で回せることを優先する。

## 採用方針

- `tests` フォルダは廃止し、テストコードは `src` 配下に同居させる。
- 推奨配置は `src/modules/Tests/` とし、filesystem 上の配置でテスト群を整理する。
- Rubberduck 互換は目的にしない。
- ただし、開発者体験として有用な概念は xlflow-native な形で再構成して取り込む。

## デザイン原則

### 1. Push / Pull / Test の整合を最優先する

- source of truth は常に `src` 配下に置く。
- テストだけ別ディレクトリ・別 import ルールにしない。
- push 対象と test 対象のズレを避ける。

### 2. 注釈より unattended 実行性を優先する

- IDE 専用の注釈互換より、CLI から確実に発見・実行・観測できることを優先する。
- テスト結果は人間だけでなくエージェントや CI からも扱いやすい構造を維持する。

### 3. 魔法を増やしすぎない

- 特別な hidden rule より、明示的な命名規約と少数の hook を使う。
- private テスト実行のために不自然なコード注入や Rubberduck 特有の前提を持ち込まない。

## 想定するテスト配置

- 標準モジュールのテストは `src/modules/Tests/` 配下に置く。
- 必要に応じて `src/modules/Tests/Smoke/` や `src/modules/Tests/Integration/` のようにサブディレクトリで整理する。
- `vba.folders = true` の既定動作と組み合わせ、path ベースの folder annotation 管理に寄せる。

例:

```txt
src/
  modules/
    Main.bas
    ReportService.bas
    Tests/
      ReportServiceTests.bas
      Smoke/
        AppSmokeTests.bas
```

## テスト発見ルール

- 基本のテスト本体は public な引数なし `Sub` とする。
- 命名規約は当面の互換性を保ち、`Test*` または `*_Test` を対象にする。
- `Private Sub` はテスト本体としては扱わない。
- private procedure はテストヘルパーや共通処理として使う。

この方針により、現在の軽量 discovery との互換を保ちつつ、runner 側の実装コストを抑える。

## ライフサイクル hook

Rubberduck の `ModuleInitialize` / `TestInitialize` をそのまま模倣するのではなく、xlflow-native な reserved name として次を導入する。

- `BeforeAll`
- `AfterAll`
- `BeforeEach`
- `AfterEach`

ルール:

- いずれも public な引数なし `Sub` とする。
- 同一モジュール内でのみ有効にする。
- 存在すれば runner が自動実行する。
- 実行順は `BeforeAll` -> 各テストごとに `BeforeEach` -> `Test...` -> `AfterEach` -> 最後に `AfterAll` とする。
- `BeforeEach` や `AfterEach` が失敗した場合も、そのテストケースの失敗として JSON に記録する。

この方式なら、private 呼び出し問題を避けつつ setup / cleanup 体験を提供できる。

## アサーション API

`XlflowAssert` を拡張して、Rubberduck の開発体験に近いが xlflow に最適化した最小セットを提供する。

候補:

- `AssertEquals expected, actual, [message]`
- `AssertNotEqual expected, actual, [message]`
- `AssertTrue condition, [message]`
- `AssertFalse condition, [message]`
- `AssertFail [message]`
- `AssertInconclusive [message]`
- `AssertIsNothing value, [message]`
- `AssertIsNotNothing value, [message]`

方針:

- まず scalar-friendly な API を優先する。
- object / array は初期段階では限定的に扱い、誤解しやすい比較は避ける。
- 失敗は明確な VBA error として raise し、既存の JSON envelope へ自然に乗せる。

## Filter / Tag

将来的には test 名だけでなく、module 名や tag による絞り込みを扱えるようにする。

候補:

- `xlflow test --filter TestCreateReport`
- `xlflow test --module ReportServiceTests`
- `xlflow test --tag smoke`

tag は Rubberduck 互換を目的にしない軽量 comment metadata として設計する。
例:

```vb
'@Tag("smoke")
Public Sub TestSmoke()
End Sub
```

ここでの comment metadata は IDE 依存ではなく、xlflow runner が読む CLI 用メタデータとして扱う。

## xlflow らしい付加価値

Rubberduck 互換よりも、以下の点で上回る UX を目指す。

### 1. Headless UI テスト

- `XlflowUI` と `--msgbox` / `--inputbox` / `--filedialog` を組み合わせ、対話 UI を unattended にテストできる。
- `--ui-stream` により、実行中の UI 分岐をリアルタイムで観測できる。

### 2. JSON で扱いやすい失敗情報

- 失敗した test 名、module 名、duration、error code、message を構造化して返す。
- 将来的には hook failure、inconclusive、tag、debug event も JSON に含められるようにする。

### 3. Session workflow との一体化

- `xlflow push --fast --session --no-save` -> `xlflow test --session --json` を標準の反復サイクルにする。
- live workbook が disk より新しいケースも、session-aware に扱う。

### 4. AI agent / CI 親和性

- 人間向け IDE 機能ではなく、CLI と structured output を中心に据える。
- エージェントが failing test をピンポイントで再実行しやすいインターフェースを保つ。

## 非目標

- Rubberduck COM の `AssertClass` や `FakesProvider` への依存
- Rubberduck 注釈との完全互換
- private test method をそのまま実行する高度な dispatcher
- IDE 側だけで成立する機能を xlflow の中核 UX に据えること

## 実装の段階案

### Phase 1

- `tests` フォルダ scaffold を廃止
- `src/modules/Tests/` を推奨配置としてドキュメント化
- `XlflowAssert` を拡張

### Phase 2

- `BeforeAll` / `AfterAll` / `BeforeEach` / `AfterEach` hook を runner に追加
- hook failure を JSON へ反映

### Phase 3

- `--module` / `--tag` filter を追加
- 軽量 comment metadata を導入

### Phase 4

- 必要であれば scaffold にテストテンプレート生成を追加
- `xlflow new test` のような補助コマンドを検討

## まとめ

- xlflow のテスト UX は、Rubberduck 互換ではなく、`src` 同居・headless 実行・JSON 可観測性・session 中心の体験として設計する。
- setup / cleanup の便利さは hook で取り込み、IDE 固有の複雑さは持ち込まない。
- 目指すのは「Rubberduck のように書けること」ではなく、「xlflow 上で速く、壊れにくく、無人でも同じように回せること」である。
