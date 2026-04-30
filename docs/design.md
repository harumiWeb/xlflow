xlflow skill は、**コマンドリファレンス**ではなく、AIエージェント用の **VBA開発ワークフロー規約** として作るのが良いです。

理想はこうです。

```text
xlflow skill = AIエージェントがExcelVBAを安全に編集・実行・検証・修正するための作業手順書
```

## 基本構成

おすすめはこの構成です。

```text
skills/
  xlflow/
    SKILL.md
    examples/
      add-vba-module.md
      fix-failing-test.md
      create-macro-from-request.md
      inspect-existing-workbook.md
    snippets/
      trace-module.bas
      test-assert.bas
```

中心は `SKILL.md` です。

---

## SKILL.md に書くべき内容

### 1. まず読むべき前提

```md
# xlflow Skill

Use xlflow when working with Excel VBA projects from the command line.

The agent must not assume VBA code is correct until it has:

1. exported or inspected the workbook,
2. run relevant tests or macros,
3. checked the resulting workbook state,
4. reviewed diagnostics and diffs.
```

日本語ならこうです。

```md
# xlflow Skill

Excel VBAを編集する場合は、コード生成だけで完了扱いにしない。
必ず xlflow を使って、インポート、実行、検証、差分確認まで行う。
```

---

### 2. 標準ワークフロー

ここが一番重要です。

```md
## Standard workflow

When modifying an Excel VBA workbook:

1. Inspect the workbook.
   - Run `xlflow inspect <book>`
   - Run `xlflow export <book> --out vba/`

2. Modify VBA source files only.
   - Prefer editing `.bas`, `.cls`, `.frm` files.
   - Do not directly modify binary `.xlsm` unless necessary.

3. Import changes.
   - Run `xlflow import <book> --src vba/ --out build/out.xlsm`

4. Run diagnostics.
   - Run `xlflow doctor build/out.xlsm`
   - Run `xlflow lint vba/`

5. Execute tests.
   - Run `xlflow test build/out.xlsm`

6. If no tests exist, run the target macro.
   - Run `xlflow run build/out.xlsm <MacroName> --trace`

7. Compare workbook state.
   - Run `xlflow diff fixtures/expected.xlsm build/out.xlsm`
   - Or run `xlflow snapshot/assert`

8. Fix errors and repeat.
```

ポイントは、**AIに「編集して終わり」を禁止する**ことです。

---

### 3. エラー時の判断ルール

AIエージェントにはここがかなり効きます。

```md
## Failure handling

If `xlflow test` fails:

- Read the failing test name.
- Read the VBA error number and description.
- Inspect the related module.
- Make the smallest code change.
- Re-run only the failing test first.
- Then run the full test suite.

If `xlflow run --trace` fails:

- Read the trace log from top to bottom.
- Identify the last successful trace entry.
- Add more `XlflowLog` calls around the suspected block.
- Re-run the macro.
```

これを入れると、エージェントが闇雲に修正しにくくなります。

---

## コマンド別の使わせ方

skillでは、コマンド一覧より **いつ使うか** を書くのが重要です。

### `xlflow doctor`

```md
Use `xlflow doctor` before blaming VBA code.

Run it when:

- Excel COM fails
- macro execution fails before entering VBA code
- import/export fails
- tests behave inconsistently
```

### `xlflow test`

```md
Use `xlflow test` as the primary correctness signal.

Never mark the task complete if tests fail.
If there are no tests, create minimal tests when practical.
```

### `xlflow run --trace`

```md
Use `xlflow run --trace` when:

- implementing a new macro
- debugging runtime behavior
- no formal tests exist
- the macro mutates workbook state
```

### `xlflow diff`

```md
Use `xlflow diff` after running macros that modify a workbook.

Check:

- changed cells
- added/removed sheets
- formulas
- named ranges
- shapes
- print areas
```

---

## skillに入れるべき判断フロー

かなりおすすめです。

```md
## Decision flow

### User asks to add or change VBA behavior

1. Export existing VBA.
2. Locate relevant modules.
3. Edit source files.
4. Import into a copied workbook.
5. Run lint.
6. Run tests.
7. If tests are missing, run the target macro with trace.
8. Diff the output workbook.
9. Report final result with commands executed.

### User reports a VBA bug

1. Reproduce with `xlflow run --trace` or `xlflow test`.
2. Identify the failing module/procedure.
3. Add temporary trace logs if needed.
4. Patch the smallest area.
5. Re-run reproduction command.
6. Run full test suite.
7. Remove unnecessary temporary logs unless useful.
```

---

## コード生成ルールも入れるべき

VBAはAIが雑に書くと壊れやすいので、skill側に規約を書きます。

```md
## VBA coding rules

- Always use `Option Explicit`.
- Avoid `Select`, `Activate`, `ActiveSheet`, and unqualified `Range`.
- Prefer explicit workbook and worksheet references.
- Use `Long` instead of `Integer`.
- Restore Application state in cleanup blocks.
- Use `On Error GoTo ErrHandler`, not broad `On Error Resume Next`.
- Keep macros small and split logic into testable procedures.
- Put business logic in standard modules, not worksheet event handlers when possible.
```

特に重要なのはこれです。

```vb
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Report")
ws.Range("A1").Value = "..."
```

AIにこう書かせたい。

---

## trace用の規約もskill化する

````md
## Trace logging

When debugging macro execution, insert trace logs at:

- procedure entry
- important branch decisions
- row/column counts
- external file paths
- before and after destructive operations
- error handlers

Example:

```vb
Call XlflowLog("start GenerateReport")
Call XlflowLog("lastRow=" & lastRow)
Call XlflowLog("creating report sheet")
Call XlflowLog("finished GenerateReport")
```
````

````

さらに、ログを入れっぱなしにしてよいもの/消すものも分けると良いです。

```md
Keep trace logs if they describe high-level macro progress.
Remove noisy temporary logs before finalizing.
````

---

## 最終報告フォーマットを固定する

エージェントに最後こう報告させると、人間がレビューしやすいです。

```md
## Final response format

Report:

- Changed files/modules
- Commands executed
- Test result
- Macro execution result
- Workbook diff summary
- Remaining risks
```

例：

```md
Changed:

- Module `ReportGenerator`
- Added `TestReportGenerator`

Verified:

- `xlflow lint vba/`
- `xlflow test build/out.xlsm`
- `xlflow diff fixtures/expected.xlsm build/out.xlsm`

Result:

- All tests passed.
- Output workbook matches expected snapshot.
```

---

## 最初のxlflow skillはこれくらいで十分

初期版の `SKILL.md` は、以下の章立てが良いです。

```md
# xlflow Skill

## Purpose

## When to use xlflow

## Standard workflow

## Command usage

## VBA coding rules

## Testing rules

## Trace logging rules

## Failure handling

## Workbook diff rules

## Final response format
```

## 結論

xlflow skill は、単なる

```text
xlflow run はマクロを実行します
xlflow test はテストを実行します
```

では弱いです。

作るべきなのは、

```text
ExcelVBAをAIが編集するときの安全な開発ループ
```

です。

特に重要なのはこの4つです。

```text
1. 編集後は必ず import / run / test / diff する
2. 失敗時は trace を追加して再現する
3. ActiveSheet / Select / Activate を禁止する
4. 最終報告に実行コマンドと検証結果を含める
```

この形にすると、xlflow本体のハーネス群をAIエージェントがかなり自然に使えるようになります。

# xlflow skill の配布方法

おすすめは **両方対応** ですが、主導線は **xlflow本体のセットアップ時に配置** が良いです。

## 結論

```text
第一導線: xlflow init --with-skill
第二導線: npx/gh などの外部skill manager対応
```

が一番現実的です。

## 1. 最初は `xlflow init/new` に含めるべき

理由はシンプルで、xlflowのskillはツール本体と強く結びつくからです。

```bash
xlflow init
```

または

```bash
xlflow init --with-skill
```

で、プロジェクトにこう配置します。

```text
project/
  .xlflow/
    config.toml
  skills/
    xlflow/
      SKILL.md
      examples/
      snippets/
```

または Claude/Codex などを意識するなら：

```text
project/
  .claude/
    skills/
      xlflow/
        SKILL.md
```

```text
project/
  .codex/
    skills/
      xlflow/
        SKILL.md
```

この方式のメリットは大きいです。

- ユーザーが迷わない
- xlflowのバージョンとskill内容を揃えやすい
- プロジェクト固有のルールを書き足せる
- エージェントがリポジトリ内で確実に読める
- オフライン/社内環境でも使いやすい

特にあなたの想定ユーザーは **業務Excel + 社内PC + AIエージェント** なので、外部skill manager前提は少し弱いです。

---

## 2. ただし上書きは避ける

一度配置したskillは、ユーザーが改変する可能性があります。

なので更新はこうした方が良いです。

```bash
xlflow skill install
xlflow skill update
xlflow skill diff
```

例：

```bash
xlflow skill install --target .claude/skills
xlflow skill update --check
xlflow skill diff
```

上書きではなく、差分確認方式が安全です。

```text
skills/xlflow/SKILL.md
  local: modified
  bundled: v0.2.0
  action: show diff / merge manually
```

---

## 3. `npx skills` / `gh skills` は後でよい

外部skill manager対応は、将来的には価値があります。

ただし初期から主導線にする必要は薄いです。

理由は：

- まだskill managerの標準が固まっていない
- ユーザーのAIエージェント環境がバラバラ
- 企業環境では `npx` や外部取得が制限されやすい
- xlflow本体とskillのバージョン同期が面倒

なので、外部配布は **補助導線** が良いです。

```bash
npx skills install xlflow
```

や

```bash
gh skills install harumiWeb/xlflow
```

に対応するとしても、まずは後回しで問題ないです。

---

## 4. 個人的な推奨設計

一番良いのはこれです。

```text
xlflow本体に公式skillテンプレートを同梱
↓
xlflow init / skill install でプロジェクトにコピー
↓
プロジェクト側でカスタマイズ可能
↓
将来的に外部skill registryにも公開
```

コマンド案：

```bash
xlflow skill install
xlflow skill install --agent claude
xlflow skill install --agent codex
xlflow skill install --target ./skills
xlflow skill update
xlflow skill doctor
```

配置先を選べると強いです。

```bash
xlflow skill install --agent claude
```

```text
.claude/skills/xlflow/SKILL.md
```

```bash
xlflow skill install --agent codex
```

```text
.codex/skills/xlflow/SKILL.md
```

```bash
xlflow skill install --target skills
```

```text
skills/xlflow/SKILL.md
```

---

## 5. 配布形態としては3段階が良い

### Phase 1

```text
xlflow本体に同梱
```

これで十分です。

### Phase 2

```text
xlflow skill install/update/diff
```

skillを明示的に管理できるようにする。

---

## 6. READMEでの見せ方

READMEではこう書くと分かりやすいです。

```bash
xlflow init --with-skill
```

```text
This installs the xlflow AI agent skill into your project, including recommended workflows for editing, testing, tracing, and validating Excel VBA from the command line.
```

つまり、単なる補助ドキュメントではなく、

```text
AIエージェント用の公式開発ガイド
```

として見せるのが良いです。

## 最終判断

最初から外部skill managerに寄せるより、まずは：

```bash
xlflow init --with-skill
```

または

```bash
xlflow skill install
```

で配布するのがベストです。

xlflowは「CLIツール + AIエージェント運用」が価値なので、skillも本体の一部として扱った方が自然です。
