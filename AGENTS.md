# AIエージェント向けガイド

## 0. プロジェクト概要

```txt
root: .
├── bridge/
│   └── dotnet/ # dotnet bridge のコード
├── cmd/
├── docs/ # ADR、仕様書、その他開発ドキュメント
├── internal/ # 内部パッケージ
├── scripts/ # 自動化スクリプト
├── tasks/ # タスク管理と学習記録
├── vitepress/ # ユーザードキュメント
├── .editorconfig
├── .goreleaser.yaml
├── AGENTS.md
├── CHANGELOG.md
├── CLAUDE.md
├── CONTRIBUTING.md
├── global.json
├── go.mod
├── go.sum
├── lefthook.yml
├── LICENSE
├── package.json
├── pnpm-lock.yaml
├── pnpm-workspace.yaml
├── PSScriptAnalyzerSettings.psd1
├── README.ja.md
├── README.md
├── SECURITY.md
├── Taskfile.yml
└── THIRD_PARTY_LICENCES.md
```

## 1. ワークフロー設計

### 1. 基本は Plan モードで進める

- 3ステップ以上に分かれる作業、またはアーキテクチャに影響する作業は、必ず Plan モードから開始すること
- 途中で進行がうまくいかなくなった場合は、無理に続行せず、いったん止めて計画を立て直すこと
- Plan モードは実装時だけでなく、検証手順の設計にも使うこと
- 実装前に仕様をできるだけ具体化し、曖昧さを減らすこと

### 2. マルチエージェント戦略

- メインのコンテキストを汚さないために、サブエージェントを積極的に活用すること
- 調査、確認、並列分析はサブエージェントへ委譲すること
- 複雑な問題では、計算資源を多く使う目的でもサブエージェントを活用すること
- 実行を集中させるため、サブエージェントには1つのタスクだけを割り当てること
- 読み取り中心のコードベース探索には explorer を使うこと
- 実装や修正には worker を使うこと
- レビューには reviewer を使うこと

### 3. 自己改善ループ

- ユーザーから修正指示を受けたら、そのパターンを `tasks/lessons.md` に記録すること
- 同じミスを繰り返さないためのルールを、自分向けに明文化すること
- エラー率が下がるまで、そのルールを継続的に改善すること
- 各セッションの開始時には、そのプロジェクトに関係する lesson を見直すこと

### 4. 完了前に必ず検証する

- 動作を証明できるまでは、タスクを完了扱いにしないこと
- 必要に応じて main ブランチと変更内容を比較すること
- 「これを staff engineer が見て承認するか？」を自問すること
- テスト実行、ログ確認、正しく動くことの提示まで行うこと
- リリース前の確認では、CI だけで十分だとみなさないこと。Windows + Excel の実機 E2E が必要な変更は、`xlflow-tmp-workspace-e2e` skill を使って `tmp_workspaces` で release 前検証を行うこと
- Windows + Excel の実機確認で `push` / `run` / `test` / `pull` / `save` を複数回行う場合は、非 session の単発コマンド連打を避け、`session start -> push --fast --session --no-save -> run/test --session -> save --session -> session stop` を基本形として使うこと。workbook を毎回開き直すと検証が極端に遅くなったり、待機中に詰まって見えることがある

### 5. バランスを保ちながら、よりエレガントな解決を目指す

- 重要な変更の前には、「もっとエレガントなやり方はないか？」と一度立ち止まって考えること
- 修正が場当たり的に感じられる場合は、「今わかっている情報を踏まえて、より洗練された形で実装する」と考え直すこと
- ただし、単純で明白な修正にまでこの手順を持ち込まないこと。過剰設計は避けること
- 成果物を出す前に、自分の実装を自分で疑って見直すこと

### 6. バグ修正は自律的に進める

- バグ報告を受けたら、逐一指示を待たずに自分で調査し、そのまま修正まで進めること
- ログ、エラー、失敗しているテストを使って、自力で原因を特定し解決すること
- ユーザーに不要なコンテキストスイッチを発生させないこと
- 指示がなくても、CI が落ちているなら修正に取り組むこと

---

## 2. 必要な作業手順

コードを生成・変更する前に、作業規模に応じて以下を行うこと。

1. 要件を理解する：関連する仕様書、ADR、既存実装を確認する。
2. 設計を検討する：影響範囲、既存設計との整合性、代替案を確認する。
3. 必要に応じて作業用メモを作成する：
   - 複雑な作業: `tasks/feature_spec.md`
   - 進捗管理: `tasks/todo.md`
   - 再発防止: `tasks/lessons.md`
4. テストを追加・更新する。
5. 実装する。
6. 動作確認する。
7. テストを実行する。
8. 自己レビューする。
9. 必要に応じてドキュメント、ADR、仕様書、CHANGELOG を更新する。

- ADR、仕様書に更新がある場合は次のディレクトリに記録すること
  - ADR: `docs/adr/`
  - 仕様書: `docs/specs/`
- 公開APIに変更がある場合は、以下のドキュメントに記録が必要なことがあります
  - `docs/specs/` 内の仕様書
  - `vitepress/` 内のユーザードキュメント
  - `README.md` 内の概要説明
- ユーザーに影響する変更を行う場合は、`CHANGELOG.md` に追記すること

### リリース前 E2E

- Windows + Excel COM / VBIDE access が関わる変更をリリースする前には、repo-local の `xlflow-tmp-workspace-e2e` skill を使って実機 E2E を実施すること
- 少なくとも blank workbook、standard module round-trip、class module round-trip、UserForm + `.frx` round-trip、`init` の各経路を確認すること
- session-aware workflow を変更した場合は、`session start -> push --fast --session --no-save -> run/test -> save -> session stop` も release gate に含めること
- Windows + Excel の実機 E2E で workbook-backed command を複数回組み合わせる場合は、session-aware workflow を変更していなくても、まず `session start -> push --fast --session --no-save -> run/test --session -> save --session -> session stop` を優先すること
- 検証に使った `tmp_workspaces` の絶対パス、実行コマンド、結果、未検証項目を最終報告へ残すこと

---

## 3. ドキュメント保持ポリシー

### 役割の分離

- `tasks/todo.md` には、セッション単位の進捗管理だけでなく、検証結果、未解決事項、判断理由の要約などを一時的に記録してよい
- `tasks/feature_spec.md` は実装前の作業用仕様書として使ってよいが、将来参照する仕様、制約、検証条件が含まれる場合は使い捨てにしないこと
- `tasks/lessons.md` は再発防止ルールを記録する場所であり、設計判断や仕様そのものを置く場所として使わないこと
- 設計判断やトレードオフは `docs/adr/` に、現行の内部仕様や制約は `docs/specs/` に移すこと

### ADR と仕様書の使い分け

- ADR には、将来の実装者が同じ問題に再び直面したときに役立つ、判断の背景や複数案を比較したうえで採用した方針を記録すること
  - ADR を編集する場合は `adr-manager` skill　を使うこと
- 仕様書には、レビュー、CI、障害対応を通じて確立された恒久的なルールや、CLI、バリデーション、互換性に関わる契約事項を記録すること
- 追加した回帰テストについて、理由を忘れると再発につながるような仕様上の文脈がある場合は、仕様書に記録すること

### 残すべき情報

- 将来の実装者が同じ問題に再び直面したときに役立つ、判断の背景
- 複数案を比較したうえで採用した方針
- レビュー、CI、障害対応を通じて確立された恒久的なルール
- CLI、バリデーション、互換性に関わる契約事項
- 追加した回帰テストについて、理由を忘れると再発につながるような仕様上の文脈

### 捨ててよい情報

- 単発の作業順メモ
- 途中で終わった仮説や中間メモ
- 完了後に参照価値のない進捗ログ
- 判断理由を伴わない単純な手順一覧

---

## 4. コア原則

- **まずはシンプルに**: すべての変更は、可能な限りシンプルに保つこと。影響範囲を最小限にすること
- **手を抜かない**: 根本原因を特定すること。場当たり的な修正は避けること。シニアエンジニア水準を保つこと
- **影響を最小化する**: 必要な部分だけを変更すること。新たなバグを持ち込まないこと

## 5. 注意事項

- xlflowはメインバイナリとdotnetブリッジバイナリの二つで動くため、E2E動作確認を行う際、`go install ./cmd/xlflow` でインストールしても、dotnet bridgeバイナリをインストールすることができない。必ず`task install`でインストールすること

## grepai usage

Use `grepai` for semantic code discovery before broad file reads.

Recommended flow:

1. Use `grepai search "<task intent>"` to find candidate files.
2. Use `grepai trace callers "<symbol>"` or `grepai trace callees "<symbol>"` to identify likely call sites.
3. Treat trace results as candidates, not ground truth.
4. Verify important symbols and call sites with exact search:
   - `rtk rg "<symbol>"`
   - `rtk rg "new <TypeName>"`
   - `rtk rg "<methodName>"`
5. Read only the files confirmed by grepai + exact search.

Branch/index safety:

- Prefer running `grepai watch` in a separate terminal.
- After `git switch`, validate important grepai hits with `rtk rg` before editing.
- If grepai returns files or symbols that do not exist in the current branch, treat the index as stale and restart `grepai watch`.

Use grepai for:

- semantic code discovery
- finding related implementation files
- locating design/spec documents
- finding likely callers/callees

Use `rtk rg` for:

- exact symbol names
- CLI flags
- error messages
- config keys
- test names

Do not rely on grepai trace alone for complete impact analysis.

<!-- headroom:rtk-instructions -->

# RTK (Rust Token Killer) - Token-Optimized Commands

When running shell commands, **always prefix with `rtk`**. This reduces context
usage with zero behavior change. If rtk has no filter for a command, it passes
through unchanged — so it is always safe to use.

This project is developed on **Windows**, so prefer PowerShell-compatible
commands and paths.

## Key Commands

```powershell
# Git
rtk git status
rtk git diff
rtk git log --oneline -20

# Files & Search
rtk dir
rtk dir .\src
rtk read .\path\to\file.txt
rtk rg "pattern"
rtk rg "pattern" .\src
rtk find "pattern"
rtk diff .\path\to\file.txt

# Analysis
rtk err <command>
rtk log .\path\to\log.txt
rtk json .\path\to\file.json
rtk summary <command>
rtk deps
rtk env

# GitHub
rtk gh pr view <number>
rtk gh run list
rtk gh issue list
```

<!-- /headroom:rtk-instructions -->
