# Issue #830: push --fast can skip import for a fresh managed session

## 原因

ExcelPushService.Execute の changed-only スキップ判定が .xlflow/state/push.json
のソース fingerprint 一致のみを見て、workbook attach 前に早期 return する。
push.json は {workbook_path, files[], line_numbers_enabled} のみで、
「どの workbook/session インスタンスに適用したか」の証跡を持たない。
--no-save session push → discard → 新セッションが古いディスクファイルを開く、
という境界をまたぐと誤スキップする。

## 方針 (案A: 適用先バインド — ユーザー承認済み)

push.json を {fingerprint, applied_to} に拡張し、スキップは
「fingerprint 一致」かつ「現在のターゲットが applied_to でカバーされる」場合のみ。

- applied_to.session_id / session_pid / session_hwnd: session-attached push の記録
- applied_to.saved_file: Save() 後の {path, last_write_time_utc_ticks, length}
- session.json に session_id (GUID) を追加。start/attach で新規発行、save 等の再書き込みでは維持
- legacy push.json (applied_to 無し) はスキップ不可 → 一度だけフル import で新形式へ移行

## 実装タスク

- [x] SessionMetadata に SessionId 追加、ReadSessionMetadata/WriteSessionMetadata/MarkSessionPoisoned 更新
- [x] VbaSourceHelper に PushState/PushAppliedTo/PushSavedFile、TryReadPushState、FingerprintEquals、WritePushState を追加
- [x] ExcelPushService: TrySkipUnchangedImport/EvaluatePushStateCoverage (pure 判定) + スキップ応答の session ペイロード実反映 + 成功時の applied_to 記録
- [x] bridge unit tests: 判定マトリクス、session_id 維持、legacy state、Execute レベル (file-target skip / fresh-session non-skip) — 463/463 pass
- [x] docs/specs/cli-contract.md の changed-only 契約更新、CHANGELOG
- [x] 実機 E2E (tmp_workspaces/issue-830-e2e): 2 セッションライフサイクルで macro 実行確認済み
  - session A: push --fast --session --no-save → Extra.bas import、push.json が session A id にバインド
  - session stop --discard → session start (新 id) → push --fast --session --no-save → changed:true (import 実行、誤スキップ無し) → run Extra.Ping --session → ok
  - 同一セッション再 push → changed:false で skip (session.active:true, mode:explicit)
  - save 済み push → 新セッション → saved_file スタンプ一致で skip (source_of_truth: saved_workbook)

## 判定ルール

- metadata が workbook 一致 && !poisoned && attach 解決で記録セッションの Excel が
  workbook を保持 → session target
  - session_id (or pid+hwnd fallback) 一致 → skip
  - saved_file スタンプ一致 → skip
- attach 解決が別 Excel の live workbook に着く → スキップしない
- session target でなく UseSession → スキップしない (attach で session_required)
- それ以外 → file target → saved_file スタンプ一致のみ skip
- poisoned metadata が workbook 一致 → 常にスキップしない (poisoned error を優先)

## レビュー対応 (PR #836 Devin/CodeRabbit)

- [x] BUG1: pid 生存だけでは「session の workbook がまだ開いている」ことを示せない。
      attach 解決 (GetExcelFromSessionMetadata: ROT→hwnd→pid) と同じ経路で Excel を解決し、
      その Excel が workbook を保持し、かつ記録セッション本人 (hwnd 優先/pid fallback) かを検証する
      ResolveSessionWorkbookTarget を追加。RunningExcelHasOpenWorkbook は置き換え。
- [x] BUG2/CodeRabbit: session target スキップ応答の workbook.dirty/needs_save と
      session.save_required/live_newer_than_disk を確定値 false ではなく null (unknown) にする。

## E2E 証跡 (release-gate 記録)

- workspace: C:\Users\HARUMI\orca\workspaces\xlflow\trumpetfish\tmp_workspaces\issue-830-e2e
- 実行コマンド (xlflow は task install 済みの本 worktree ビルド):
  - session A: `xlflow session start --json` → `xlflow push --fast --session --no-save --json`
    (Extra.bas 追加済みソースを import、push.json が session A の session_id にバインド)
  - `xlflow session stop --discard --json` → `xlflow session start --json` (新 session_id)
  - `xlflow push --fast --session --no-save --json` → `changed:true` (import 実行、誤スキップ無し)
  - `xlflow run Extra.Ping --session --json` → ok、B2 = "issue830 ok"
  - 同一セッション再 `push --fast --session --no-save --json` → `changed:false` skip
    (session.active:true, mode:explicit, dirty:null)
  - save あり push → `session stop`/`session start` → `push --fast` → saved_file スタンプ一致で skip
    (source_of_truth: saved_workbook)
- 未検証項目: poisoned session の実機再現、複数 Excel インスタンスでの同一 workbook 同時オープン
