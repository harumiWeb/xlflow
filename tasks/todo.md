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

- metadata が workbook 一致 && !poisoned && pid 生存 → session target
  - session_id (or pid+hwnd fallback) 一致 → skip
  - saved_file スタンプ一致 → skip
- session target でなく UseSession → スキップしない (attach で session_required)
- それ以外 → file target → saved_file スタンプ一致のみ skip
- poisoned metadata が workbook 一致 → 常にスキップしない (poisoned error を優先)
