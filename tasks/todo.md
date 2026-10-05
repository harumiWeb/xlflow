# Issue #917 final review follow-up

- [x] Preserve unchanged axes during move/snap and reject gestures that leave existing overflow; add geometry regressions.
- [x] Enforce a single move/resize operation or distinct move+resize pair for one control at the RPC boundary; test rejected raw RPC batches and accepted pairs.
- [x] Run affected Go and VS Code checks, preserve Pass 1 evidence and clean up review resources.
- [x] Commit the fixes and run independent Pass 2 against the full Issue #917 change; both Pass 1 findings resolved.
- [x] Reject moves of existing sub-point controls; verify explicit resize repair and real Webview no-edit behavior.
- [x] Preserve Pass 2 report and clean up its resources; stop the full review loop at two passes and verify the isolated P2 correction through regression tests/self-review.

# PR #925 review follow-up

- [x] 操作エラーの中間状態検証による上書きを防ぐ。
- [x] ID 重複診断を canonical validator に委譲し、追加対象 ID の文脈を保持する。
- [x] 複数行引用 scalar のコメント二重化と indentless sequence の最後の要素削除を修正する。
- [x] reorder は整数 gap を利用し、必要な最小連続範囲だけを変更する。
- [x] 回帰テスト・vet・docs チェックを確認する。

# Issue #916 final review follow-up

- [x] 明示タグ付き引用 scalar の境界検出を修正する。
- [x] 引用内の #、エスケープ、flow 区切り、LF/CRLF のコメント保持を回帰テストする。
- [x] 関連 Go テストと vet を確認し、Pass 1 の報告と再現コードを保持してレビュー環境を cleanup する。
- [ ] 修正をコミットして独立レビュー Pass 2 を実行する。

# Issue #915 designer review follow-up

- [x] PR #924: prefer client/inside dimensions on both axes; render explicit ComboBox values including empty strings; localize Webview UI through host l10n messages.
- [x] Preserve zIndex-ordered MultiPage selection to match pure-Go generation/topology; add reordered Page regression and explain unsupported source-order recommendation.
- [x] Run renderer and real Webview tests, lint/docs checks; publish with regression evidence. Docstring coverage is advisory, not a repository gate.

- [x] Honor explicit visibility for controls and Page tabs, retaining original Page collection indexes.
- [x] Add hidden-parent and hidden-Page selection regressions and rerun Designer/Windows Webview checks.

# PR #923 review follow-up

- [x] Resolve picture roles before every canonical VBA component collector; cover module, class and document roots through CLI pack/file push.
- [x] Preserve case-distinct Code/ loose modules and fingerprints on case-sensitive hosts.
- [x] Correct Windows CI assertions to resolve expected physical paths before host-key comparison, retaining symlink target checks. Directory-alias fixtures cover resolved roots as well as filename casing; the initial case-only correction was insufficient on the runner.
- [x] Remove the CodeQL allocation-size sum, propagate input path-resolution errors, and correct obsolete pack restrictions.
- [x] Run affected Windows/Linux checks, lint and installed Excel gate; publish fixes and report review dispositions. CodeQL alert 11 is fixed and Linux CI passes; Windows CI is rerun after the physical-path assertion correction.

# PR #913 review follow-up

The persisted-caption finding is valid: caption projection respects the
control contract, but unsupported-property reporting still treats every
Caption as projected and hides nonempty unmodeled values.

- [x] Add full-projection regressions for nonempty unsupported captions,
      empty defaults, and supported captions; confirm the regression fails.
- [x] Align unsupported Caption reporting with the canonical control contract
      while keeping empty unsupported defaults silent.
- [x] Verify affected Go tests, lint/docs and staged formatting.

Publish the verified fix to the existing PR and reply with the regression
evidence. Test-function docstrings are an advisory suggestion rather than a
repository requirement and are outside this behavior fix.

# PR #907 review follow-up

All five comments are valid presentation/projection defects; retain global
resolution, native metrics identities and incident-edge selection.

1. [x] Filter displayed root candidates while retaining unresolved requests.
2. [x] Retain all related SCC member dependency nodes and boundary IDs.
3. [x] Join conditional procedure metrics by declaration start position.
4. [x] Allow existing empty directories as display scopes; reject invalid files.
5. [x] Distinguish the field inventory from mutable access counts in human output.
6. [x] Add regressions and verify architecture/output/CLI tests, affected race,
       lint/docs/staged formatting and whitespace checks.

# Issue #463 final-review Pass 1 follow-up

Validated against `d143e7e8`: implicit Application callbacks are extracted
from unresolved IR and can lose target evidence and possible reachability.

1. [x] Project callback references from the existing resolved IR snapshot.
2. [x] Cover receiverless and With Application Run/OnTime/OnKey, including
       project-procedure and non-callable local shadow controls.
3. [x] Run focused architecture/calls/reachability/CLI tests, race and lint.
4. [x] Pass 2 reviewed `4a0fda3c`: the P2 is resolved, the focused callback
       regression passes independently, and no confirmed findings remain.

# PR #906 review follow-up

CodeQL #10 identifies unchecked addition in a capacity hint; use the existing
raw length directly. The ADR's graduation boundary is historical and must be
identified as superseded by the later UserForm amendments. The repeated-root-
Caption report is not reproduced: applyPersistenceEdit updates the clone's
Caption before the next iteration, confirmed by the exact reported sequence.

1. [x] Remove the allocation-size sum and verify caption byte preservation.
2. [x] Cover repeated root Caption edits and classify the review with evidence:
       Original -> Changed -> Original passes without changing ApplyEdits.
3. [x] Mark historical ADR authority consistently with current amendments.
4. [x] Run full/focused tests, affected race/vet, docs and formatting checks.

Publish the verified changes to PR #906 and reply with each classification.

# Issue #887 final-review Pass 1 follow-up

Validated against `c87cac40`: the cleanup PID race, workspace reparse-point
boundary and unstable inventory errors are valid and in scope. Duplicate
MSForms-reference handling is unchanged from the parent commit; its VBE
acceptance remains unverified and is a separate follow-up, not a review fix.

1. [x] Pin the verified process handle through cleanup; never kill by PID.
2. [x] Reject existing reparse-point ancestors before creating a workspace.
3. [x] Sort inventory validation keys and add multi-error regressions.
4. [x] Run focused/full verification and the real Excel gate.
5. [x] Commit fixes and complete the bounded Pass 2 review.

Pass 2 reviewed `8af2ff79` with gpt-6-luna xhigh: PASS, all three in-scope
findings resolved, no new confirmed findings. Independent harness and
inventory regressions passed; retained Excel evidence was read back.

# Issue #887

- [x] Add explicit template/source UserForm topology configuration (default template).
- [x] Discover canonical template specs alongside legacy code-only forms; retain explicit property intent.
- [x] Plan supported Designer updates/additions/removals atomically with lossless preservation.
- [x] Reconcile project metadata and reparse output before publication.
- [x] Cover Go/CLI regressions and run focused/full/race/vet/lint checks.
- [x] Verify template and blank artifacts in real Excel, including save/reopen and final-form removal.
- [x] Update ADR/spec/config docs/CHANGELOG and review the complete diff.

Local commands, retained workspaces and results:
`internal/pack/testdata/userform-integration/README.md`.

# PR #904 review follow-up

REGISTERED aggregate-size and malformed Forms LIBID findings are valid.
PROJECT has an aggregate Size (the comment's first-length explanation is
incorrect), but it too must be ignored on read in favor of two bounded paths.

1. [x] Share bounded reference payload sizing between both parser paths.
2. [x] Validate LIBID grammar and cover malformed Forms admission atomically.
3. [x] Cover PROJECT and REGISTERED aggregate-size variations, truncations
       and reference byte preservation; remove the unused E2E workbook table.
4. [x] Affected pack/sourceinventory/CLI tests, race, vet, parser fuzz,
       lint/docs/PowerShell and staged formatting checks pass.

Publish the verified fix to the existing PR and reply with these results.

# Issue #886 final-review Pass 1 follow-up

Validated against `852a57b7`: the P2 NUL validation finding is in scope.

1. [x] Validate every decoded reference string, including ORIGINAL, CONTROL
       twiddled LIBID and extended name, before accepting a reference group.
2. [x] Add corrupted-field regressions and atomic Forms admission coverage;
       retain valid Excel reference byte-preservation controls.
3. [x] Run pack tests, race and vet; commit hooks check lint and formatting.
4. [x] Pass 2 reviewed `eb21a91f` with gpt-6-luna xhigh: the P2 is resolved,
       no confirmed in-scope defects remain, and affected package tests pass.

# Issue #883 final-review Pass 1 follow-up

Validated against `76cae6b8`: both P2 findings are in scope.

1. [x] Require canonical MSForms LIBID evidence in raw reference records;
       reject an unrelated twiddled GUID even if its display name is MSForms.
2. [x] Reject equal VB_Base GUIDs at the project-addition boundary, including
       case variants, without mutating the input project.
3. [x] Add regressions and run focused userforms/pack/excel tests, vet and docs
       checks; commit hooks run lint and staged formatting checks.
4. [x] Obtain Pass 2 with gpt-6-luna xhigh: `ace0f284` passed with both P2 fixes
       verified and no new findings. Preserve #886/#887 scope. The admission
       checks do not change generated Designer bytes;
       committed Excel normalization evidence still passes in ordinary tests.

# PR #902 review follow-up (Issue #882)

Verified against `40267db2`: both Devin findings are valid. The two CodeQL
alerts refer to the same unchecked sum used for key-slice capacity.

1. [x] Reject invalid before/after coordinate systems before interpreting
       geometry; retain omitted, points and parent-relative inputs.
2. [x] Check the supplied legacy caption baseline when build.caption first
       appears, retaining changed-build precedence and atomic failure.
3. [x] Collect the property-key union without adding untrusted map lengths;
       cover overlap, disjoint keys and deterministic edits.
4. [x] Add focused regressions, update compiler/trusted-Excel contracts, and
       run focused/full tests, race/lint/docs and the saved/reopened Excel gate.

Publish the verified fix to the existing PR and reply to each review thread
with its disposition and evidence; keep remote CI status separate.

CodeRabbit's pack integration warning belongs to explicit follow-up #887;
private-helper docstring coverage is not a repository gate. The developer
Excel harness intentionally executes a sentinel, so explicitly document that
only trusted generated workbooks and known derivatives may be verified.

# Issue #882 final-review follow-up

Independent review of `d63cde3a` confirmed one in-scope P3: seven new
snake-case test labels are incorrectly extracted as structured error codes.
Keep this fix local to the new test labels rather than changing the existing
inventory generator's repository-wide extraction contract.

1. [x] Rename the seven test labels and regenerate the reference inventory.
2. [x] Run focused userforms/pack tests and docs/format/lint checks; confirm
       only the four real `userform_edit_*` codes remain newly inventoried.

Publication follows verification: commit the review fix, push the feature
branch, create the PR, and validate its GitHub read-back.

# PR #901 review follow-up (Issue #881)

Verified against commit `7396d1fe`; all three Devin bug reports are valid and
in scope. The temporary-allocation observation is a performance hardening note,
not a demonstrated regression, and remains a separate follow-up.

1. [x] Resolve known Designer stream names with CFB case-insensitive identity,
       retain their original spelling, and replay every captured stream under
       that spelling. Add case-varied root and optional-stream round-trip tests.
2. [x] Let `pull --backend auto` inspect module topology without parsing
       Designer data when `code_source = "frm"`, so unsupported Designer
       layouts can still select Excel; keep strict parsing for the file backend
       and sidecar probe. Add malformed-Designer probe regressions.
3. [x] Count parsed Designer streams in `pack.modules.carried_streams` and add
       a form-bearing output-contract regression.
4. [x] Run focused and full validation, commit and push the fixes, then reply
       to all review threads with evidence and the allocation disposition.

# PR #900 review follow-up

Verified against commit `c4ed68d8`; the stale compatibility artifact portion
of the Devin finding is valid and in scope. Direct FormSpec-to-MS-OFORMS push
remains assigned to parent Issue #876.

1. [x] Mark specs emitted by `pull --backend file` as having unsynchronized
       compatibility `.frm` / `.frx` artifacts.
2. [x] Make the shared Excel/file push preflight reject that persistent marker
       even when a same-named stale `.frm` exists; keep the artifacts
       non-destructive and unmanaged.
3. [x] Add focused fresh/stale artifact regressions, update ADR/spec/docs and
       the real-Excel gate, and run validation. The review reply must preserve
       the Issue #876 scope boundary.

# PR #893 review follow-up 2

Verified against commit `4c57c077`; the stale delegated warning finding is
valid and in scope:

1. [x] Explicitly clear `XLFLOW_PULL_AUTO_PROBE_WARNING` for warning-free WSL
       auto-pull delegation, even when the parent environment and `WSLENV`
       contain a stale value.
2. [x] Add focused delegation regression coverage, run repository validation,
       push the fix, and reply to the review thread with evidence.

# PR #893 review follow-up

Verified against commit `bb327504`; three inline findings are valid and in
scope, while WSL detection of arbitrary externally opened Excel workbooks is
the approved local-auto boundary rather than a defect:

1. [x] Skip workbook coordination before `auto` selection, then acquire the
       workbook lease only when auto actually selects Excel. Keep file pulls
       under source-tree coordination and preserve explicit backend behavior.
2. [x] On an indeterminate Windows open-state probe, require attachment to an
       already-open matching workbook instead of opening a separate saved copy.
3. [x] Preserve the WSL session-probe failure warning in the delegated Windows
       JSON result, with focused end-to-end delegation coverage.
4. [x] Investigate the VS Code CI failure independently, run focused/full
       validation, push the review fix, and reply to every inline thread with
       evidence or the documented design disposition.

# Issue #892 final-review pass 1 follow-up

Independent review of commit `4e08cf77` confirmed one in-scope P2:

1. [x] Restrict the auto backend capability probe to saved-workbook
       properties (parse/protection/UserForm support) so it never reads managed
       source trees before source-tree leases and `--wait` take effect.
2. [x] Keep source-path, stale-file, and publication planning validation inside
       `PullContext` while its leases are held; add focused regression coverage.
3. [x] Run focused/full tests and lint, commit the fix, and request final-review
       pass 2 on the exact fix commit.

# Issue #877 final-review pass 1 follow-up

Independent review of commit `d20ca701` confirmed one in-scope P1:

1. [x] Key every CFB sibling tree by `DirectoryNameKey` so case-variant
       storage definitions deduplicate only when metadata matches, conflicting
       metadata and storage/stream collisions fail deterministically, and the
       first spelling remains the serialized display name.
2. [x] Reject case-insensitive storage/storage and storage/stream duplicates
       while reading malformed CFB directory trees.
3. [x] Add writer, reader, determinism, and malformed-input regressions; run
       focused/full tests, lint, format checks, and the real-Excel pack gate.

# PR #874 review follow-up (Issue #871)

Verified against commit `f5bef996`; the five inline findings and one
documentation finding are valid and in scope. The generic docstring coverage
warning does not identify a repository contract violation and needs no change.

1. [x] Preserve case-only module rename semantics on case-sensitive filesystems
       without deleting an aliased target on case-insensitive filesystems.
2. [x] Classify wrapped publication failures as `pull_source_publish_failed`
       before generic workbook-path errors.
3. [x] Reject overlap in either direction between the forms root and each
       managed source root before reconciliation.
4. [x] Keep `target.kind` within the stable `file` vocabulary and make pull
       prerequisites backend-specific.
5. [x] Add focused regressions, run affected tests and lint/docs checks, commit,
       push, and reply to each review thread with evidence.

# PR #864 review follow-up (Issue #855)

Verified against commit `b8d8a2e5`; both Devin comments are valid and in
scope:

1. Preserve the existing `pack` contract that configured absolute source
   roots may live outside the project while keeping the containment contract
   of ordinary encoding commands unchanged. Add a pack-only sourceencoding
   option, retain managed-root symlink containment, and report external
   diagnostic paths as absolute paths.
2. Normalize configured source-root separators before discovery so Unix
   preflight and `sourceinventory` consume the same paths instead of silently
   skipping backslash-configured roots.
3. Add sourceencoding and pack CLI regressions for valid/invalid external
   roots and invalid source under a backslash-configured root. Run focused and
   full tests, lint/docs/format checks, push the follow-up commit, and reply to
   both review threads.

# PR #863 review follow-up (Issue #854)

Verified against commit `0e6c6145`; the six Devin observations and two
CodeRabbit comments reduce to six valid in-scope fixes:

1. Publish through the canonical output path so an existing output symlink is
   preserved and its referent is replaced; retain both lexical and canonical
   candidates for Office lock-file detection. Extend identity comparison with
   `os.SameFile` for existing hard-link aliases.
2. Create staging files exclusively with mode `0644` so Unix umask applies to
   new outputs, and preserve an existing destination's permission bits during
   replacement.
3. Keep the exclusive replacement probe on Windows, where sharing modes are
   meaningful, but do not require write access to the old file on Unix; final
   rename errors remain the source of truth there.
4. Use Linux `renameat2(RENAME_NOREPLACE)` for atomic create and retain hard
   links only as the ENOSYS/EINVAL fallback and for other non-Windows systems.
   Never classify create-time EPERM/EACCES as a busy destination.
5. Surface cleanup failure together with pre-publication errors while keeping
   `errors.Is` classification intact.
6. Read the staged `xl/vbaProject.bin` entry fully so ZIP CRC/data corruption
   is detected before publication.

Add focused Windows/Unix/CLI regressions, run affected tests on Windows and
WSL Linux, run lint/docs checks, commit, push to PR #863, and reply to each
review thread with the disposition.

# Final review pass 3 follow-ups (validate-static-rules-default, aa889dfc)

Independent whole-branch review verdict: approve-with-notes. Remaining items:

1. DONE (PR #860 review round 4): review_test.go now bounds Unreviewed by
   the committed backlog (<= 9231). Review-driven decreases still pass;
   new snapshot findings fail until triaged or the ceiling is raised.
2. Follow-up: addressOfEntryNames is a name-only project-wide set; a
   same-named non-callback procedure can be over-suppressed (bounded
   false negative, consistent with the suppression-only fail-open
   contract). Revisit only if a real misreport is observed.

# Final review pass 1 fixes (validate-static-rules-default, cc016ad9)

Independent review verdict: approve-with-notes. Verified findings:

1. P2 cross-module AddressOf signature constraint: VBA resolves a bare
   AddressOf target project-wide across public procedures in standard
   modules, but parameterPassingSignatureConstrained only consulted the
   candidate file's own facts.addressOfNames. A callback declared in a
   dedicated module still received VBA273/VBA275/VBA277 advice - the same FP
   class this branch fixed in-module. VBA265 is unaffected (its candidates
   are Private-only) and keeps its current merged dynamic-entry contract.
   Fix: add analysisContext.addressOfEntryNames (project-wide union of
   facts.addressOfNames) populated when any signature-constrained
   parameter-passing rule is enabled (VBA273/275/276/277), thread it into
   parameterPassingSignatureConstrained alongside the file-local check, and
   widen it in the realtime path from projectDocuments the same way
   dynamicEntryNames is widened. Add a two-module regression test.
2. P3 metrics assertion: review_test.go now uses floors + dropped
   Unreviewed==0. Intentional (opt-in profile emits unreviewed observations
   by design); add a TODO documenting when the bounds can be tightened.
3. P3 missing config row: vitepress/reference/config-file.md [analyze]
   table lacks detect_invalid_ismissing_usage (pre-existing gap, now
   user-facing because VBA283 is default-on). Add the row + narrative line.
4. P3 stale comments/doc: analyzer.go dynamicEntryNames comment (mentions
   VBA260, actual consumers are VBA265 + now includes AddressOf names),
   module_facts.go "three source scans" comment (now four), and
   vba-parameter-passing-diagnostics.md VBA276 exclusion list (code applies
   the AddressOf constraint too, though unreachable for Property Let/Set -
   document the check or the dead branch).
5. Unsupported/follow-up observations recorded by the reviewer (realtime
   private-name-collision parity, Application.Run string-literal names not
   consulted by parameter passing) are pre-existing design boundaries; the
   realtime addressOfEntryNames widening shares the VBA265 scan loop so
   parity bookkeeping stays symmetric. No separate fix.

# Issue #826: VBA class/interface public-API hazard diagnostics (VBA278-VBA282)

Parent issue #816. Five opt-in warning rules, all default-disabled,
warning-level, inline-suppressible, non-blocking, on batch + realtime + LSP.
Renumbered from the originally implemented VBA273-VBA277 because origin/main
PR #845 (issue #823) already owns those IDs for parameter-passing rules.

- VBA278 public member underscore names (class/form/document; skips
  control-verified/intrinsic event handlers, verified Implements bindings,
  unresolved-Implements prefixes, and Class_Initialize/Class_Terminate).
- VBA279 non-Private Enum in document modules.
- VBA280 write-only Property (Let/Set without Get), one finding per name;
  Friend counts as public-facing; conditional/recovered accessors mark the
  name uncertain and fail open.
- VBA281 public members whose <Interface>\_<Member> binding resolves to a
  declared public member of the implemented interface class, and explicitly
  Public recognized event handlers.
- VBA282 self-name access inside predeclared modules (document/form/
  VB_PredeclaredId class); local/parameter/module-scope shadows and TypeOf
  type operands excluded; recovered/conditional procedures skipped.

PR #847 review fixes applied (Devin 1-6 + CodeRabbit 1-4; all verified valid):

- Form helper misclassified as event: recognizedEventMember validates the
  control prefix against designer controls when known and falls back to
  intrinsic UserForm\_\* names when unknown. Lazily resolves controls because
  VBA220's gate only covers \_change/\_click names.
- IFace_Utility FP: interfaceBindingMatch requires the suffix to name a
  public member of the resolved, unambiguous interface class; resolved but
  absent suffixes fall back to VBA278, unresolved interfaces fail open.
- VBA280 uncertain accessors: Recovered/ConditionalBranches mark the
  property name uncertain; Friend writers now count as public-facing.
- VBA282 guards: proc.IR.Symbol.Recovered/ConditionalBranches fail open;
  ScopeLocal/ScopeParameter/ScopeModule accesses excluded; TypeOf exclusion
  narrowed to the last (type-operand) child of type_of_expression.
- Spec fail-open text updated to match implementation.

# Conflict resolution (merge origin/main @ aabb6575)

- Rule ID collision: our VBA273-277 -> VBA278-282; config keys unchanged.
- registry.json: main entries kept, ours appended renumbered.
- diagnostics.md regenerated from registry; oracle fixture dirs + ids and
  fixture_contract_test.go renumbered; workspace_test.go keys renumbered.
- Corpus snapshots to be regenerated (corpus:update-snapshots) and ledger
  rows renumbered to VBA278-282; metrics numbers rechecked after regen.

Remaining: corpus regen + metrics, format:check/docs:check, full test pass,
commit merge, push, update PR #847 body (new rule IDs + oracle case IDs
vba278-public-underscore-member, vba279-document-public-enum,
vba280-write-only-property, vba281-public-event-handler,
vba281-public-interface-member-collision, vba282-predeclared-self-access).

# PR #845 review follow-up (Issue #823, VBA273-VBA277)

Review comments on internal/analyze/parameter_passing.go. Verified against
code; both bug reports are valid and in scope. Plan:

1. [Devin + CodeRabbit] Property accessor collision: mutation records are
   keyed by qualified name only, so Property Get Item and Property Let Item
   share one summary (first registered wins). VBA275 can falsely claim the
   Let index parameter is never written; VBA274 can miss a written Let value
   parameter. Fix: kind-qualify record keys (kind|name) for summary storage
   and finding lookup; call-edge resolution uses candidate kind when present
   and falls back to unique-name match or possiblyWritten when ambiguous.
   Regression test: Property Get before Property Let, Let writes index and
   assigns value parameter.
2. [Devin] applyArgumentMutation checks parameterTypeIsObject before
   parameterTypeIsUDT, so a member argument on a UDT that shadows a builtin
   object name (Type Collection) drops the propagated write and VBA275
   falsely fires. Same latent issue family as the recordParameterMemberWrite
   fix. Fix: UDT check first, matching VBA name-resolution precedence.
   Regression test: module UDT named Collection, member forwarded to a
   writing ByRef call.

Advisory (Devin): no executed VBE oracle case IDs. Oracle v1 supports
compile probes only; the modeled semantics (property value ByVal, forced
ByVal parens) are covered by byref-parenthesized-variable (accepted) and
property-signature-valid/invalid at compile level; runtime cases are out of
scope for the current probe contract. Rules remain opt-in and fail open.

# Issue #879 final-review follow-up

- [x] Keep TabStrip `ListIndex` out of unsupported FormSpec fields and report it as unsupported; cover Project -> WriteSnapshot -> LoadFormSpec.
- [x] Report non-empty ControlSource and RowSource site state as unsupported, and audit all decoded site strings so meaningful unprojected state cannot disappear silently; add focused warning regressions.
- [x] Run focused userform projection/spec tests and the relevant package validation; review the complete diff.

# PR #897 review follow-up

- [x] Report non-default unsupported nested `Level` state on its owning container, and report unmodeled non-default site/control bitfield state without flagging file-format defaults; cover root and nested fixtures.
- [x] Surface non-default form `BooleanProperties` through the existing unsupported warning contract and add a regression.
- [x] Correct the YAML parentId completion test cursor offset; confirm Page/MultiPage remain projection/snapshot types rather than additions to the built-in authoring contract, and production CLI wiring belongs to parent issue #876.
- [x] Add and run same-workbook Excel snapshot versus pure-Go projection parity verification; run focused and affected package tests. Commit/push fixes and reply to review threads with evidence.

# PR #837 review follow-up (Issue #822, VBA257/VBA258)

Review comments on internal/analyze/discarded_return.go. Verified against IR
dump and resolver source; all six are valid. Plan:

1. [Devin] Parse errors permit false all-discard claims: VBA258 claims every
   caller discards, but a parse error can hide a call or a dynamic-dispatch
   reference (Application.Run reaches Private procedures too). Fix: bail out
   of the whole pass when any file.IR.Parse has HasError/HasMissing.
2. [CodeRabbit] PathFilter removes modules before parsedFiles; hidden modules
   can contain callers or dynamic references. Note: resolutionComplete is too
   broad (AnalyzeProject always sets typeDBResolutionIncomplete), so gate on
   PathFilter==nil instead. Combined with #1 this also covers the Friend case.
3. [Devin] Same-name arguments hidden by isCalleePart: 'Foo Foo' argument is
   skipped because it shares the callee name inside the call range. Fix:
   locate the exact callee token range (bytes.Index of Callee.Text inside
   call.Range) and exempt only expressions inside that range.
4. [CodeRabbit] Implicit Application dispatch: receiverless 'Run "x"' and
   With-block '.Run' produce no dynamic references. Fix: in calls package,
   map callee member run/ontime/onkey with nil/empty/"." receiver to
   "application.<member>" when resolution ran and did not bind a project
   procedure (keeps inspect path unchanged via ResolutionNotAttempted).
5. [Devin] Underscore exclusion too broad: Implements modules exclude every
   underscored proc. Fix: collect implemented interface names from
   TypeReferences (kind=="implements") and exclude only <Interface>\_<Member>.
6. [CodeRabbit nitpick] Implements exclusion test cannot fail (unknownDynamic
   suppresses everything and IWorker_DoWork is never called). Fix: add an
   independent test with a standalone IWorker_DoWork call and no dynamic
   dispatch.

Non-actionable: CodeRabbit docstring-coverage warning (bot threshold, repo
convention does not doc-comment every function).

# PR #837 review follow-up 2 (post-merge)

Second-round comments verified against merged HEAD (bfa7e23a). Devin's three
comments and CodeRabbit's two actionable comments are stale: they target the
pre-fix head and are already covered by 92cfacb3 (parse-error bail-out,
calleeTokenRange, implementedInterfaceNames, projectViewComplete gate,
implicitApplicationAPIName) - CodeRabbit marks both as addressed.

One valid finding remains:

1. [CodeRabbit] Interface names containing underscores: `Implements I_Foo`
   makes the implementation `I_Foo_Bar`, but strings.Cut splits at the first
   underscore so qualifier `i` never matches `i_foo`. Fix: prefix-match the
   lowered symbol name against each complete interface name + `_`. Add a
   regression test with an underscored interface name.

# PR #838 review follow-up (Issue #827, VBA260/VBA261/VBA262)

Verified against commit `2ebcf3f6`; actionable plan:

1. [Devin] Exclude bracket-escaped identifiers that bind to procedure, module,
   or project declarations from VBA262. Add declared enum/local/module and
   undeclared host-expression regressions.
2. [Devin] Require VBA260's `ThisWorkbook` token to be an unqualified,
   unshadowed project-global receiver. Reject member-qualified and locally
   shadowed forms with focused tests.
3. [Devin/CodeRabbit] Preserve logical-to-physical source positions for VBA260
   continuation statements so ranges and line suppressions bind to the actual
   physical line. Cover indentation and a match beginning on a continued line.
4. [CodeRabbit] Distinguish partial worksheet catalogs from unavailable
   catalogs in structured warning text while retaining valid mappings.
5. [CodeRabbit] Skip `sheetData` subtrees through `xml.Decoder.Skip` to reduce
   token handling while still validating XML through EOF and rejecting later
   conflicting `sheetPr` metadata.
6. [CodeRabbit acceptance check] Extend VBA261 diagnostics and tests to explain
   the error-behavior difference as well as binding and return-type semantics.
7. Run focused tests, full affected packages, corpus snapshots, lint/docs, and
   push the review-fix commit. Re-read PR comments and remote checks.

Non-actionable/follow-up: repository style does not require doc comments on all
private helpers. The security job reports GO-2026-6452 in the pre-existing
excelize v2.11.0 dependency with no fixed version, so it is tracked separately
from this review fix. VBE-oracle need will be reassessed after checking existing
bracket fixtures; VBA262 is not compile-equivalent and will not be promoted as
compile evidence.

# Issue #823 final-review pass 1 follow-up

Review run_7c453f1dc6a5 on commit 202af7ff. All four confirmed findings
verified against IR probes and real Excel runtime checks; all valid. Plan:

1. F1: With-block implicit member writes invisible to the mutation summary.
   Fix: build a shared withReceiver map (statement ID to innermost With
   receiver expression), resolve implicit member targets and implicit member
   call arguments through it, and fail open on the unresolved spaced-callee
   form (Replace .Left parses into the callee expression).
2. F2: Erase/Input #/Line Input #/Get # operands surface as unknown
   statements with read accesses. Fix: record write targets for those
   syntax kinds (file_number_literal excluded; Get uses TargetID only).
3. F3: VBA273 fires on ParamArray with an impossible suggestion. Fix:
   exclude ParamArray from VBA273 and VBA277.
4. F4: cfg/clone.go cloneParameters missed PassingRange. Fix: clone it and
   extend the snapshot-isolation test.
5. Follow-ups: VBA276 gains the constrained exclusion (VBE rejects
   Implements passing-modifier mismatches in both directions; verified on
   real Excel). Indexed element writes (v(0) = 1) become possiblyWritten:
   not a VBA274 reassignment, still caller-visible so VBA275 stays
   suppressed. Call Mutate((x)) is forced ByVal and no longer propagates.
   Missing mutation summary now returns nil explicitly (fail-open).

# Issue #823 final-review pass 2 follow-up

Review run_7c453f1dc6a5 pass 2 on commit 378b1333. Three confirmed findings
verified against the report's IR dumps, probe outputs, and the implicated
source; all valid. Plan:

1. F-P2-1: named-argument ExpressionIDs bind the widest expression in the
   whole `a:=x` range, so a name at least as long as its value hides the
   value from propagation (VBA275 false positive) and from every consumer
   that assumes ExpressionIDs[i] == Named[].ExpressionID
   (object_use_before_set, array_safety_source_order, object_container_flow,
   variable_assignment). Fix in procedureir/visitor.go: use
   argument.valueRange when the argument is named; add a width-inverted
   regression test for parameter passing and named-argument coverage for the
   sibling consumers.
2. F-P2-2: indexed call expressions (arr(0)) emit no receiver access, so
   element-valued arguments, With receivers, and file-statement operands are
   invisible to propagation (VBA275 false positives). Fix: resolve
   ExpressionCall roots through the callee child in parameterExprRootName,
   generalize the implicit-member argument fallback to a root-resolution
   fallback for any unhandled argument, and mark ExpressionCall write
   targets possiblyWritten in recordParameterWriteTarget.
3. F-P2-3: parameterMemberTargetRoot stops at spaces so [My Field] loses
   member writes, and parameterTypeIsObject runs before the UDT check so a
   UDT shadowing a builtin object name (e.g. Collection) drops member
   writes. Fix: extract bracketed names whole, fall back to expression
   resolution whenever the text root matches no parameter, and check UDT
   membership before the builtin object list.
4. Regression tests for each, plus docs/spec updates (named mapping is now
   width-independent; element-valued arguments and receivers covered).

# PR #849 review and security follow-up (Issue #828)

- [x] Update tree-sitter-vba v0.14.5 in the third-party licence inventory; verify the exact security gate locally.
- [x] Fix conditional-procedure scope for VB087. VB074 checks every parsed branch body; VB090 fails open for branch-dependent labels.
- [x] Make VB088 inspect only declaration visibility and cache module kind once per scan.
- [x] Cover numeric labels and their branch references for VB090.
- [x] Report VB081 on same-source declared suffixed callees while leaving unknown and intrinsic callees unreported.
- [x] Confirm focused and full Go suites, corpus snapshots, lint/docs, and security inventory locally; push a review-fix commit to PR #849.

# Issue #852 final-review pass 2 follow-up

- [x] Make template-owned PROJECT topology validation deterministic by checking removals/renames in original declaration order and additions in requested spec order.
- [x] Add regressions that assert the first reported conflict for multiple invalid components.
- [x] Run focused pack tests, formatting checks, and commit the review fix.

# PR #861 review follow-up

- [x] Validate `Attribute VB_Name` for existing standard and class component updates before assigning normalized source.
- [x] Add regression coverage for mismatched existing standard and class modules while preserving document/UserForm header behavior.
- [x] Run focused and full validation, then push the fix and reply to the review thread.

# PR #903 review follow-up

- [x] Require REFERENCECONTROL OriginalTypeLib identity rather than trusting a twiddled LIBID; cover conflicting identities and atomic rejection.
- [x] Reject incomplete generated-form expected maps before Excel starts; compare SpinButton Delay and ScrollBar Delay/ProportionalThumb with regression coverage.
- [x] Align .NET builtin control contracts with the canonical registry and fix the failing Windows/Linux CI test.
- [x] Verify focused Go/.NET/PowerShell tests, lint/docs, and staged formatting.
- [x] Rerun the fresh Excel generation gate after the concurrent session finishes: Delay/ScrollBar ProportionalThumb, both forms, save/reopen, sentinel, and normalized binary readback passed; owned PID 204972 cleanup confirmed. Evidence and command recorded in the fixture README.
- [x] Push the fixes (`4e735e72`) and reply to the three review comments with focused regression and fresh Excel evidence; verify the updated PR body on GitHub.

# Issue #886

- [x] Parse complete references and ensure first-UserForm Forms reference atomically.
- [x] Enable canonical spec/code discovery and common-control generation in blank pack.
- [x] Cover reference preservation, code modes, capability failures and publication safety.
- [x] Validate generated artifacts in Excel, retain fixtures and run release gates.
- [x] Update contracts/docs, run checks and independent review. Reference reviewer found no confirmed defects and independently passed pack/fuzz tests; blank-input reviewer found no confirmed defects. Case-variant directory spelling remains an unconfirmed follow-up, not a changed contract. Evidence: `internal/pack/vbaproject/testdata/forms-reference-excel/README.md`.

# PR #908 review follow-up

- [x] Clarify case-sensitive retained names and case-insensitive type labels.
- [x] Check the edited pack output's module identity rather than the original output.
- [x] Reproduce default TabIndex collisions during topology edits; reserve final sibling retained/explicit values before assigning generated defaults, including property-bag aliases, and preserve existing values.
- [x] Run focused regressions, affected package tests, CLI pack tests, and required lint/docs checks.

# Issue #885 MultiPage / Page / TabStrip

- [x] PR #910 review follow-up: skip observed Page geometry; retain nested omitted tabs using the inspection type resolver; defer and verify sibling TabIndex; populate lists before selection/text; prevent caption normalization from adding exported source lines. Bridge regressions pass (507 full / 59 focused); session-backed Excel build/apply/snapshot/save and reopened snapshot succeed in `tmp_workspaces/issue-910-review-e2e`.
- [x] PR #910 final-selection acceptance: the pack gate exposed copied observed selection from the old Page collection. Discard unchanged selection observations before final-topology validation and retain original observations in the property-only stage. Six template/edit regressions and related Go tests/lint pass. Pure-Go pack, Excel runtime, save/reopen runtime and confirmed cleanup pass in `tmp_workspaces/issue-910-selection-pack` and `tmp_workspaces/issue-910-selection-excel`.

- [x] Final review P2: synchronize Page Site visibility after topology reconciles the final selected identity; cover selected-page removal, additions, reorder and empty topology. Userforms/pack/filepull tests, lint, docs checks and the retained pure-Go reproduction pass. Review pass 1 report and reproduction are preserved in `tmp_workspaces/final-review/pass-1/`; review resources are cleaned up.
- [x] Final review pass 2 verified `ffb16e22579fd3931aace7ee6bca6c68fe9f257b`: no confirmed findings; compiler/oforms tests and retained reproduction pass. Report preserved in `tmp_workspaces/final-review/pass-2/`; reviewer, watcher, setup terminal and temporary worktree released/removed. Excel/COM was not rerun after the selection fix; remote CI remains unrun.
- [x] PR #910 CI follow-up: update CLI/LSP fixture assertions for the added MultiPage/Page/TabStrip controls and supported MultiPage selection; both formerly failing tests pass. Local full Go run also hit the unchanged analyzer package's default 10-minute timeout; do not report it as passed.

- [x] Capture sequential Excel-authored and saved/reopened differential evidence, including empty Pages/Tabs and layout.
- [x] Add bounded lossless TabStrip arrays/flags and MultiPage x bookkeeping parsing and writing.
- [x] Extend new generation and atomic before/after topology/property compilation, preserving opaque state.
- [x] Align FormSpec, projection, editor intelligence, and Excel authoring/inspection contracts.
- [x] Verify blank/template pack, file pull, Excel persistence/runtime, and relevant package/bridge tests.
- [x] Update MS-OFORMS/pack/form specifications, ADR-0012, user documentation and changelogs; self-review.

Evidence and exact local commands are retained in the compiler's
`testdata/multipage-excel-authored/README.md` and
`testdata/multipage-excel-generated/README.md`. Focused persistence review
confirmed and then rechecked fixes for empty cached-tab edits and quadratic
signature membership checks. Local checks passed for userforms/pack/filepull,
CLI pack/form, 51 bridge tests, Go lint, VS Code TypeScript and docs. No remote
CI or publication is claimed.

# Issue #912 final review follow-up

- [x] Validate the pass 1 P2: CLI encoding preflight misclassifies referenced pictures with source suffixes.
- [x] Resolve validated picture roles before encoding and share command-local role metadata with symbol discovery, preserving full-project diagnostics.
- [x] Add CLI pack/file-push publication regressions for .bas/.cls/.frm pictures, a code-sidecar collision, and an internal symlink; verify persisted bytes and ordinary invalid-VBA rejection.
- [x] Run affected/full Go checks, corpus snapshots, lint and fresh installed-binary Excel picture gates.
- [x] Complete analyzer performance comparison: deterministic counters unchanged, no suspicious metric increases; concurrent test activity prevents isolated timing claims. Pass 2 confirmed the first fix and found a Linux case-distinct picture/sidecar identity defect.
- [x] Reproduce the pass 2 P2 in Linux, share host path keys across inventory/preflight/symbol discovery/file-push dependencies, and verify actual CLI pack/push preserve case-distinct sidecar code and reject invalid code before publication.
- [x] Complete final full Go/lint/format checks and fresh installed-binary Excel file-push/blank-pack/template-pack picture gates after the structural path-identity correction.
- [x] Run exceptional pass 3 on 2ab038f7: prior findings resolved, one new in-scope P2 confirmed for case-distinct Assets/ nested forms. Stop the full review loop at three passes.
- [x] Preserve pass 3 evidence and clean up all review resources. Correct reserved assets comparisons using the shared host path key; Linux inventory/actual-push regressions and full Windows Go/lint/format checks pass. Fresh installed-binary Excel file-push/blank-pack/template-pack gates pass with confirmed cleanup. Final correction receives coordinator regression/self-review verification; no fourth independent pass or remote CI is claimed.

# v0.35.0 bundled agent skill refresh

- [x] Correct backend routing and obsolete UserForm restrictions in the entrypoint; retain session-first guidance for Excel work.
- [x] Refresh forms reference with canonical source authority, containers, pictures and omission semantics; add focused pack reference and architecture discovery.
- [x] Verify installed reference links across providers and execute the documented examples through blank/template pack and file push with artifact readback. Full agentskill/CLI tests PASS (1.350s/136.836s), lint and skill validator PASS; installed-binary skill smoke PASS in tmp_workspaces/issue912-agent-skill-release/xlflow.
- [x] Complete self-review and docs/lint checks; prepare the skill update for commit and resume the authorized PR preparation. Validation evidence: tmp_workspaces/final-review/agent-skill-validation.md.
