---
name: review-static-analysis-corpus
description: Review xlflow real-world static-analysis corpus diagnostics, classify committed evidence as true positive, investigate suspected false positives, remediate confirmed false positives with focused regression tests, update the diagnostic review ledger and snapshots, and interpret reviewed-only quality metrics. Use for work involving `testdata/static-analysis-corpus`, `reviews/diagnostics.jsonl`, corpus snapshot deltas, TP/FP/unreviewed classification, forbidden diagnostic contracts, or `corpus:metrics`.
---

# Review Static-Analysis Corpus

Use the corpus as two complementary records: snapshots preserve observed output, while the review ledger preserves semantic judgments. Never treat a snapshot row as proof that a diagnostic is correct, and never update a snapshot merely to silence a delta.

## Ground the Work

Before classifying or changing evidence:

1. Read `docs/specs/static-analysis-corpus.md`, especially **Reviewed diagnostic evidence**, **Operational workflow**, and **Reading a snapshot delta**.
2. Read `docs/adr/ADR-0032-reviewed-static-analysis-corpus-evidence.md` when the task touches the evidence model or ownership boundary.
3. Inspect `testdata/static-analysis-corpus/reviews/diagnostics.jsonl`, the relevant snapshot rows, the rule registry entry, the rule implementation, and its focused tests.
4. Check `tasks/lessons.md` for relevant recurrence-prevention rules.
5. Inspect the worktree and preserve unrelated user changes.

If the rule specification, registry metadata, implementation, and focused tests disagree about severity, surface, or semantics, stop classification and report the contract drift first. Do not choose whichever source makes the proposed classification convenient.

Treat the specification as the schema authority. If this Skill and the specification disagree, follow the specification and report the drift. Modify this Skill only when the user explicitly asks to maintain it.

## Choose One Workflow

Classify the task before editing:

- **Review-only:** decide whether current emitted diagnostics are true positives. Add only evidence that has actually been reviewed.
- **Suspected FP:** investigate and reduce a diagnostic, but do not add a false-positive ledger row while the analyzer still emits the reviewed false positive.
- **FP remediation:** prove the false positive, add a focused failing fixture, fix the root cause, update observations, then commit the forbidden evidence.
- **Snapshot delta:** explain added, removed, or moved observations before accepting any baseline update.

Keep one pull request centered on one rule or one analyzer root cause. A small batch of diagnostics for that same rule is acceptable; unrelated rule judgments should be separate work.

## Review Current Diagnostics

1. Select one rule and a manageable batch, normally 10-20 occurrences. Use `rtk task corpus:review-candidates -- <RULE> [LIMIT]` to list current unreviewed start-only snapshot rows without running the analyzer.
2. Locate every candidate in the snapshot and source. Preserve duplicate occurrences; do not deduplicate by eye.
3. Read the rule contract and focused tests before deciding. Evaluate the source facts required by that rule, not whether the VBA merely compiles.
4. Record a true positive only when the diagnostic matches the intended rule semantics at that exact source location.
5. Use the full normalized identity: project, project-relative file, surface, code, severity, and all four range coordinates. Record multiplicity separately in `count`; it is not part of the identity.
6. Do not infer an end position from the start-only snapshot format. Obtain the normalized full range from the actual analyzer/corpus diagnostic or a focused diagnostic probe. Prefer an existing owning-package test helper that exposes normalized diagnostics. If no targeted probe exists, add a small probe only when repository edits are in scope and either retain it as a useful focused test or remove that probe-only change before handoff. Otherwise report range collection as a blocker and leave the ledger unchanged.
7. Add canonical, sorted `true-positive` rows to `testdata/static-analysis-corpus/reviews/diagnostics.jsonl`, with a concise rationale explaining the source fact that makes the finding valid.
8. Leave every undecided or unexamined occurrence absent from the ledger. Absence means `unreviewed`; never serialize an unreviewed classification.

If review indicates a likely false positive, switch to the suspected-FP workflow. Do not label it true positive for convenience.

## Investigate a Suspected False Positive

1. State the rule contract and the exact reason the real-world diagnostic appears inconsistent with it.
2. Reduce the relevant syntax and data/control-flow facts to the smallest focused fixture that still reproduces the diagnostic.
3. Search existing fixtures and VBE oracle evidence before adding a duplicate case.
4. Decide whether the behavior is compile-equivalent:
   - For compile-equivalent rules, follow `docs/specs/vbe-oracle.md` and the repository oracle instructions. Run known controls and focused cases sequentially on Windows with Excel/VBIDE.
   - For policy, maintainability, runtime-risk, portability, or Excel-specific rules, VBE compile success is not evidence of a false positive.
5. If the task does not include implementation, report or create a focused follow-up issue with the project, file, range, rule, reduced example, expected behavior, and evidence. Leave the ledger unchanged.

Do not create a forbidden row for an outstanding defect: false-positive ledger rows are durable evidence for a remediated diagnostic and must be enforceable by the current analyzer.

## Remediate a Confirmed False Positive

Perform these steps in order:

1. Add a focused regression test that reproduces the analyzer defect and observe it fail for the intended reason.
2. Fix the analyzer root cause without weakening unrelated valid diagnostics.
3. Run the focused test and neighboring rule tests.
4. Run the Excel-free static-analysis and committed-oracle contracts.
5. Run the real-world corpus in verify-only mode and inspect the complete delta.
6. If the behavior change is deliberate and supported by the focused test, update snapshots explicitly and review the generated diff.
7. Add or retain the real-world `false-positive` row as a forbidden contract. Point `regression_test` to the exact Go test using `<path>::<test-name>`.
8. Use `regression_exception` only when a focused fixture is genuinely impractical, and explain why. Do not use it to avoid writing a feasible test.
9. Run the corpus again without update mode so both snapshot and review contracts are checked together.

The focused test protects the analyzer root cause. The real-world forbidden row protects the original integration boundary. Both are normally required.

## Maintain the Ledger Safely

Follow the exact schema and ordering rules in `docs/specs/static-analysis-corpus.md`.

- `count` is explicit multiplicity for one exact identity. For TP rows it is the minimum expected count; extra matching emissions remain visible as unreviewed.
- FP rows remain in the ledger after remediation and continue to contribute to reviewed-only metrics.
- Use `allowed_occurrences` only when separately reviewed, legitimate diagnostics share the exact normalized identity with the remediated FP because the contract lacks rule-specific path/context detail.
- Before setting `allowed_occurrences`, enumerate and explain the legitimate colliding occurrences. Use the smallest observed ceiling.
- Remember its limitation: an allowed occurrence disappearing while the forbidden occurrence returns under the same identity cannot be distinguished.
- Keep rationales factual and specific. Avoid conclusions based only on VBE compilation or general-language intuition.
- Never hand-edit snapshot output to make a contract pass.

After editing the ledger, let loader validation catch unknown fields, non-canonical JSON, ordering, duplicates, invalid paths/ranges, unsupported metadata, and stale regression-test references.

## Diagnose Snapshot Deltas

Read deltas by rule and retain multiplicity:

1. Check whether project source, vendor pin, file mapping, or line positions changed.
2. If source is unchanged, reproduce the rule behavior in a focused fixture.
3. Run verification twice when nondeterminism is plausible. Do not update an unstable baseline.
4. Explain each deliberate diagnostic addition/removal category before updating snapshots.
5. Update snapshots only after the focused semantic decision is established.

A removed diagnostic is not automatically a fixed FP, and an added diagnostic is not automatically a TP.

## Verify

Match verification to the workflow:

- **Read-only candidate scouting:** use `corpus:review-candidates`, then run the owning rule tests and `corpus:metrics`; do not update snapshots or run the full corpus unless needed to reproduce the candidate.
- **TP ledger change:** run the committed review/metrics tests and `corpus:test` so exact normalized ranges are checked against current output.
- **FP remediation:** run focused and neighboring tests, static-analysis contracts, committed oracle contracts, verify-only corpus, the explicit snapshot update when justified, and a final verify-only corpus run.
- **Documentation-only change:** run the relevant docs and format checks; do not run the corpus merely because this Skill was edited.

On Windows, run Go through the repository wrapper. Start narrow, then broaden as required above:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test <owning-packages> -run '<focused-test>' -count=1
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/staticanalysis/... -count=1
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/oracle -run '^TestCommittedFixtureContractsWithoutExcel$' -count=1
rtk task corpus:review-candidates -- <RULE> [LIMIT]
rtk task corpus:metrics
rtk task corpus:test
```

For an intentional observation change, run this only after reviewing the verify-only delta:

```powershell
rtk task corpus:update-snapshots
rtk task corpus:test
```

Finish with the relevant repository checks, including:

```powershell
rtk pnpm docs:check
rtk git diff --check
```

Do not run the Excel/VBE oracle unless the decision depends on actual VBE compile semantics. If it is required, treat timeout, modal dialogs, compile invocation failure, COM failure, or uncertain cleanup as infrastructure failure and do not promote evidence.

## Report the Evidence

Report:

- workflow used: review-only, suspected FP, FP remediation, or snapshot delta
- rule, project, file, surface, exact range, and multiplicity reviewed
- TP/FP decision and the concrete rationale
- focused regression test and analyzer change for each remediated FP
- snapshot delta and why it is accepted or left unchanged
- reviewed, TP, FP, unreviewed, and precision metrics before/after when changed
- commands and results, including VBE case IDs when run
- any unreviewed, ambiguous, or unverified items left in scope

Interpret `corpus:metrics` as the committed evidence ledger's reviewed-only precision, not the analyzer's defect rate across all corpus diagnostics.

If the workflow itself or evidence architecture changes, update the corpus specification and use `$adr-manager` to decide whether ADR-0032 needs an amendment or a successor decision.
