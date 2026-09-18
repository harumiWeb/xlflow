---
name: analyzer-performance-review
description: Detect, explain, and prevent xlflow static-analyzer performance regressions by comparing the current worktree with origin/main using deterministic analysis telemetry and the ROneCOne/std-vba real-world benchmarks. Use this skill whenever changes touch analyzer participation, interprocedural summaries, fixed-point worklists, CFG/data-flow traversal, parser interpretation consumed by diagnostics, semantic-query invalidation, caches or indexes, or when preparing/reviewing a PR with non-trivial changes under internal/analyze or internal/staticanalysis—even if the user only asks whether performance regressed.
---

# Review Analyzer Performance

Treat analyzer performance as a work-scaling contract. Wall-clock time is useful evidence, but deterministic counters reveal accidental broadening even when machine noise makes elapsed time look unchanged or faster.

## Establish the scope

1. Inspect `git status`, the diff against `origin/main`, and `tasks/lessons.md`.
2. Identify which dimensions could grow: candidate procedures, participant closure, summary evaluations, CFG walks, invalidations, semantic kernels, allocations, or retained state.
3. Fetch `origin/main` before a PR gate unless the user explicitly requires an offline comparison. Record the resolved base and head SHAs.
4. Preserve unrelated changes. Benchmark the current worktree as-is; use the helper's detached temporary worktree for the base.

Run this skill proactively when a change broadens an interprocedural analysis or adds a fallback that can revisit procedures. Unchanged findings do not prove unchanged scalability.

## Add deterministic coverage first

When the changed mechanism has a countable unit of work, add or extend a synthetic test or benchmark telemetry before relying on real-world timing.

- Include many unrelated procedures and a small positive dependency chain.
- Assert that unrelated work remains constant or absent.
- Assert that actual participants grow linearly when the feature requires it.
- Keep timing assertions out of ordinary tests.
- Prefer counters such as summary evaluations, CFG walks, worklist revisits, participant procedures, index builds, and semantic kernel runs.

A useful test demonstrates both sides: sparse/unrelated code does not increase work, while a positive control proves the counter observes real work.

## Run the base/head comparison

On Windows, invoke the bundled helper through `rtk`:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\.agents\skills\analyzer-performance-review\scripts\Compare-AnalyzerPerformance.ps1 -BaseRef origin/main -Count 3
```

The helper:

- resolves and records the exact base/head commits;
- creates a detached base worktree under the system temporary directory;
- runs the same `-benchtime=1x -count=N` command for both sides through `scripts/dev/go.ps1`;
- covers ROneCOne and std-vba cold, warm, local-edit, and dependency-edit modes;
- writes raw logs, `comparison.json`, and `report.md` beneath an ignored `.tmp*` directory;
- removes only the validated temporary base worktree.

Use `-Projects ronecone` for a quick stress-fixture investigation. Use both default projects before declaring a PR free of analyzer performance regressions. Keep the command, report, and raw logs on the same `Count`; never combine evidence collected with different sample counts.

Use `-FailOnRegression` only as an explicit gate. By default the helper reports suspicious increases without deciding whether an intentional semantic expansion is acceptable.

## Interpret the evidence

Investigate every unexplained positive delta in a deterministic `counter_*` metric connected to the change. A one-unit increase can be meaningful when it changes complexity or invalidation scope.

Use these default noise thresholds as prompts for investigation, not universal correctness limits:

- `B/op` or `allocs/op`: more than 5% increase;
- `ns/op`: more than 10% increase, followed by a rerun when other evidence does not agree;
- deterministic counters: any unexplained increase.

Compare related metrics together. For example, stable candidate procedures with sharply higher summary evaluations usually means repeated work rather than legitimately broader participation.

Do not average away structural regressions. Prefer the median for timing and allocation samples, while expecting deterministic counters to be identical across repeated samples.

## Diagnose and suppress a regression

1. Locate the first counter that expands and map it to the owning worklist, traversal, cache, or invalidation path.
2. Distinguish broader legitimate participation from repeated unrelated work.
3. Add a focused regression test that fails on the structural cause.
4. Fix the root cause by narrowing dependencies, caching immutable indexes, bounding state, or removing obsolete broad fallbacks.
5. Re-run focused tests and the base/head comparison.
6. Run `rtk task corpus:test` to prove diagnostic snapshots and review contracts remain intact.
7. Run the affected package tests, race coverage where shared state is involved, lint, and `git diff --check`.

Do not update corpus snapshots merely to make a performance fix pass. Any diagnostic delta needs its own semantic explanation and the `review-static-analysis-corpus` workflow.

## Review and report

Before completion, use an independent final-review workflow with an explicit
scope suitable for the current diff. On Orca-enabled machines, use the
machine-global `orca-supervised-final-review` skill; if Orca is unavailable,
report the final review as unverified rather than invoking the removed
repository-local fallback. Report:

- resolved base/head revisions and whether the head included uncommitted changes;
- exact benchmark command and sample count;
- median `ns/op`, `B/op`, and `allocs/op` by project/mode;
- relevant deterministic counter deltas and their explanation;
- focused scalability test and positive control;
- corpus/package/race/lint results;
- raw log and generated report paths;
- remaining noise, infrastructure failures, or unverified cases.

Do not claim “no regression” from elapsed time alone or when the base and head commands, corpora, sample counts, or machine differ.
