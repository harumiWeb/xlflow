# Go Test Parallelization Policy

This repository uses Go package-level parallelism by default and adds
intra-package parallelism only where test isolation is explicit. The purpose is
to improve throughput without making process-wide state or timing-sensitive
tests order-dependent.

## Safety policy

Tests may use `t.Parallel()` when they only perform pure computation, read
immutable fixtures, use an independent `t.TempDir()` workspace, run an
isolated analyzer/linter invocation, or verify read-only deterministic
snapshots.

Tests remain serial when they use `t.Setenv`/`t.Chdir`, mutate package or
process-wide state, use fixed shared paths, coordinate locks or child
processes, depend on timing or execution order, invoke Excel/COM/VBE or the
local oracle, or publish/update snapshots. A `t.Parallel()` test must not call
an API that changes process-wide environment or working-directory state.

The `analyze`, `lint`, and pure VBA data-flow fixtures follow the safe policy.
The existing `cfg` and `procedureir` parallel tests are retained. CLI, LSP,
Excel/COM, coordination, process-sensitive, and environment-mutating tests
are intentionally outside the initial parallelization scope.

## Real-world corpus

Each corpus project is analyzed in its own temporary workspace. Native and
third-party projects run as named parallel subtests, with a fixed maximum of
two active projects. Results are stored by input index and merged in manifest
order before the existing deterministic diagnostic and failure sorting. This
keeps output stable while limiting peak memory and GC pressure on smaller CI
runners.

Diagnostic normalization happens inside each isolated worker. Snapshot set
construction, review evaluation, verification, and publication happen only
after all project workers complete. Snapshot updates retain the existing
all-or-nothing contract: workers never publish individual files, and
`WriteSnapshotSet` replaces the complete tree once, after successful analysis.

## Timing reference

These are observations from the Windows development workstation on 2026-08-08
using Go 1.26.5, `-count=1`, and the repository's
`scripts/dev/go.ps1` wrapper. They are comparison points, not CI thresholds.

| Run                           |    Before |                After |
| ----------------------------- | --------: | -------------------: |
| `go test ./...` (uncached)    |  59.198 s | 50.951 s median of 3 |
| real-world corpus verify-only | 171.054 s | 86.584 s median of 2 |

Repeat timing measurements with the same toolchain and machine before drawing
performance conclusions. Use `-json` test output to retain package/test timing,
and run the corpus through `rtk task corpus:test` so the opt-in environment and
verify-only snapshot mode are explicit.

## Verification commands

On Windows, invoke Go through `scripts/dev/go.ps1`:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test -count=1 -json ./...
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test -race ./...
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test -shuffle=on -count=10 ./...
rtk task corpus:test
```

The opt-in corpus suite should also be exercised with race detection and
shuffled repetition during development. Snapshot update runs must be followed
by a worktree check; normal verification must never modify committed snapshot
files.
