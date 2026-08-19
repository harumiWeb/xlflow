# ADR-0043: Bounded Parallelism for Batch Analysis

## Status

Accepted

## Context

Issue #646 established shared parsing, IR/CFG, project indexes, and reusable
analysis artifacts as the scalability baseline for `xlflow analyze`. Issue
`#654` follows those changes because parallelizing duplicated or quadratic
work would amplify the original problem rather than solve it.

The batch pipeline constructs project-wide immutable context first, then runs
multiple independent diagnostics for each parsed source file. That file phase
was still entirely sequential and dominated the synthetic 500- and
1000-procedure benchmarks.

## Decision

Run the independent per-file diagnostic phase through a bounded worker pool.
The default worker limit is the smaller of `runtime.GOMAXPROCS(0)` and the
number of parsed files, with a minimum of one worker. Project-wide context
construction, source discovery, parsing, IR/CFG construction, and final
suppression remain coordinated stages.

Workers receive file indexes and write only to their own result slots. The
coordinator concatenates results in input-file order and applies the existing
finding sort and suppression finalization. This preserves deterministic
findings, preflight findings, and warnings regardless of worker completion
order.

Workers use the existing cancellable context. An error or cancellation stops
new work and is returned without publishing partial analysis results. Shared
IR, CFG, TypeDB, project indexes, and Intel snapshots are read-only during the
parallel phase.

## Consequences

- Large multi-file projects use available CPU cores during independent rule
  execution and reduce wall-clock analysis time.
- Small projects avoid unnecessary workers because the pool is capped by file
  count.
- Per-stage elapsed totals may exceed wall-clock batch time because stage
  measurements now include concurrent worker execution; the existing metrics
  remain useful for work attribution, while `analyze_total` remains the
  wall-clock measure.
- Procedure-level parallelism and parallel parsing remain deferred until
  separate measurements prove additional benefit and safety.

## Alternatives Considered

1. **One goroutine per file** - Rejected because large projects could create
   unbounded scheduler and memory pressure.
2. **Parallelize project-wide fixed-point work** - Rejected because those
   stages have shared dependency state and were already improved algorithmically
   by #646; parallelizing them would complicate convergence and determinism.
3. **Expose a user-configurable worker-count CLI option** - Rejected for this
   change because it would add a public tuning contract without evidence that
   users need it. The runtime CPU limit is a safe default.
4. **Parallelize parsing and IR/CFG construction first** - Deferred because
   those artifacts are shared inputs and the file diagnostic phase has the
   clearest measured wall-clock opportunity.

## Evidence

- Requirements: issues #646 and #654.
- Batch pipeline and shared artifact implementation:
  `internal/analyze/analyzer.go`.
- Deterministic stage profiling and synthetic scaling benchmark:
  `internal/analyze/benchmark_test.go` and
  `internal/analyze/profiling_test.go`.
- Deterministic serial/parallel regression:
  `internal/analyze/bounded_parallel_test.go`.
