# ADR-0045: Bounded Procedure-Level Parallelism for Large Batch Files

## Status

Accepted

## Context

ADR-0043 added a bounded worker pool for independent file-level diagnostic
work. That pool scales projects with many files, but a project whose work is
concentrated in one module still has only one file job. The procedure-local
loop in `internal/analyze` can therefore serialize hundreds or thousands of
independent rule evaluations.

Procedure evaluation is not safe to parallelize by simply starting another
full-size worker pool inside every file worker. The analyzer also has to keep
the existing contracts for immutable project/file facts, tree-sitter parser
lifetime, one-time file diagnostics, cancellation, finding multiplicity, and
deterministic suppression/JSON output.

## Decision

Add bounded procedure-level parallelism to batch `xlflow analyze` for files
whose procedure workload is large enough to justify scheduling overhead. The
threshold is an internal, benchmark-calibrated scheduling gate; it is not a
public configuration contract. Ordinary small files retain the existing
direct loop fast path.

Procedure batches run only after project and file analysis state has been
constructed and frozen. Shared project resolution, effect/object summaries,
analyzer indexes, module facts, and procedure facts are read-only inputs.
Flow state, rule overlays, statistics deltas, caches that are not explicitly
concurrent-safe, and finding buffers are procedure-local. File-global or
one-time diagnostics remain in the file coordinator where their existing
deduplication semantics can be preserved.

File jobs and procedure batches use one shared bounded execution budget. A
large file may be scheduled as procedure batches, but it must not create an
independent full-size pool while occupying a file-level pool of the same
size. The scheduler may choose file-level or intra-file work according to the
workload shape, provided the total active analyzer work stays within the
shared bound.

Each procedure or batch writes to a stable procedure-indexed result slot.
The coordinator merges completed slots in source order and then invokes the
existing deterministic finding sort and suppression/finalization stages.
Completion order, map iteration order, and worker assignment therefore do
not affect finding content, multiplicity, source order, suppression, or JSON
output. Any file-level one-time diagnostic is deduplicated during this stable
merge rather than through an unsynchronized shared map.

Procedure workers consume owned IR, immutable facts, and owned source values;
they do not retain tree-sitter nodes or a `ParsedDocument`. Tree and borrowed
source access remains inside the `ParsedDocument.Read` lifetime boundary. A
rule that still requires that boundary is executed by the coordinator or
converted to an owned, read-only projection before procedure work is queued.

The shared analysis context propagates cancellation to queued and active
procedure work. Queue submission stops promptly, active rule scans continue
to check the context at their existing cancellation boundaries, and an
error or cancellation waits for active workers before returning. Partial
findings are not published after cancellation or worker failure.

This decision changes the execution strategy only. It adds no CLI flag,
configuration key, diagnostic ID, finding field, LSP capability, or JSON
schema field.

## Consequences

- Very large single-module projects can use multiple cores while retaining a
  bounded total analyzer concurrency.
- Many-small-file projects continue to use the file-level pool without
  multiplying concurrency, and ordinary small modules avoid procedure-worker
  overhead.
- Stable result slots and a source-order merge add bounded result storage and
  coordination work, but preserve the existing deterministic finalization
  contract.
- Stage totals can include concurrent procedure work and therefore remain
  work-attribution metrics; the existing total elapsed measurement remains
  the wall-clock measure.
- Every mutable value reachable from procedure analysis must either be frozen,
  made explicitly concurrent-safe, or moved into a procedure-local buffer.

## Alternatives Considered

1. **Create one full-size procedure pool inside each file worker** - Rejected
   because the product of file and procedure worker counts oversubscribes CPU
   and memory on multi-file projects.
2. **Start one goroutine per procedure** - Rejected because a large module can
   create unbounded scheduling, stack, result-buffer, and cancellation pressure.
3. **Parallelize parser/IR construction at the same time** - Rejected because
   procedure workers need a frozen owned input, and parser lifetime/read
   boundaries are a separate concern with less direct evidence of benefit.
4. **Expose a procedure-worker count or threshold to users** - Rejected
   because it would create a tuning and compatibility contract before the
   benchmark data establishes a user need.
5. **Publish findings as workers finish** - Rejected because completion order
   would make deduplication, multiplicity, suppression, and JSON output
   dependent on scheduling.

## Evidence

- Requirements: xlflow issues #675 and #670.
- Existing bounded file execution and procedure-local loop:
  `internal/analyze/analyzer.go`.
- Existing file-level cancellation and deterministic result-slot regression:
  `internal/analyze/bounded_parallel_test.go`.
- Synthetic large-module benchmark and profiling dimensions:
  `internal/analyze/benchmark_test.go` and
  `internal/analyze/profiling_test.go`.
- Immutable IR and parser lifetime contract:
  `docs/specs/vba-analysis-ir.md` and ADR-0021.
- Existing file-level concurrency decision amended by this ADR:
  `docs/adr/ADR-0043-bounded-batch-analysis-parallelism.md`.

On the Windows development host used for this change, three `-benchtime=1x`
trials at `-cpu=4` showed median serial versus bounded times of approximately
62 versus 64 ms for 100 procedures, 346 versus 253 ms for 500, 840 versus
574 ms for 1,000, and 1,923 versus 1,331 ms for 2,000. The 100-procedure
case is within ordinary run-to-run noise while the larger cases show the
intended wall-clock improvement. These measurements calibrate the internal
500-procedure gate; they are hardware-local evidence, not CI timing thresholds.

## Supersedes

- The procedure-level-parallelism deferral in ADR-0043's Consequences.
  ADR-0043 remains authoritative for file-level bounded execution and its
  shared immutable-state and no-partial-publication requirements.

## Superseded by

- None

## Related

- Issue #675
- Issue #670
- ADR-0021, ADR-0022, ADR-0023, ADR-0043
- `docs/specs/vba-analysis-ir.md`
- `docs/specs/static-analysis-corpus.md`
