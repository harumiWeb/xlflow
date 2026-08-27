# LSP startup and project-preparation profiling

This specification defines the developer-only measurement contract for large
workspace LSP performance work. It is additive observability: it does not
change diagnostics, symbol visibility, cancellation, generation checks,
overlay publication, or request results.

## Enabling telemetry

Performance logging is opt-in. Start the server with `xlflow lsp --stdio
--performance-log` (or set `xlflow.lsp.performanceLogging=true` in the VS Code
client). Records go to stderr, or to the path supplied by `--log-file`. With
the flag absent, the recorder is nil and stage measurements do not read the
clock or update a counter map.

Each record is a single structured log line. `elapsed_ms` is time spent
executing the stage and `wait_ms` is time spent waiting for a scheduler permit
when the stage has a wait component. `class` is `background` for workspace
index work and `interactive` for editor-request work. `path` is populated for
per-file stages. The existing request records remain compatible; their
`stage` field is the request operation and diagnostics additionally include a
`phase` of `fast` or `full`.

Workspace readiness is recorded independently: declaration indexing publishes
`operation="workspaceDeclarations/index/initial"` when interactive symbol
postings are complete, while semantic preparation publishes the existing
`operation="workspaceSymbols/index/initial"` after IR/CFG/call-site entries are
complete. Neither record changes LSP response fields.

## Stage names

The following names are stable and intended for benchmark/profile scripts:

| Stage                              | Boundary                                                                     |
| ---------------------------------- | ---------------------------------------------------------------------------- |
| `workspaceDiscovery`               | Initial source-file discovery under the configured source roots.             |
| `declarationIndexing`              | Per-file declaration/symbol extraction.                                      |
| `semanticIndexing`                 | Per-file procedure IR, raw call-site, and control-flow construction.         |
| `projectSnapshot`                  | Coherent project snapshot assembly from indexed files.                       |
| `projectResolver`                  | Project resolver construction, including TypeLib candidates when applicable. |
| `projectResolutionView`            | Resolution view/call-resolver setup for a project revision.                  |
| `projectResolutionMaterialization` | Resolution of indexed document IR into the project view.                     |
| `projectEffectSummary`             | Project effect-summary construction.                                         |
| `projectCapabilityPlan`            | Project capability requirement planning.                                     |
| `projectConstants`                 | Project-visible constant and constant-value construction.                    |
| `projectChange`                    | Revision change bookkeeping and impacted-file calculation.                   |
| `dependencyFingerprintUpdate`      | Procedure fingerprinting and dependency propagation.                         |
| `permitWait`                       | Time blocked waiting for a bounded analysis permit.                          |

Fast and Full diagnostics are also emitted as request records with
`phase="fast"` or `phase="full"`; hover, definition, completion, and
signature help retain their request operation records. The diagnostics stage
records produced by `analysisstats` continue to report analyzer-owned stages
and capability counters.

## Counter names

The initial workspace-index record emits a complete counter snapshot, including
zero values, so a profile can distinguish “not observed” from an omitted
instrumentation field. The counters are stderr-only and never enter an LSP
payload.

| Counter                            | Meaning                                                          |
| ---------------------------------- | ---------------------------------------------------------------- |
| `workspace_files_discovered`       | Source files returned by workspace discovery.                    |
| `workspace_declaration_builds`     | Successful per-file declaration builds.                          |
| `workspace_semantic_builds`        | Successful per-file semantic builds.                             |
| `project_snapshot_builds`          | Project snapshots assembled.                                     |
| `resolution_resolver_builds`       | Resolver constructions (revision cache misses).                  |
| `resolution_view_builds`           | Resolution views constructed (revision cache misses).            |
| `canonical_resolver_builds`        | Canonical project resolver/index constructions.                  |
| `procedure_resolver_views`         | Procedure-only resolver views derived from the index.            |
| `full_resolver_views`              | Full resolver views derived from the index.                      |
| `resolution_overlay_builds`        | Per-document resolution overlays built for both modes.           |
| `resolution_materializations`      | Document IR materializations into resolution views.              |
| `procedure_fingerprint_builds`     | Procedure fingerprints computed for dependency comparison.       |
| `procedure_fingerprint_reuses`     | Fingerprints served from a reusable dependency cache.            |
| `dependency_nodes_updated`         | Procedure dependency entries replaced or removed.                |
| `dependency_edges_updated`         | Reverse or outgoing call edges added or removed.                 |
| `procedures_revisited`             | Procedures inspected during an incremental dependency update.    |
| `fast_diagnostic_runs`             | Fast diagnostic runs started.                                    |
| `full_diagnostic_runs`             | Full diagnostic runs started.                                    |
| `background_permit_waits`          | Background workers that waited for an analysis permit.           |
| `interactive_permit_waits`         | Interactive workers that waited for an analysis permit.          |
| `project_cache_hits`               | Completed project products served from a dependency key.         |
| `project_cache_misses`             | Project product lookups without a reusable entry.                |
| `project_cache_rebuilds`           | Project product builds started after a dependency miss.          |
| `project_dependency_invalidations` | Files or dependency roots invalidated for a project product.     |
| `project_cache_reused_entries`     | Immutable resolver, constant, effect, or overlay entries reused. |

`procedure_fingerprint_reuses` is reserved for reuse reporting and is emitted
as zero until a reusable fingerprint cache is used. A counter increment is
reported with the stage and path that caused it; the initial snapshot reports
the accumulated totals.

The LSP dependency index stores catalog fingerprints and call edges from the
document analysis revision. A body-only edit should increment
`dependency_nodes_updated` only for changed procedures and should report
unchanged procedures through `procedure_fingerprint_reuses`; it must not invoke
JSON serialization of a complete `ProcedureIR`. Signature, module-context, or
resolution changes may revisit all procedures, while edge publication remains
incremental. Module declaration references use a module-level dependency node;
uncertain (dynamic, ambiguous, incomplete, or project-local unresolved)
resolution attaches to a project boundary so its reverse closure is invalidated
on the next revision. These counters are structural observations and are never copied
into LSP responses, normal CLI JSON, corpus snapshots, or review ledgers.

Project preparation has a second, product-level cache layer. Resolver indexes,
document overlays, constants, and effects are keyed by their explicit semantic
inputs and may be reused across adjacent workspace revisions. The effects
product retains direct procedure summaries and recalculates the changed
procedure plus its reverse caller closure; a declaration, TypeDB, recovery, or
completeness change may invalidate the project boundary. These five project
cache counters are emitted at the same stderr-only boundary.
They are observations for comparing cold, warm, body-edit, and dependency-edit
benchmarks and are not a performance threshold or a protocol contract.

Delta counter records use `outcome="counter"` with the increment in `value`.
The initial aggregate snapshot uses `outcome="counter_snapshot"`,
`value=0`, and the accumulated count in `total`; profile scripts should not
sum snapshot records as additional work.

## Deterministic benchmarks

Run commands from the repository root. On Windows, use the repository Go
wrapper so the tree-sitter CGO toolchain is selected:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/lspserver -run '^TestLSPStartupBenchmarkFixture$' -bench '^BenchmarkLSPStartup$' -benchmem -benchtime=1x -count 5
```

`BenchmarkLSPStartup` creates a deterministic multi-module workspace and
measures `Initialization`, `ImmediateInteractive`, `WhileIndexing`,
`AfterDeclarationBeforeSemantic`, and `AfterSemanticReady`. The latter three
use parser checkpoints rather than sleeps, so hover and definition are sampled
while the background index is at a known readiness boundary. The declaration
checkpoint must already resolve the cross-file target even though semantic
preparation is blocked; the semantic checkpoint must preserve the same result
after IR/CFG/call-site publication.

The giant single-module lifecycle benchmark uses the existing generated
ROneCOne-scale shape:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/lspserver -run '^$' -bench '^BenchmarkLSPIssue491LargeClass/Lifecycle/' -benchmem -benchtime=1x -count 1 -timeout 25m
```

It reports the actual first diagnostics publication after `didOpen`, first Fast
and Full diagnostic costs, hover during Full diagnostics, and definition during
workspace indexing. The generated fixture is deterministic and does not add
third-party source to the repository. `Lifecycle/DidOpenFirstPublication`
measures time-to-first-publication, while `Lifecycle/InitialFastDiagnostics`
isolates the bounded first-open Fast preview cost.

The opt-in local ROneCOne path remains available when a developer has a local
specimen:

```powershell
$env:XLFLOW_LSP_BENCH_FILE = 'C:\path\to\ROneCOne.cls'
$env:XLFLOW_LSP_BENCH_ROOT = 'C:\path\to\ronecone'
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/lspserver -run '^$' -bench '^BenchmarkLSPIssue491OptInROneCOne/' -benchmem -benchtime=1x -count 1 -timeout 25m
```

The equivalent POSIX command is:

```bash
XLFLOW_LSP_BENCH_FILE=/path/to/ROneCOne.cls XLFLOW_LSP_BENCH_ROOT=/path/to/ronecone \
  rtk proxy go test ./internal/lspserver -run '^$' -bench '^BenchmarkLSPIssue491OptInROneCOne/' -benchmem -benchtime=1x -count 1 -timeout 25m
```

The aggregate `task bench:lsp` task runs the existing and new deterministic
LSP benchmarks; `task bench:lsp-startup` selects only the startup scenario.

## Reproducible profiling

Collect one scenario at a time and keep profiles outside the repository. The
CPU profile identifies scheduler and execution time; the memory profile can be
viewed as allocation space, allocation objects, or live heap. For example:

```powershell
$profileDir = Join-Path $env:TEMP 'xlflow-lsp-profiles'
New-Item -ItemType Directory -Force -Path $profileDir | Out-Null
$bin = Join-Path $profileDir 'lsp-benchmark.test.exe'
$cpu = Join-Path $profileDir 'startup.cpu.pprof'
$mem = Join-Path $profileDir 'startup.mem.pprof'
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/lspserver -run '^TestLSPStartupBenchmarkFixture$' -bench '^BenchmarkLSPStartup/WhileIndexing$' -benchtime=1x -count=1 -timeout 25m -o $bin -cpuprofile $cpu
rtk go tool pprof -top $bin $cpu
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/lspserver -run '^TestLSPStartupBenchmarkFixture$' -bench '^BenchmarkLSPStartup/WhileIndexing$' -benchtime=1x -count=1 -timeout 25m -o $bin -memprofile $mem
rtk go tool pprof -sample_index=alloc_space -top $bin $mem
rtk go tool pprof -sample_index=alloc_objects -top $bin $mem
rtk go tool pprof -sample_index=inuse_space -top $bin $mem
```

The `-o` path above keeps the test binary available for `go tool pprof`. Repeat the same procedure
with `-bench '^BenchmarkLSPIssue491LargeClass/Lifecycle/FirstFullDiagnostics$'`
to isolate Full-diagnostic cost. CPU and allocation profiles must be collected
in separate runs when comparing startup discovery, semantic preparation, and
Full diagnostics. Use a warm repeated-request sub-benchmark after the cold
measurements to identify cache reuse.

On POSIX hosts replace the wrapper invocation with `rtk proxy go test`; the benchmark
names, profile flags, and pprof commands are unchanged. Do not enable the
ROneCOne benchmark in CI unless the source path is supplied explicitly.

## Interpretation and correctness boundary

`projectResolver`, `projectResolutionView`, `projectCapabilityPlan`, and
`projectResolutionMaterialization` are cache-miss/build boundaries. A repeated
request on the same revision should show request latency without another build;
the counter snapshot makes that distinction visible. For one project
preparation, `canonical_resolver_builds`, `procedure_resolver_views`, and
`full_resolver_views` should each increase once, while
`resolution_overlay_builds` increases once per document and mode. Permit wait
time is separate from execution time so an interactive request blocked behind a
background worker can be measured without attributing the delay to analysis.

Project snapshots retain the canonical syntax-local `DocumentIR` and attach a
read-only `ResolvedDocumentView`. The normal snapshot and Full-diagnostic paths
must therefore report `projectResolutionMaterialization` with zero results and
`resolution_materializations` must remain zero. `Resolve`/`Materialize` remains
available for compatibility consumers that explicitly require an independently
owned resolved document; such fallback calls are the only work counted by that
counter.

All instrumentation is developer telemetry. It must remain disabled by
default, use no LSP response fields, and preserve cancellation and generation
safety. A profile is evidence about the selected benchmark checkpoint, not a
claim that all workspace work has completed.
