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

## End-to-end startup correlation

When performance logging is enabled, the VS Code client creates an anonymous
`startup_id` for each language-client start attempt and passes it to the child
process through `XLFLOW_LSP_STARTUP_ID`. Manual and automatic process restarts
begin a fresh attempt with a new ID. The client writes records with
`operation="lsp/startup"`, `side="client"`, `event`, `elapsed_ms`, and
`wall_time_unix_ms`. The server writes the same operation and ID with its local
monotonic elapsed time and `wall_time_unix_ns`. IDs are omitted when performance
logging is disabled, and no source, workspace name, or user-identifying data is
included.

The client records `extensionActivationStart`, CLI availability start and
completion, workspace/configuration discovery, `languageClientStart`, the
nearest public process-spawn boundary, `initializeResponseReceived`,
`initializedSent`, and the first successful `didOpen`. The
`vscode-languageclient` API does not expose every JSON-RPC boundary, so the
initialize and process events are explicitly nearest stable boundaries; the
server lifecycle events provide the precise server-side counterparts.

The server records `serverProcessStart`, `serverConstructed`,
`initializeHandled`, `initializedHandled`, `didOpenHandled`, declaration and
semantic index start/ready events, and the first non-empty hover, definition,
and completion result. The CLI captures the server baseline before its
best-effort TypeLib preparation hook, so that cold preparation remains visible
in the server startup interval. An empty result, error, or cancellation is not
a first-success event. Each event is emitted at most once per startup ID.

Startup telemetry is observational only. It does not add an LSP notification,
change a response, delay document lifecycle handling, or change availability,
restart, generation, or cancellation behavior. `wall_time_unix_*` is used only
to correlate the two processes; elapsed durations are compared within the
process that produced them.
Initial declaration jobs use a bounded priority queue. Active/open documents
are P0, uniquely matched direct module hints are P1, and the remaining files
are P2. A promotion updates an existing queued job; it never creates duplicate
parsing work, and an already running job is not preempted. Open-document
overlays remain authoritative through their existing generation checks. The
priority counters describe scheduling only and do not change readiness or
result ordering.

Readiness records use `outcome="readiness"` and report the elapsed time from
workspace startup to the first accepted declaration at each boundary:
`active_document_declaration_ready_ms`,
`referenced_document_declaration_ready_ms`, and
`workspace_declaration_ready_ms`. An active overlay declaration counts as the
P0 boundary when it supersedes a queued saved-file job.

The VS Code client supplies `initializationOptions.declarationPriority` with
`activeDocumentUri` and `openDocumentUris`, then sends the optional
`xlflow/didChangeActiveDocument` notification as the active editor changes.
These values are scheduling hints only; `didOpen` and `didChange` remain the
source of truth for unsaved text.

## Automatic code actions and transport readiness

Code-action requests use one worker and at most 16 pending requests; the
ordered stdio receive loop captures the document revision and continues to
serve interactive requests and lifecycle notifications. `$/cancelRequest`
cancels matching numeric or string request IDs. A newer automatic action for
the same document supersedes older automatic actions, while explicitly invoked
actions are not coalesced. Changes, close/reopen, and shutdown cancel affected
work. Canceled, superseded, or excess requests return `RequestCancelled`
(`-32800`), and edits from obsolete snapshots are never returned.

Cancellation and final response publication share a request-local mutex in
addition to document-lifecycle validation. If the client advertises
`workspace.workspaceEdit.documentChanges`, open-document actions carry their
captured version; disk-owned documents use a null version. Clients without
that capability retain `changes` edits. Versioned edits protect against
client-side changes that have not reached the server yet.

Cancellation is cooperative, not a bound on response or shutdown latency.
The shared initial `AnalysisSnapshot.ParsedDocument` / `ast.ParseDocument`
construction does not accept a context, and waiting for its parse mutex is
also not interruptible. On an unparsed large file, cancellation is observed
after that construction returns; the single action worker and shutdown may
wait until then. Canceled results are still discarded. Making initial parsing
interruptible requires a separate change to shared parse ownership, retry
after cancellation, and snapshot retirement; the action queue does not detach
unowned parser goroutines to simulate immediate cancellation.

`context.only` filters action families before their computation. VB044 being
disabled skips its parse and fix scan. When enabled, its code actions evaluate
only the procedure-name-constant rule and apply the same inline suppressions;
they must not call whole-file `LintParsed` or construct ProcedureIR/CFG.
Documentation actions retain their document-symbol projection.

`didChange` must not wait for a background reader or an in-progress parse of
the previous revision. It uses incremental tree-sitter parsing only when both
the parsed snapshot and its tree lease are immediately available. Otherwise it
publishes a lazy successor snapshot with `parse_mode="deferred_full"`; the new
revision is parsed by its first consumer. `fallback_reason` distinguishes a
busy/unavailable incremental tree from a full-document replacement. Completed
immutable procedure artifacts may carry into the successor, but no stale tree
or result may be published for the new version. Retiring the obsolete snapshot
invalidates it immediately; if its parse mutex is busy, parsed resource cleanup
finishes asynchronously. Server shutdown waits for cleanup of active snapshots.

`operation="textDocument/codeAction"` starts at request dispatch and ends when
its response is prepared, including time in the bounded action queue. Other
handler timings do not include time spent waiting before dispatch. Measure
client request-to-first-nonempty-response latency as well as handler duration,
and include the automatic code actions emitted by the editor. Fast diagnostics
publication is not proof of Full diagnostics completion.

`TestJSONRPCCodeActionDoesNotBlockInteractiveRequestsOrCancellation` blocks
the action projection at a deterministic checkpoint and verifies a real
JSON-RPC hover response, cancellation, and change/close/reopen invalidation.
`TestCodeActionRequestsBoundedQueueAndAutomaticSupersession` covers the work
bound and removal of obsolete queued requests. These are correctness and
scheduling checks, not wall-clock speedup claims.

For manual comparisons record the exact extension source, executable path,
build revision, active file, profile settings, and request ordering. F5 extension
compilation alone does not update a separately installed Go executable. Use a
matching development build, and do not infer its revision from its modification
time or a `dev` version string.

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

| Counter                             | Meaning                                                                               |
| ----------------------------------- | ------------------------------------------------------------------------------------- |
| `workspace_files_discovered`        | Source files returned by workspace discovery.                                         |
| `workspace_declaration_builds`      | Successful per-file declaration builds.                                               |
| `workspace_semantic_builds`         | Successful per-file semantic builds.                                                  |
| `declaration_priority_p0_jobs`      | Initial declaration jobs started for active/open documents.                           |
| `declaration_priority_p1_jobs`      | Initial declaration jobs started for uniquely identified direct dependencies.         |
| `declaration_priority_p2_jobs`      | Initial declaration jobs started for the remaining workspace.                         |
| `declaration_promotions`            | Queued declaration jobs moved to a higher priority.                                   |
| `declaration_priority_hits`         | Promotion hints that matched an already queued declaration job.                       |
| `interactive_index_builds`          | Snapshot-scoped interactive declaration indexes built.                                |
| `interactive_index_hits`            | Interactive lookups served from an existing snapshot index.                           |
| `procedure_catalog_builds`          | Procedure catalogs built for an interactive document revision.                        |
| `procedure_catalog_reuses`          | Interactive lookups that reused a snapshot-owned procedure catalog.                   |
| `interactive_exact_queries`         | Exact-name lookups issued through the interactive index.                              |
| `interactive_prefix_queries`        | Prefix lookups issued through the interactive index.                                  |
| `interactive_qualified_queries`     | Qualified-name lookups issued through the interactive index.                          |
| `full_document_symbol_builds`       | Full document symbol models built for LSP requests.                                   |
| `interactive_full_symbol_fallbacks` | Interactive lookups that fell back to full document symbol extraction.                |
| `project_snapshot_builds`           | Project snapshots assembled.                                                          |
| `resolution_resolver_builds`        | Resolver constructions (revision cache misses).                                       |
| `resolution_view_builds`            | Resolution views constructed (revision cache misses).                                 |
| `canonical_resolver_builds`         | Canonical project resolver/index constructions.                                       |
| `procedure_resolver_views`          | Procedure-only resolver views derived from the index.                                 |
| `full_resolver_views`               | Full resolver views derived from the index.                                           |
| `resolution_overlay_builds`         | Per-document resolution overlays built for both modes.                                |
| `resolution_materializations`       | Document IR materializations into resolution views.                                   |
| `procedure_fingerprint_builds`      | Procedure fingerprints computed for dependency comparison.                            |
| `procedure_fingerprint_reuses`      | Fingerprints served from a reusable dependency cache.                                 |
| `dependency_nodes_updated`          | Procedure dependency entries replaced or removed.                                     |
| `dependency_edges_updated`          | Reverse or outgoing call edges added or removed.                                      |
| `procedures_revisited`              | Procedures inspected during an incremental dependency update.                         |
| `fast_diagnostic_runs`              | Fast diagnostic runs started.                                                         |
| `full_diagnostic_runs`              | Full diagnostic runs started.                                                         |
| `background_permit_waits`           | Background workers that waited for an analysis permit.                                |
| `interactive_permit_waits`          | Interactive workers that waited for an analysis permit.                               |
| `project_cache_hits`                | Completed project products (including capability plans) served from a dependency key. |
| `project_cache_misses`              | Project product (including capability-plan) lookups without a reusable entry.         |
| `project_cache_rebuilds`            | Project product builds started after a dependency miss.                               |
| `project_dependency_invalidations`  | Files or dependency roots invalidated for a project product.                          |
| `project_cache_reused_entries`      | Immutable resolver, constant, effect, overlay, or capability-plan entries reused.     |

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

Interactive index counters are scoped to the immutable document revision and
are intended to expose cold versus warm request behavior. For one unchanged
revision, `interactive_index_builds` and `procedure_catalog_builds` should each
increase at most once; subsequent hover, definition, completion, and signature
help requests should report `interactive_index_hits` and
`procedure_catalog_reuses` rather than rebuilding the same source declarations.
Exact, qualified, and prefix requests are counted separately. An ordinary
prefix completion should not increment `interactive_full_symbol_fallbacks`, and
normal hover or definition should not require a `full_document_symbol_builds`
increment. Explicit full document-symbol requests may still build a complete
symbol model. These counters are structural observations only and are never
copied into LSP responses, normal CLI JSON, corpus snapshots, or review ledgers.

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

Issue #756 interactive declaration-index benchmarks report separate cold and
warm hover, definition, and prefix-completion requests for an ordinary module
and the generated 1,200-procedure module:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/lspserver -run '^$' -bench '^BenchmarkLSPInteractiveIndex(Ordinary|Issue756)$' -benchmem -benchtime=1x -count 1 -timeout 20m
```

Issue #758 startup comparisons should use the same fixture and permit budget
with the target module placed after unrelated files, plus one generated giant
unrelated module. Record P0 declaration readiness, P1 helper readiness, the
first successful cross-file definition, and complete declaration readiness
separately. Compare the priority queue with a deterministic discovery-order
run; these measurements are diagnostic evidence, not a CI latency threshold.
The repository benchmark reproduces this comparison with
`BenchmarkLSPDeclarationPriority`:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\dev\go.ps1 test ./internal/lspserver -run '^$' -bench '^BenchmarkLSPDeclarationPriority$' -benchmem -benchtime=1x -count 1 -timeout 10m
```

The declaration-priority benchmark injects a parser stub with a 2 ms wait for
the giant file and fixes the worker count at one. It isolates queue ordering;
it does not measure real VBA parsing, automatic code actions, transport waits,
or VS Code activation. Its ratios must not be reported as editor speedups.

Its `active_declaration_ready_ms` and `helper_declaration_ready_ms` metrics
are reported independently from `workspace_declaration_ready_ms`. The
existing large-class lifecycle benchmark covers the giant-active path, where
an active overlay is promoted before startup work can restore saved
declarations.

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
