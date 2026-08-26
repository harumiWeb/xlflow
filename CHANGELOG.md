# Changelog

All notable changes to xlflow will be documented in this file.

## Unreleased

- Reduced static-analysis allocation pressure for giant VBA modules by sharing
  immutable module array-variable catalogs, lazily copying branch-local
  Dictionary/Collection, Application-state, and generic data-flow maps, and
  resolving VBA205 project procedure shadowing on demand. Added recovered and
  unknown synthetic benchmark workloads and a ten-sample ROneCOne benchmark
  task; diagnostic identity, ordering, JSON/LSP, batch/realtime, cancellation,
  and process-local cache boundaries remain unchanged.

## v0.31.0

- Updated the tree-sitter-vba parser dependency to v0.12.2. Exported VBA
  procedure attributes with multi-segment names now parse without recovery,
  preserving downstream lint and analysis behavior.

- Reduced revision-scoped semantic query overhead by sharing prepared,
  kernel-specific fingerprints across procedure lanes, using comparable cache
  keys and exact document/procedure invalidation indexes, and separating cold,
  warm, local-edit, and dependency-edit corpus benchmarks. Diagnostic, JSON,
  LSP, cancellation, and process-local cache contracts remain unchanged.

- Added process-local, revision-scoped semantic query reuse for unchanged
  analyzer work. Dependency-aware invalidation, cancellation safety, and
  stderr-only query telemetry preserve existing batch/LSP diagnostics; disk
  and cross-process caching remain out of scope.

- Replaced copied derived CFGs with immutable, revision-scoped `CFGView`
  values. Stable block ordinals, canonical filter identities, shared
  adjacency/query indexes, reachability, and dominators are reused across
  equivalent views with bounded procedure/revision lifetime and safe
  concurrent reads. Existing Graph compatibility and defensive ownership
  boundaries remain available; Array, Object, Error, and return-path
  diagnostics, CLI JSON, and LSP contracts are unchanged.

- Completed the HTTP portion of the indexed semantic-state solver migration.
  HTTP transport facts now use revision-local scalar slots with in-place joins
  and changed-state worklist updates; the fixed-point path no longer clones or
  joins a nested `httpAnalysisState`. Deterministic post-convergence evidence
  replay preserves recognized-client, TLS, credential, timeout,
  module-constant, and download/execute diagnostics, including VBA224
  duplicate suppression, batch/realtime ordering, CLI JSON, and LSP contracts.
  Array source-line/edge-refinement paths remain follow-up migrations.

- Extended the indexed semantic-state solver to advanced Array CFG paths,
  including source-line lifecycle ordering, normal-edge refinement,
  exceptional/uncertain edge handling, and combined runtime/evidence lanes.
  Indexed cursors keep semantic slots separate from diagnostic sidecars while
  preserving `VBA208`, `VBA227`, and `VBA249` findings, evidence, ordering, and
  conservative `ReliableExceptional` behavior. A preflight compatibility
  boundary retains the legacy walker for unsupported participant or transfer
  contracts; compact/legacy counts and fallback reasons remain
  developer-only telemetry. The five-sample ROneCOne measurement and focused
  CPU/heap alloc-space profile are recorded in the corpus specification; the
  retained leaf profile shows a material reduction versus the #713 baseline.

- Restricted array interprocedural fixed-point planning to semantic
  participants discovered from array operations, parameters/returns,
  ByRef/module-state dependencies, and resolved caller/callee closure.
  Deterministic worklist requeues now avoid rescanning unchanged participants,
  while unresolved, dynamic, recovered, and incomplete boundaries fail open
  within the smallest known dependency boundary. Added developer-only
  `array_participant_procedures`, `array_interprocedural_cfg_walks`, and
  `array_worklist_revisits` telemetry; diagnostic, snapshot, CLI JSON, and LSP
  contracts remain unchanged.

- Precomputed immutable module semantic facts once per module/revision,
  including conservative `Option Private Module` state and the compact array
  operation observations used by idempotent setup analysis. Procedure workers
  now share the facts safely, avoiding repeated module-wide source scans and
  normalization while preserving diagnostic, snapshot, CLI JSON, and LSP
  contracts. Extended the opt-in fact-build telemetry with
  `module_option_scans`; cross-revision caching and a complete normalized AST
  remain out of scope.

- Split procedure data-flow execution into lazy generic and HTTP lanes, so
  procedures with no applicable HTTP projections avoid HTTP CFG/state work
  while preserving diagnostic ownership, ordering, suppression, and output
  contracts. Added lane-level planned/skipped, kernel, and CFG telemetry to
  the opt-in performance log; normal CLI JSON and LSP payloads are unchanged.

- Reduced giant-module analyzer allocation pressure by representing procedure
  projections as immutable views over canonical Procedure IR. Facts now retain
  compact indexes and read-only access paths instead of copied declaration,
  statement, expression, call, access, and parameter collections. Hot rule
  consumers now iterate procedure and statement-grouped member facts without
  defensive copies, while explicit owned-copy compatibility boundaries remain.
  Diagnostic and JSON/LSP contracts remain unchanged.

- Added a read-only procedure-resolution overlay for batch `VB052`-`VB054`
  diagnostics, avoiding a second full `DocumentIR` clone while preserving
  materialized `Resolve` compatibility and deterministic resolution results.

- Added immutable tri-state procedure feature summaries and applicability-based
  planning for expensive semantic analyzer domains. Proven-irrelevant domains
  can be skipped before setup, while recovered, unresolved, dynamic, or
  otherwise unknown applicability remains planned conservatively. Opt-in
  `--performance-log` output adds planned/skipped domain counters; findings,
  exit codes, and normal JSON/LSP contracts remain unchanged.

- Consolidated compatible array diagnostics onto one immutable,
  procedure-local array semantic result per analysis revision. `VBA227`,
  deterministic array `VBA249`, `VBA208`, object-array `VBA101`/`VBA102`,
  `VBA241`, and applicable `VBA226` projections reuse shared preparation and
  fixed-point facts while retaining their existing diagnostic ownership and
  conservatism. Performance logs expose `array_kernel_runs`,
  `array_cfg_walks`, and `array_projection_runs`; findings, snapshots, exit
  codes, and normal JSON/LSP contracts remain unchanged.

- Added dependency-driven project semantic capability planning. Enabled
  diagnostics now retain only the transitive capabilities they require, with
  shared resolution/effect/index state constructed at most once per analysis
  revision and irrelevant project setup omitted. Batch and Full LSP analysis
  share the same conservative dependency and participant rules; recovered,
  ambiguous, dynamic, or incomplete evidence fails open, and
  compile-equivalent diagnostics remain unconditional. Capability build counts
  and optional elapsed stages are developer-only stderr telemetry; normal
  findings, snapshots, exit codes, and CLI/LSP payloads are unchanged.

- Added explicit procedure semantic execution plans. The planner now selects
  applicable kernels, projections, and semantic-result dependencies in a
  canonical deterministic order, while immutable kernel results are reused only
  within one procedure/revision. Existing bounded procedure workers and
  cancellation remain in force; rule-level or nested kernel pools and
  persistent cross-run caching are out of scope. `--performance-log` adds
  `analysis_plans`, `planned_kernel_runs`, `skipped_kernel_runs`, and
  `semantic_results_reused` as stderr-only telemetry; findings, snapshots, exit
  codes, and normal JSON/LSP contracts remain unchanged.

- Added opt-in semantic-domain profiling beneath
  `procedure_local_diagnostics` for giant-module `xlflow analyze` workloads.
  `--performance-log` now attributes aggregate source-scan, runtime, array,
  object, dictionary, error, dataflow, resource, Excel, application-state,
  and other procedure-local work, and reports candidate/traversal/kernel work
  counters. Parallel domain timings are cumulative worker time, not an additive
  wall-time partition. Output remains stderr-only and findings, exit codes, and
  JSON output are unchanged when profiling is enabled.

- Added bounded procedure-level parallelism for large single-file
  `xlflow analyze` workloads. Procedure batches share the analyzer's bounded
  execution budget with file workers, merge into source-ordered result slots,
  and preserve deterministic findings, suppression, JSON output, and
  cancellation behavior. Ordinary small files retain the existing fast path.

- Fixed `VB060` false positives for writable member assignments such as
  Excel `Range.Hidden`, including member access with an omitted `With`
  receiver. Incomplete member names are no longer treated as proven constants.

- Reused immutable file and procedure analysis facts across analyzer rule
  families. Shared declaration/constant/procedure indexes, statement and
  expression lookups, member-expression projections, and constant overlays
  avoid repeated preparation for large single-module analyses; diagnostic IDs,
  ranges, severities, and public output remain unchanged.

- Improved project-local call graph construction by indexing canonical
  procedure candidate identities, avoiding a full procedure-node scan for each
  uniquely matched call while preserving conservative resolution behavior.
- Bounded `effects.Build` fixed-point propagation for large call graphs by
  separating compact semantic state from deterministic diagnostic witnesses,
  retaining the existing effect/error/uncertainty provenance contract while
  avoiding redundant transitive growth. Added effect-summary worklist and
  propagated-fact performance counters, indexed reachable CFG queries, and
  dedicated large-callgraph benchmark guidance for before/after measurements.
  On the recorded Windows host, provenance-heavy effect workloads reduced
  median wall time by 74.7–88.5% and `B/op` by 64.2–67.5%; ROneCOne's
  `effect_summaries` stage was approximately 0.3% of total analysis time and
  was not a confirmed end-to-end hotspot.

- Added canonical large-single-module analyzer performance telemetry for
  `xlflow analyze --performance-log`, including aggregate workload counters and
  maximum file/procedure dimensions. Profiling remains opt-in and stderr-only;
  analyzer findings, exit codes, and `--json` output remain compatible. Added
  developer-only synthetic and ROneCOne benchmark guidance.

- Fixed false positives in the shared `VBA227`/`VBA249` array-use scan when a
  qualified member call such as `Application.OnTime(...)` or
  `driver_.TableToArray(...)` shares its name with a local scalar or
  array-returning procedure. Member names are no longer mistaken for local
  array indexing.

- Improved large-project `VBA227` analysis throughput by materializing each
  file's Procedure IR and module declaration projections once and reusing the
  immutable inputs across array fixed-point solvers.

- Reduced `VBA227` false positives after line-continuation `Split` assignments
  inside conditional and `With` blocks. Recognized array-factory assignments
  now establish allocation before subsequent bounds and indexed accesses.

- Reduced `VBA227` false positives for valid colon-separated dynamic-array
  declarations such as `Dim values(): ReDim values(...)`. The dedicated
  lifecycle pass now carries that allocation into subsequent array operations
  while preserving fixed-array `ReDim` diagnostics.

- Reduced `VBA227` false positives when an allocated project-local array return
  is passed directly to a private `ByRef` array helper, including nested calls
  on one source line. Unknown and conditionally allocated returns remain
  conservative.

- Reduced `VBA227` false positives in sorted collection helpers. Sorted map
  and sorted set kind branches now carry generic collection array allocation
  into their private `ByRef` helpers.

- Reduced `VBA227` false positives in class collection helpers. Rejecting
  guards that validate a configured collection role or kind now carry the
  corresponding generic collection array allocation into private `ByRef`
  helpers, while unguarded collection access remains reported.

- Reduced `VBA227` false positives in configured class storage members. Internal
  collection, data-row, and aggregate-error accessors/mutators now retain the
  owning instance's proven array allocation when they call private `ByRef`
  helpers.

- Reduced `VBA227` false positives for private, idempotent module-array setup
  helpers. A module-scoped Boolean ready guard followed by plain `ReDim` and a
  final `True` assignment now carries its allocation into downstream private
  helpers, while externally writable guards and resettable arrays remain
  conservative.

- Reduced `VBA227` false positives when a private `ByRef` array-output helper
  allocates a caller-local array that is then passed to another private helper.
  Proven output allocation now follows local array arguments without widening
  module-array propagation.

- Reduced `VBA227` false positives when a module-level array is allocated by a
  public entry or class initializer and then read through a same-module
  private helper chain. The entry-state proof requires every resolved caller
  to establish the allocation and remains conservative for cross-module or
  unresolved calls.

- Reduced `VBA227` false positives in Variant-to-Byte-array `ByRef` adapters.
  `(vbArray Or vbByte)` guards, binary-stream `Read(-1)` results, and the
  `vbNullString` empty-array idiom now flow through normal CFG exits while
  project-local non-returning error helpers are excluded from normal state;
  element access on a known empty Byte array remains diagnosed.

- Reduced `VBA227` false positives when a private helper returns a dynamic
  array and its successful length through paired `ByRef` outputs. Positive
  length guards now carry the array allocation proof into downstream helpers,
  while calls whose guarded array use is unreachable under literal `False` or
  zero arguments no longer poison that proof.

- Reduced `VBA227` false positives for invocation adapters that allocate a
  `ByRef` output array only when a collection has a positive count. The proof
  now preserves that conditional state through positive-count branches and
  numeric `Select Case` dispatches, including nested project-local helpers.

- Fixed Procedure IR argument slots for explicit VBA line continuations so
  interprocedural `VBA202` object-state and `VBA227` array-allocation proofs
  follow the actual arguments.

- Reduced `VBA227` false positives when a project-local array-length helper
  returns `UBound(values) - LBound(values) + 1` on success and zero from its
  error-recovery path, including a typed VBA function that falls through from
  its recovery label to the implicit zero return. The same proof now recognizes
  `UBound(values) + 1` helpers and Variant parameters, so positive-length
  branches establish array allocation, including when the positive result is
  first stored in a scalar local and compared later. The helper's handled probe
  and invalid allocations remain diagnosed conservatively.

- Reduced `VBA227` false positives for whole-array `arr() = ...` assignments,
  `ParamArray` allocation, qualified/nested array factory calls, and proven
  allocated arrays passed through unique project-local `ByRef` helpers.

- Reduced `VBA227` false positives for unique project-local array-return helper
  chains. Chains are now resolved independent of declaration order; public and
  ambiguous calls remain conservative. The same restricted proof carries
  allocated arrays through nested project-local `ByRef` helper chains.

- Reduced `VBA227` false positives by excluding definitely failing constant
  `ReDim` branches from normal array return summaries when no local error
  handler can recover them. The corrected whole-array state also removes the
  corresponding deterministic `VBA249` duplicate runtime findings.

- Fixed lifecycle ordering within a multiline CFG block: `VBA227` now evaluates
  operations in source order, while `VBA208` and `VBA249` retain their existing
  CFG-block semantics.

- Reduced `VBA227` false positives for deterministic plain `ReDim` and
  recognized array-factory assignments on `On Error Resume Next` exceptional
  edges. Whole-array arguments such as `ConvertToJson(values(), 4)` are no
  longer treated as indexed accesses at the call site; `IsArray(variant)` true
  branches now establish allocation for guarded whole-array assignments; and
  `ReDim Preserve` remains conservative when its prior allocation or shape is
  unknown.

- Reduced `VBA227` false positives when module-level arrays initialized by a
  dominating private setup call, including form/class initialization, are
  passed through private `ByRef` helpers. Conditional setup calls remain
  conservative.

- Reduced `VBA227` false positives for growable Byte buffers that use a guarded
  `UBound` capacity probe under `On Error Resume Next` followed by conditional
  `ReDim Preserve` and a bounded write loop. Unrelated `Resume Next` probes
  remain conservative.

- Reduced `VBA227` false positives for class-level dynamic arrays configured by
  project-local `Friend`/`Private` helpers: proven `ByRef` `ReDim` effects now
  flow through matching rejecting role guards and validated role branches.

- Reduced `VBA227` false positives for guarded non-empty `String` assignments
  to dynamic `Byte` arrays, while keeping unguarded empty-string paths
  conservative.

- Reduced `VBA227` false positives for private recursive `ByRef` array helpers
  when an allocated external entry is proven.
- Reduced repeated canonical CFG `IsReachable` query cost by reading cached
  reachability membership directly instead of cloning the full reachable set;
  default and `NormalOnly` semantics are unchanged.

## v0.30.2

- Fixed LSP `VB029` false positives for declared UDT member receivers and VBA
  built-in functions used in member chains such as `GetObject(...).ExecQuery(...)`.

- Added bounded parallel execution for independent per-file batch analysis
  stages. Findings remain deterministic, cancellation is preserved, and large
  synthetic projects show substantial wall-clock reductions.

- Improved repeated CFG query performance by reusing immutable statement,
  adjacency, and reachability indexes for analyzer graph queries. CFG semantics
  and diagnostic output are unchanged.

- Improved batch ByRef analysis by computing each file revision once and
  reusing the typed diagnostics for both runtime-safety and compile-equivalent
  projections. Diagnostic output and the `VBA206`/`VBA228` contracts are
  unchanged.

- Improved `VBA202` batch analysis scalability by gating object analysis behind
  the rule, reusing per-procedure CFG/object-flow artifacts, and propagating
  interprocedural summaries and entry states through dependency worklists.
  Diagnostic output and conservative propagation semantics are unchanged.

- Fixed `VB050` false positives for `Friend` procedures in worksheet and
  `ThisWorkbook` document modules. `Friend` remains rejected in standard
  modules, matching Excel/VBE behavior and issue #656.
- Added opt-in `xlflow analyze --performance-log` batch-analyzer profiling.
  Stage timings, result counts, outcomes, and stable workload counters are
  written to stderr without changing analyzer findings, exit codes, or the
  `--json` stdout envelope. Added deterministic 100/500/1000-procedure scaling
  and real-world corpus benchmark guidance for comparing stage costs and
  allocations.
- Improved batch ByRef call resolution by reusing parsed project symbols and
  indexing exact and qualified-name lookups, avoiding a full project-symbol
  scan for every call site while preserving ambiguity and visibility behavior.
- Reused batch-built procedure IR and CFG artifacts in Intel diagnostics to
  avoid rebuilding the same immutable file revision during `xlflow analyze`.

## v0.30.1

- Fixed `VB053` false positives for valid unqualified Excel/TypeLib enum
  constants such as `xlCenter`, `xlThin`, and `xlContinuous` when metadata
  contains the same name under multiple enum definitions.

- Fixed `VBA225` false positives when a local, parameter, or module-level VBA
  binding such as `cells` shadows Excel's unqualified `Cells` function. The
  Procedure IR now gates the textual fallback, including propagated helper
  summaries, while explicitly typed `Range` and `Worksheet` bindings remain
  eligible.

## v0.30.0

- Completed the `VB012` procedure-terminator compatibility audit across the
  `Sub`, `Function`, `Property Get`, `Property Let`, and `Property Set` opener
  matrix in standard, class, `ThisWorkbook`, and worksheet document modules.
  Excel 16.0 build 17932 accepts `Property Get` with `End Sub` or
  `End Function`; those accepted forms now use non-blocking, suppressible `VB066`
  style warnings, while VBE-rejected mismatches remain `VB012` compile errors
  that block source preflight. Parser structure, VBE validity, and style policy
  are documented and covered separately.
- Fixed `VBA214` false positives when a narrow `On Error Resume Next`
  compatibility probe captures `Err.Description` before clearing the error and
  restoring normal error handling.
- Added a shared, Excel-free VBA constant-expression evaluator for issue #595.
  Typed `Known` / `Unknown` / `Invalid` results now power Optional defaults,
  Const and Enum references, fixed-array bounds, and `ReDim` bounds across
  batch and LSP analysis. Runtime calls, ambiguous or unresolved values, and
  safely unmodelled overflow remain fail-open instead of becoming diagnostics.

- Added default-enabled `VBA249` deterministic runtime-error diagnostics for
  proven division, numeric-operand, conversion, and array allocation/bound
  failures derived from shared constant, type, CFG, and dataflow facts.
  `runtime-error` findings use `error` severity but remain inline-suppressible
  and non-preflight-blocking so they stay distinct from VBE compile-equivalent
  errors; unknown `Variant`, late-bound, external, locale-dependent, and
  branch-merged values remain silent.

- Fixed `VBA204` false positives for project-specific labels used as shared
  cleanup or loop-finalization paths, while retaining findings for normal
  fallthrough into logging, error reporting, or arbitrary handler code. Cleanup
  recognition now uses parsed assignments and exact IR call members, so string
  contents and longer cleanup-like identifiers cannot suppress a finding.

- Added default-enabled, compile-equivalent `VB062`–`VB065` syntax
  diagnostics for provably invalid conditional branches, `Select Case`
  branch ordering, malformed `Open ... For <mode>` shapes, and malformed
  `TypeOf` expressions. Lint, LSP, and source preflight share unsuppressible,
  preflight-blocking errors; ambiguous parser recovery remains `VB014`.

- Reduced `VBA202` false positives for private object helpers, successful Excel
  member/factory assignments, read-only `ByRef` aliases, UserForm controls,
  and error-handler cleanup paths. Public or unresolved boundaries,
  `On Error Resume Next`, and nullable object results remain conservative.
- Fixed `VBA222` false positives for unresolved external public API types when
  the project or generated TypeLib resolution view is incomplete. Project-local
  inaccessible and ambiguous type findings remain unchanged.

- Fixed the bundled `XlflowDebug` scaffold helper so `new` projects and helper
  modules installed by `init --with-module` produce no default `xlflow analyze`
  findings. ParamArray array-bound checks use narrow `VBA227` suppressions, and
  optional debug-pipe write/close failures now return checked best-effort results.

- Expanded static array-operation validation for issue #593. Shared array
  shapes now cover fixed and dynamic arrays, multidimensional declarations,
  `Variant` values, array-returning procedures, `Array(...)`, `ParamArray`, and
  `ByRef` flows. The analyzer can validate statically provable `ReDim`,
  `Erase`, `LBound` / `UBound`, and `For Each` source misuse while preserving
  `VBA208` ownership for `ReDim Preserve` and conservative handling for
  unknown `Variant` state. Compile-equivalent declaration and constant
  assignment cases remain separate from runtime-safety warnings.

- Added default-enabled, compile-equivalent `VB059` call-syntax diagnostics.
  Lint, LSP, and source preflight now reject explicit `Call` statements without
  argument parentheses, standalone empty or multi-argument parenthesized calls,
  Function calls used in expressions without required parentheses, and
  syntactically identifiable invalid `Call` targets. Legal parenthesized
  ByRef/ByVal idioms remain separate as the existing `VB022` maintainability
  warning.

- Added opt-in `VBA248` maintainability analysis for opaque positional Boolean
  control arguments. Multiple positional `True`/`False` literals are reported
  with structured call-site context, while named arguments suppress the
  single-literal heuristic and conventional single flags remain clean.
  `xlflow metrics` now exposes additive Boolean-parameter and direct-control
  branching measurements without adding a declaration-level diagnostic.

- Added default-enabled, compile-equivalent `VB055`–`VB058` diagnostics for
  duplicate and undefined procedure labels, mismatched `Next` control
  variables, and invalid loop/procedure `Exit` statements. Lint, analyze,
  preflight, and LSP surfaces share conservative procedure IR/CFG facts;
  recovery and conditional-compilation uncertainty fails open, and the
  errors remain preflight-blocking and unavailable to inline suppression.

- Fixed `VBA201` to require a receiver resolved as `Excel.Range` before
  tracking nullable `Range.Find` results, avoiding false positives for
  project-defined `.Find` methods.

- Fixed `VBA209` to ignore `Set object = Nothing` assignments, including inline
  conditional assignments, while retaining diagnostics for scalar `= Nothing`
  object comparisons.

- Added deterministic procedure and module hotspot ranking to `xlflow metrics`
  (`metrics.hotspots`, schema version 1, score model
  `percentile_equal_weight_v1`). The report exposes raw and normalized
  architectural signals, supports independent top-N and percentage selectors
  through `[metrics.hotspots]`, and emits metrics-owned `MX002` entries (top-N
  informational, threshold-selected warning) while retaining the complete
  metrics payload. Hotspot scores are maintainability review signals rather
  than definite analyzer findings.

- Added default-enabled `VBA246` HTTP transport-security analysis and `VBA247`
  timeout reliability analysis for common XMLHTTP, ServerXMLHTTP, WinHTTP, and
  identifiable ADODB.Stream download-and-launch patterns. Findings are
  procedure-local, realtime, warning-level, non-blocking, inline-suppressible,
  and expose redacted `http_security` / `http_reliability` context. Added exact
  `[analyze].development_http_origins` exceptions for plain-HTTP credential
  development flows while retaining `VBA224` as the generic fallback.

- Added project-level `[preflight].allowed_diagnostics` waivers for every
  registry-owned source-preflight blocker. On each supported lint, analyze, or
  LSP surface, waived diagnostics retain their error severity and suppression
  behavior, while workbook commands proceed with deterministic
  occurrence-count warnings; parser, infrastructure, duplicate-component, and
  UserForm integrity failures remain non-waivable.

- Added default-enabled, compile-equivalent `VB052`–`VB054` diagnostics for
  provably invalid project-local call targets, ambiguous bare Enum members,
  and undeclared `RaiseEvent` targets. Batch lint/analyze and LSP Full share
  the canonical procedure resolver and fail open for external, built-in,
  late-bound, dynamic, conditional, partial, or otherwise incomplete models;
  all three errors are unsuppressible and block source preflight.

- Added default-enabled `VBA245` file/path safety analysis for destructive VBA
  statements, FileSystemObject write/delete operations, workbook SaveAs paths,
  wildcard and traversal hazards, unchecked overwrites, external-input-derived
  paths, and missing local temporary-file cleanup. Findings include an additive
  `file_operation` context and retain `VBA224` as the compatibility fallback
  when the specialized rule is disabled.

- Added source-only `xlflow metrics` procedure complexity reporting with a
  versioned JSON schema (`metrics.schema_version = 1`) and deterministic values
  for cyclomatic complexity, nesting, statements, source lines, branches,
  loops, `GoTo`, exits, parameters, `ByRef` parameters, locals, and call
  fan-out. Optional `[metrics.thresholds]` settings emit metrics-specific
  `MX001` warnings and exit `1` while retaining the complete metrics payload;
  exclusions are configured independently through `[metrics].exclude`.

- Added default-enabled, compile-equivalent `VB050` module declaration-context
  diagnostics and `VB051` invalid-`Me` diagnostics. Checks use canonical
  standard/class/document/UserForm metadata across lint, LSP, and source
  preflight, keep unknown type-dependent `WithEvents` cases fail-open, and
  remain separate from style rule `VB006`. Extended the developer-only VBE
  oracle schema v1 with class, UserForm, and document module probes while
  retaining `.bas` fixture compatibility and cleanup-gated promotion.

- Added default-enabled, batch-only `VBA244` analysis for recursive and cyclic
  project-local procedure dependencies. Normal analysis now reports one
  deterministic representative `call_cycle` witness per cyclic SCC, retaining
  severity, event-handler, dangerous-effect, cross-module, and uncertainty
  context aggregated across the component. Explicit graph inspection may still
  expose exhaustive elementary cycles; this intentional finding-multiplicity
  change prevents dense call graphs from causing elementary-cycle explosion.
- Fixed `VB023` false positives for valid composite `For Each` control targets
  such as array elements. The rule now resolves bare control identifiers from
  Procedure IR declarations and remains conservative for unresolved composite
  targets.
- Fixed `VB004` false positives for narrow `On Error Resume Next` probes that
  restore handling on the same physical line, after a colon-separated
  statement, or by switching to an explicit error handler. Procedure exits
  without restoration remain covered by the rule.
- Fixed `VB010` false positives for mutually exclusive conditional-compilation
  procedure declarations that share a common terminator after `#End If`.
- Fixed the developer-only VBE oracle cleanup gate so unrelated Excel
  instances started concurrently remain untouched and observable without
  invalidating cleanup of the oracle-owned process.
- Fixed `VB009` false positives for valid, closed VBA string literals where a
  literal backslash is adjacent to doubled quote characters.

- Added opt-in `VBA242` performance analysis for expensive operations over
  entire rows, columns, worksheets, or unbounded `UsedRange` expressions.
  Findings use `information` outside loops and `warning` in reachable loops,
  accept explicit bounded ranges, and recommend deriving bounded last-row or
  last-column limits. Configure it with
  `detect_expensive_full_range_operations` or disable it with
  `[analyze].disabled_rules = ["VBA242"]`.

- Added opt-in `VBA243` performance analysis for bulk or repeated
  `Range.Value` transfers where `Range.Value2` may avoid unnecessary
  Date/Currency coercion. Findings use `information` outside loops and
  `warning` for reachable repeated transfers, remain conservative around
  intentional Date/Currency handling, and can be configured with
  `detect_value2_performance_opportunities` or disabled with
  `[analyze].disabled_rules = ["VBA243"]`.

- Added default-enabled `VBA241` performance analysis for reachable
  `ReDim Preserve` statements inside VBA loops. The rule covers all supported
  loop forms, distinguishes loop-variable-dependent growth from repeated
  constant-size reallocation, raises nested-loop findings to `warning`, and
  recommends preallocation or geometric capacity growth. Configure it with
  `detect_redim_preserve_in_loops` or disable it with
  `[analyze].disabled_rules = ["VBA241"]`.

- Added default-enabled, compile-equivalent `VB046` duplicate-declaration and
  `VB047` invalid-declaration-placement lint diagnostics. Both are high-precision
  errors emitted by batch lint and the LSP, block source preflight, and cannot be
  disabled or suppressed inline.

- Added default-enabled, compile-equivalent `VB048` procedure-parameter and
  `VB049` Property-accessor contract lint diagnostics. The checks validate VBA
  parameter declaration constraints and same-name Get/Let/Set signatures in
  batch lint and the LSP, fail open only where type resolution is unknown or
  ambiguous, and block source preflight without allowing inline suppression.
  Optional-default validation recognizes VBA decimal, hexadecimal, octal, and
  numeric type-suffix literals while leaving String-suffix references unknown.

- Fixed `VB022` false positives on parenthesized `ElseIf` and conditional-
  compilation `#If` expressions by recognizing only whitespace-separated
  parenthesized procedure calls.

- Added opt-in `VBA240` project-wide analysis for module-level mutable state.
  The analyzer builds per-procedure read/write sets, aggregates field readers
  and writers across uniquely resolved call paths, reports structural lifecycle
  coupling, and exposes informational `analysis_metrics.module_state` output.
  Constants and read-only configuration are classified separately, and fan-in
  counts alone do not produce findings. Collection classification is independent
  of procedure order, built-in type matching avoids user-defined name substrings,
  and procedure metrics are emitted deterministically.
- Fixed `VBA240` false positives where multiple confirmed roots only read a
  field while its writer was unreachable from every confirmed root. Findings
  now require an observed reachable write or mutation before reporting
  cross-root lifecycle coupling.
- Updated `tree-sitter-vba` to v0.12.0. VBA declaration type characters
  (`$`, `%`, `&`, `!`, `#`, `@`, `^`) now count as explicit types for `VB005`
  and `VB019`; invalid `Dim`/`ReDim` keywords after a declaration comma are
  reported as `VB014` without synthesizing declarator diagnostics. Formatter
  recovery leaves these invalid declarations unchanged.

- Fixed `VB020` false positives for procedure-local variables and constants
  declared with VBA type-declaration characters. The declaration and reference
  sides now use the same identifier identity while write-only locals remain
  reportable.

- Added default-enabled `VBA239` SQL-construction safety analysis for ADO and
  DAO execution boundaries. Findings track SQL value/identifier/LIKE roles,
  locale-sensitive formatting, manual quoting, and parameterized
  `ADODB.Command` usage without claiming complete SQL-injection proof.
- Strengthened `VBA202` object use-before-`Set` analysis with CFG-based,
  branch- and interprocedural tracking across locals, parameters, module and
  `Static` state, explicit `Nothing` resets, `ByRef` mutations,
  `As New`/constructor initialization, exceptional paths, and nullable
  function/callee returns. Unresolved calls no longer imply initialization.

- Fixed `VB014` false positives for valid date-literal and named-argument colons,
  UDT fields named `Next`, and repeated conditional-compilation guards.
- Fixed `VB029` false positives for project-visible public assignments, intrinsic
  `Err` assignments, bang-member targets, and function names with VBA type
  declaration suffixes.
- Fixed `VB045` false positives caused by UDT fields, local shadowing,
  conditional project overloads, and incomplete `Application.Run` signatures.
  Argument binding now fails closed unless the call target and signature are
  unambiguous; batch analysis also records the rule's projected `analyze`
  surface for corpus and preflight consistency.

- Added default-enabled `VBA236` command-execution safety analysis for VBA
  `Shell`, `WScript.Shell.Run`/`Exec`, `ShellExecute`, `cmd.exe`, PowerShell,
  and script-host launches. Findings distinguish potential injection from
  general process-launch risk and expose launcher, interpreter, command-role,
  origin, credential, observability, and unquoted-path context with Windows
  quoting guidance. Process-launch ownership moves from generic `VBA224` to
  the new configurable `detect_unsafe_command_construction` rule.
- Added default-enabled `VBA237` warnings for handlers, cleanup paths,
  `On Error Resume Next` regions, and ignored Boolean success results that lose
  failure information across uniquely resolved local call chains. Batch and
  LSP Full diagnostics share provenance-preserving procedure summaries and
  report the location where the failure signal is lost.
- Added default-enabled `VBA238` performance warnings for loop-invariant Excel
  object-model member chains. Batch and real-time analysis normalize equivalent
  constant lookups, ignore iterator-dependent expressions, and suggest hoisting
  the invariant chain into a cached local. Configure it with
  `detect_loop_invariant_excel_object_resolution` or disable it with
  `[analyze].disabled_rules = ["VBA238"]`.
- Fix false-positive `VBA205` findings for local, parameter, and project symbols that shadow Excel root names, including procedure declaration and `Attribute` lines.

- Expanded `VBA212` to analyze AST operand relationships across eager `And`,
  `Or`, `IIf`, `Choose`, and `Switch` expressions, including nested
  member/index access, `IsArray` bounds operations, and uniquely resolved
  side-effecting property getters. Diagnostics now identify the guarded object
  and unsafe operand range precisely.

- Added high-confidence `Scripting.Dictionary` and `Collection` safety
  diagnostics for `CompareMode` ordering, repeated key/item materialization,
  key normalization, late-bound comparison constants, mutation during
  iteration, and one-based indexing. Guard analysis now distinguishes definite
  missing keys from uncertain existence and recognizes safe local probes,
  aliases, construction, and uniquely resolved helper effects.
- Fixed `VBA204` false positives for qualified `_Cleanup` labels and handler
  labels whose only implicit predecessor is a non-returning `Err.Raise` call.
- Fixed `VBA203` false positives when a saved `Application` property value is
  restored after a clean/dirty CFG merge, after `Err.Raise`, or under a repeated
  unchanged guard.
- Added developer-only corpus review commands for exact diagnostic ranges,
  schema-valid ledger drafts, and read-only project/rule-focused verification;
  added an extensible machine-checked rule-contract marker inventory to prevent
  protected specification metadata from drifting from the registry.
- Fixed `VBA219` false positives for `Workbooks.Open` assignments to
  non-Workbook locals or return slots and for direct ownership transfer through
  a Workbook-valued Function return.
- Fixed `VBA202` false positives for `For Each` object iterators and type
  qualifiers that share a name with a local object variable. Module fields and
  persistent `Static` objects are now tracked conservatively and are accepted
  only after a dominating or proven callee initialization.
- Fixed `VBA224` false positives where built-in, module, and local constant
  identifiers were treated as separate unknown inputs flowing to sensitive
  APIs.
- Fixed `VBA226` false positives where ordinary arrays were treated as
  `Range.Value` results, nested `Cells` calls were mistaken for the outer
  `Range` receiver, or a dynamic range was provably multi-cell from one known
  extent.
- Fixed `VBA223` false positives for credential-adjacent URL/resource names,
  qualified assignment receivers, Digest header fragments, nested `With`
  statement rescans, and obvious self-describing test credentials.
- Fixed `VBA208` false positives for one-dimensional `ReDim Preserve`, stable
  symbolic non-final dimensions across CFG joins, and indexed member-array
  targets that were incorrectly attributed to their receiver arrays.
- Fixed `VBA222` false positives for standard OLE Automation interfaces and
  the VBA `VbAppWinStyle` enum, and stopped qualified external types from
  resolving through an unrelated library's same-named type.
- Fixed `VBA228` false positives for literal ByRef temporaries, array-valued
  Function return slots, local value symbols that shadow same-named project
  procedures, and module-qualified project types. Literal temporaries remain advisory
  `VBA206` warnings instead of blocking source preflight.
- Fixed `VBA229` false positives for project UserForm/document types and
  embedded VBA/Excel enum types, and prevented a missing or incomplete generated
  TypeLib manifest from turning an unsupported type-name absence claim into a
  blocking diagnostic.
- Fixed VBA209 false positives for whole-array assignments and arrays nested in
  function, `LBound`, or `UBound` arguments, including incorrect CFG block-start
  source attribution.
- Reduced duplicate VBA220 event-reentry warnings when one resolved call
  carries multiple effect and uncertainty facts; one representative warning is
  retained per statement/call boundary.
- Reduced conservative data-flow analysis allocation pressure by comparing
  deterministic propagation paths without repeatedly serializing them, and
  indexed effect-summary candidate lookup so matching one call no longer
  clones the complete project summary. Added developer-only real-world corpus
  benchmarks and profiling commands for the `std-vba` and `ronecone` analyzer
  hotspots.
- Added context-aware cancellation to procedure-level `VBA224` dataflow
  analysis and propagated explicit cancellation errors through batch and
  real-time analyzer paths.
- Fixed VBA LSP hover for procedures in open documents so `'''` and Rubberduck
  documentation comments are displayed through the lightweight symbol path.
- Defined reviewed static-analysis corpus evidence with exact expected and
  forbidden diagnostic contracts, durable false-positive records, and
  reviewed-only per-rule precision metrics. Added developer tasks for printing
  the metrics and selecting bounded, deterministic unreviewed candidates by
  rule from committed evidence.
- Made `VBA224` finding emission independent of dataflow worklist scheduling by
  projecting diagnostics only after block-entry states reach a fixed point.
- Re-enabled all 16 pinned static-analysis corpus projects. Analyzer type
  inference now terminates on self-referential `Set` assignments without losing
  earlier assignment context, project-effect propagation uses indexed
  membership, and procedure dataflow converges through a deterministic
  reverse-postorder worklist on large nested control-flow graphs.
- Added `VBA229`, an unsuppressible compile-equivalent diagnostic for
  unresolved procedure-local `As <Type>` names. It shares production
  built-in, host/TypeLib, and project-symbol resolution across analyze and LSP.
- Defined VBE-verified compile-equivalent severity and preflight policy. Split
  deterministic argument binding into `VB045`, scalar `Set` errors into `VB037`,
  and definite ByRef type mismatches into `VBA228`; `VB030`, `VBA206`, `VB001`,
  `VBA214`, and `VBA225` remain warning-level inference/runtime-safety/policy
  diagnostics, and compile-equivalent findings cannot be hidden by inline
  suppression.
- Clarified VBE oracle fixture binding/evidence roles and coverage reporting.
  Compile-equivalent bindings now remain distinct from language, policy, and
  maintainability observations; `sub-parenthesized-call` is unbound with a
  maintainability role, `missing-set-object-assignment` records a policy role,
  and `unknown-named-argument` is now fully bound to `VB045`.
- Stabilized the local VBE oracle cleanup path by closing the disposable
  workbook before quitting Excel, allowing delayed owned-process exit within a
  bounded drain, preserving fail-closed behavior for unexpected Excel
  processes, and emitting structured cleanup diagnostics. Oracle batches now
  use a crash-released Windows cross-process lock and reject concurrent local
  runs with `oracle_already_running`.
- Added Excel-free VBE oracle binding coverage validation. Rejected fixtures
  can now declare accepted `negative_controls`; bound rules require rejected
  positive evidence and accepted forbidden coverage across all declared
  analyzer surfaces, while historical unbound fixtures remain visible in a
  deterministic coverage report.
- Extended the VBE oracle binding contract to resolve diagnostic codes against
  the shared static-analysis registry and validate supported surfaces and
  severities. The public `xlflow rules --json` metadata now includes additive
  `surfaces` and `supported_severities` fields.
- Added the local-only Windows Excel/VBE oracle harness for focused static-
  analysis compile evidence, with sequential disposable workbooks, known
  accept/reject controls, observational and strict JSON modes, explicit
  promotion, provenance metadata, and deterministic Excel-free fixture
  contracts. Fixtures now also carry explicit diagnostic binding metadata so
  unbound, partially-bound, bound, and not-applicable evidence can be validated
  independently of later binding-coverage work. The oracle is
  not invoked by production commands or GitHub Actions; contributors should run
  relevant cases locally when changing VBA semantics.
- Added default-enabled `VBA223` security warnings for likely hardcoded VBA
  credentials, with structural matching, placeholder filtering, fixed
  `[REDACTED]` diagnostics, and independent configuration/inline suppression.
- Added default-enabled batch `VBA222` warnings for public APIs that expose
  private, unexposed, ambiguous, or unresolved types. The rule covers
  procedures, properties, custom events, and `VB_Exposed` classes/interfaces,
  while excluding host-required event handlers and remaining non-blocking.
- Added default-enabled `VBA224` conservative procedure-local source-to-sink
  analysis with shared batch/LSP diagnostics, explicit source/sink/path JSON
  context, narrow sanitizer contracts, and non-blocking inline suppression.
- Added default-enabled `VBA226` warnings for unsafe `Range.Value` / `Value2`
  scalar, one-dimensional, dimension/order, bounds, and destination-shape
  assumptions, with conservative procedure-local shape tracking in batch and
  real-time analysis.
- Added default-enabled `VBA227` array lifecycle and dimension-safety warnings
  with CFG allocation/erase tracking, conservative Variant handling, object-array
  missing-`Set` checks, and project-local array return summaries. Existing
  `VBA208`, `VBA209`, and `VBA226` ownership and finding contracts remain stable.
- Improved opt-in `VB021` to analyze project call-graph roots, host events, tests,
  `WithEvents` handlers, UserForm metadata, and dynamic callbacks before
  reporting unreachable private procedures.
- Improved opt-in `VB021` to treat externally callable public standard-module
  APIs as possible roots, preventing false warnings for private helpers in
  scaffolded and user-authored helper libraries.
- Documented the `VBA225` analyzer contract for issue #452: repeated
  cell-by-cell Excel object-model work inside loops, nested-loop context,
  helper-call coverage, small fixed-loop exemptions, and bulk array/range
  remediation guidance.
- Strengthened opt-in `VBA210` return-path analysis to cover `Function` and
  `Property Get` CFG exits, early exits, error-handler paths, shared cleanup,
  dominating assignments, object/value return assignment syntax, and
  non-returning `Err.Raise` error paths.

## v0.29.0

- Fixed `VB004` false positives for bounded `On Error Resume Next` probes that
  check `Err.Number` and restore normal handling, `VB022` false positives for
  intrinsic function argument expressions, and `VB029` false positives for
  multiline comparison arguments such as `vbTextCompare) = 0`.
- Added procedure-scoped incremental LSP diagnostics with 300 ms Fast and
  two-second idle Full publication, request-scoped workspace resolution,
  reusable IR/CFG procedure fragments, indexed source positions, and stage
  performance counters. Large-module Full analysis no longer repeats
  workspace symbol copies or line-prefix scans, while generation, suppression,
  diagnostic range, and close/reopen safety are preserved. Obsolete IR/CFG
  fragment revisions are pruned after successful rebuilds so repeated edits do
  not accumulate unreachable large-module artifacts.
- Fixed Fast LSP procedure diagnostics to retain enclosing conditional-
  compilation branches when a procedure follows an earlier module procedure.
- Fixed new-project helper scaffolds to avoid unused `VB004` inline-suppression
  warnings during `xlflow lint`.
- Fixed `VB014` false positives when valid VBA identifiers such as `nextChar`,
  `nextSlot`, or `NextHashCapacity` begin with the `Next` keyword text.
- Improved LSP responsiveness for very large VBA modules by removing a
  superlinear real-time diagnostic traversal, making document open analysis
  asynchronous, preventing stale workspace overlays from publishing after
  changes or reopen, and propagating cancellation through expensive analysis.
- Updated `tree-sitter-vba` to v0.11.1 so VBE-exported procedure `Attribute`
  statements, including `Load` and `Name` targets, parse without recovery nodes.
- Made `VBA206` a default-enabled, real-time non-blocking warning for runtime-
  safety ByRef forms. Definite VBE-rejected type mismatches are reported by the
  unsuppressible, preflight-blocking `VBA228` compile-equivalent diagnostic;
  `VBA206` continues to cover temporary values, indirect member/array
  expressions, and common `PtrSafe` pointer-width declaration mistakes.
- Updated `golang.org/x/text` to v0.39.0 to include the latest invalid-input
  handling security fix.

## v0.28.0

- Added default-enabled batch `VBA221` warnings for direct calls to local
  helpers that can leave an `Application` property changed. The warning keeps
  `VBA203` at the root assignment, reports only the immediate caller, preserves
  callee uncertainty, and remains non-blocking and locally suppressible.
- Fixed `VBA220` to ignore non-Excel `Select`/`Activate` calls, avoid duplicate cell-write and workbook-structure risks, and recognize safely restored event guards around delegated cell work.
- Fixed `xlflow analyze` performance for `VBA215`, `VBA218`, and `VBA220`: typed Excel-call rules now reuse each batch file's parsed analysis snapshot, and `VBA220` preparation is skipped when the rule is disabled.
- Added default-enabled batch `VBA220` warnings for supported Excel and UserForm event handlers that can re-enter themselves or trigger related event chains. The rule uses uniquely resolved local call effects, reports unresolved calls as uncertainty, recognizes CFG-proven `Application.EnableEvents` cleanup for Excel events, and remains non-blocking and locally suppressible.
- Added default-enabled real-time `VBA219` warnings for locally acquired
  Workbooks and VBA file handles that can reach a normal, error, termination,
  or unknown exit without a recognized Close path. The rule recognizes local
  aliases, cleanup handlers, and object-returning Function ownership transfer;
  it remains non-blocking and locally suppressible.
- Expanded default-enabled `VBA205` warnings to identify ambiguous Excel workbook and worksheet roots: active UI objects, unqualified sheet collections, positional workbook/window access, uncaptured `Workbooks.Open` calls, and `ThisWorkbook` from add-in standard modules. Findings remain non-blocking and can be suppressed locally for intentional interactive macros.
- Added default-enabled real-time `VBA218` warnings for resolved Excel APIs whose exceptional or `Variant/Error` failure contracts are consumed without an appropriate local guard.
- Added default-enabled real-time `VBA216` preflight errors for VBA range expressions that provably mix distinct worksheet roots, plus `VBA217` guidance for implicit-root and unstable last-row calculation patterns.
- Added default-enabled real-time `VBA215` warnings for resolved `Range.Find` and `Range.Replace` calls that omit saved Excel search settings, including support for positional, named, mixed, and multiline calls plus normal rule-specific suppression.
- Fixed scaffolded `XlflowAssert.bas` helpers to restore `On Error` handling immediately after each compatibility probe, so new projects pass the default-enabled `VBA214` analysis without warnings.
- Fixed `VB030` false positives for valid VBA `Array` calls with zero or multiple arguments by modeling its optional argument list as a `ParamArray`.
- Added default-enabled `VBA214` analysis for `On Error Resume Next` scopes that extend beyond one compatibility probe or can exit without restoration. Findings report the scope's start and effective end line; project-local calls remain warning-level and do not block `push` or `run` preflight.
- Strengthened `VBA203` so all paths after a tracked Excel `Application` state change, including early exits and error handlers, must restore the saved prior value. The diagnostic now covers `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, `Calculation`, `StatusBar`, `Cursor`, `Interactive`, `AskToUpdateLinks`, `AutomationSecurity`, and `CutCopyMode`.
- Added a shared, protocol-neutral static-analysis rule registry for `VB...` and `VBA...` diagnostics, plus the source-only `xlflow rules` command, generated rule catalog, LSP documentation links, and registry-driven VS Code inline-suppression eligibility while preserving existing `disabled_rules`, legacy boolean, and inline suppression behavior.
- Updated `tree-sitter-vba` to v0.11.0 and preserved lint and formatter behavior across its flat multiline-`If` CST: VB014 now runs structural block validation even when the parser accepts fragment nodes, and formatting pairs `if_statement`, `elseif_fragment`, `else_fragment`, and `end_if_fragment` ranges.
- Added a conservative procedure-level VBA control-flow graph shared by batch analysis and LSP snapshots, and moved `VBA204` error-handler fallthrough detection from preceding-text heuristics to normal-flow reachability without changing its public diagnostic contract.
- Added deterministic, provenance-bearing VBA procedure effect summaries with fixed-point propagation across uniquely resolved reachable project-local calls, and updated the existing `VBA203` Push/Pop helper exemption to recognize indirect restores without changing public diagnostics, configuration, or JSON output.

## v0.27.1

- Improved `VB014` parser-recovery diagnostics for likely unclosed VBA blocks. When a missing `End If`, `Next`, `Loop`, `Wend`, `End With`, or `End Select` can be identified confidently, lint, LSP, and push/run preflight now report the unmatched opening line and expected closer; a reliably indented parent closer is highlighted for an unclosed nested block, while ambiguous recovery remains generic.
- Fixed `--wait-timeout` so uncontended workbook coordination metadata work no longer consumes the lock-wait budget.
- Updated `tree-sitter-vba` to v0.10.1, adding support for nested colon-separated single-line control flow, shared `Next` counter lists, and complete single-line `If` statement sequences.

## v0.27.0

- Added `xlflow test --visible` to run GUI-bound VBA integration tests, such as `Application.OnKey` registration, with a visible Excel window without changing the project's `[excel].visible` configuration.
- Refactored the bundled `xlflow` Agent Skill into an orchestration-focused proof loop with dedicated testing, formula, UserForm, UI, debugging, recovery, and structural-analysis references. The Skill now also includes an initial `xlflow --help` discovery rule, a purpose-to-command capability map, a session-first executable quick start, explicit recovery stop conditions, and visual-verification guidance through `export-image` while leaving complete CLI syntax and diagnostics to xlflow help and documentation.
- Fixed `xlflow analyze` `VBA209` false positives for bundled `XlflowUI` functions that return array values.
- Fixed .NET bridge temporary VBA module injection for `run`, `test`, and related form/UI helpers when VBE's **Require Variable Declaration** option pre-populates a new module with `Option Explicit`. Generated code now replaces the initial module content, avoiding duplicate declarations and compile failures.
- Added `xlflow impact <Module.Procedure>` for deterministic, AI-friendly VBA call impact analysis. It reports confirmed direct/transitive callers and callees, cycles, affected modules, source locations, and explicit conservative-resolution uncertainty without opening Excel.
- Added standard LSP Call Hierarchy for confidently resolved project-local VBA procedures. Incoming and outgoing calls reuse the incremental workspace index and current unsaved editor overlays; ambiguous, unresolved, external, built-in-like, dynamic-member, and module-level calls remain excluded, and private procedures stay module-scoped.
- Extended LSP `VB030` argument diagnostics to flag known project-local arrays passed to non-array `Object` parameters, including both `ByVal` and `ByRef` calls, before Excel can surface a runtime type mismatch.
- Added `VB014` parser-compatibility guidance to the bundled Agent Skill so agents distinguish tree-sitter recovery from Excel/VBE compile failures, preserve valid idiomatic VBA, and report compatibility limits without bypassing preflight.
- Added a versioned `xlflow build` manifest to JSON output and as a companion `<output>.build.json` artifact. Build now reports resolved included/excluded components and validation evidence; if only companion publication fails, the validated workbook remains successful with `build_manifest_publish_failed`. `capabilities --json` now also reports whether the normal command path requires Excel.
- Hardened `xlflow build` publication: Excel now builds only in a bridge-owned staging directory beside the requested output, verifies the saved staged artifact after Excel exits, and then uses atomic create or replace publication. Existing artifacts remain untouched on pre-publication failure; unsafe output locks, staging/output-directory failures, and non-atomic replacement fail with structured build errors. Staging cleanup failures preserve the published artifact and report a warning plus cleanup metadata.

## v0.26.1

- Fixed VBA source encoding to use the Windows system ANSI code page (`GetACP()`) instead of a hard-coded `932` when reading exported components and preparing components for import. Non-ASCII source (for example German umlauts and em dashes on a `1252` machine) now round-trips through `pull`/`push` without corruption on non-Japanese locales, while Japanese (`932`) environments are unaffected.
- Added the `xlflow build` CLI contract and deterministic `--dry-run` release-plan output. Non-dry builds now use isolated Excel staging with VBE compilation, validated save/close/process cleanup, atomic output replacement, dirty-session blocking, and recovery quarantine for uncertain cleanup.
- Fixed `.NET` temporary macro runner invocation to qualify the generated runner with its workbook name, preventing Excel error 1004 when another workbook or VBA project is active.

## v0.26.0

- Fixed generated and updated project `.gitignore` files to ignore Excel add-ins (`*.xlam`).
- Fixed macro discovery and CodeLens run actions to recognize runnable no-argument private, public, and friend procedures in worksheet and `ThisWorkbook` document modules, while hiding workbook event procedures. `xlflow run` now reports an unavailable injected harness as `runner_not_invocable` instead of a target-macro failure.
- Added context-aware LSP completion for UserForm YAML specifications under `src/forms/specs`. Completion derives structural keys, control properties, fixed values, built-in control types, ProgIDs, and parent IDs from the shared UserForm contract; it also remains available for common incomplete YAML states and includes authoring snippets. Snapshot-oriented fields are offered only when explicitly typed, and JSON completion remains future work.
- Added type-aware LSP diagnostics for UserForm YAML specifications under `src.forms/specs`. The shared UserForm validator now publishes stable `UFV001`–`UFV014` errors and support-level warnings with precise YAML key/value ranges alongside syntax diagnostic `UFY001`; malformed, invalid-kind, and corrected editor buffers update without reopening while unrelated YAML remains ignored. JSON diagnostics, completion, and Hover documentation remain future work.
- Changed `form new` to omit the snapshot-only `warnings` field from blank authoring specifications, and added a UserForm specification reference covering fields, controls, support levels, validation, and diagnostics.

## v0.25.0

- Added opt-in `VB044` lint diagnostics for local `PROCEDURE_NAME`-style constants whose direct string literals drift from their enclosing `Sub`, `Function`, or `Property` name. The `[lint.procedure_name_constant]` configuration supports standard, class, document, and UserForm modules; the LSP supplies a Quick Fix that updates only the string literal, while `xlflow lint` remains read-only.
- Added opt-in `[vba.line_numbers].enabled` instrumentation for meaningful VBA `Erl` diagnostics: `push` injects fixed-width, space-padded physical line numbers only into temporary import copies, while `pull` strips recognized generated labels to keep tracked source unnumbered. Unsafe existing numeric labels and numeric `GoTo` / `GoSub` / `Resume` targets stop the operation without transformation; no new CLI flags were added.
- Added LSP semantic-token delta responses with stable result IDs, a bounded per-open-document history, and automatic full-response fallback when a delta base is unavailable or not smaller on the wire.
- Changed VBA LSP document synchronization to incremental UTF-16 ranged edits while retaining full-document replacements for client compatibility; document-version snapshots invalidate only the changed document, while semantic-token caches refresh when workspace symbols change.
- Added tree-sitter incremental parsing for open VBA documents. New immutable snapshots parse against an edited clone of the prior tree, preserve in-flight readers of the old tree, safely retain the last valid snapshot for unreconcilable changes, and log incremental/full-fallback outcomes with `--performance-log`.
- Added immutable, document-version-scoped LSP analysis snapshots that share normalized source lines, procedure ranges, source symbols, and semantic-token metadata across requests while preserving open-buffer and UserForm freshness.
- Added opt-in structured LSP performance logging through `xlflow lsp --performance-log`, including operation timing and result metadata for diagnostics and language features.
- Improved LSP semantic token performance by reusing document and workspace symbol data within each request and caching encoded tokens for unchanged document versions.
- Improved LSP responsiveness by caching document source symbols per document version and content, sharing the result across symbols, completion, hover, CodeLens, semantic tokens, and type inference while preserving unsaved-buffer and UserForm freshness.
- Coalesced debounced LSP diagnostics so each open document uses at most one worker, obsolete generations are never published, and closing a document leaves an empty diagnostics result as the final notification.

## v0.24.0

- Added `xlflow capabilities --json`, a versioned, machine-readable projection
  of the central command-coordination registry for VS Code and external
  integrations. The v1 contract publishes stable command IDs, CLI paths, and
  safety/wait/recovery policy metadata without exposing bridge selectors; it is
  advisory and the CLI lock remains authoritative.
- Added Windows cross-process workbook coordination with crash-released
  `LockFileEx` ownership, immediate structured `workbook_busy` diagnostics, and
  guarded atomic owner metadata for conflicting CLI processes and
  WSL-delegated commands.
- Added opt-in `--wait` and `--wait-timeout` workbook coordination with a finite
  30-second default, Ctrl+C cancellation, and structured timeout/cancel errors.
- Added an observational top-level `coordination` section to `session status`,
  including current workbook busy state and guarded public owner metadata when
  available without changing existing session fields or failure behavior.
- Added persistent workbook recovery quarantine for uncertain Excel/VBA
  termination, including atomic per-workbook markers, fail-closed
  `workbook_recovery_required` diagnostics, recovery-aware status/session/process
  output, managed `session stop --discard`, and verified or force
  `recovery clear` workflows that remain distinct from the operating-system
  workbook lock.
- Fixed recovery quarantine edge cases so fatal non-session VBA component
  replacement failures publish recovery metadata, and managed
  `session stop --discard` fails with `session_stop_cleanup_unconfirmed` when
  the recorded Excel process is still running but cannot be reacquired.
- Documented and exhaustively tested UserForm/Designer coordination so migration,
  snapshot, build/apply, image export, form inspection/listing, pull, and push all
  converge on the same canonical workbook lock before Excel or VBIDE starts.
- Changed new-project scaffolds and `module install` to place bundled `Xlflow*.bas` helpers under `src/modules/Xlflow/`; existing root-level helper files are not moved, and `module install` now refuses those legacy collisions rather than creating duplicate VBA components.
- Fixed .NET `push` to stop before importing when Excel cannot remove an existing VBA component and to reject Excel-renamed imports, preventing duplicate modules with a `1` suffix; every removal/import failure poisons a partial session replacement, managed sessions discard the unsafe unsaved state on stop, and external sessions receive owner-correct recovery guidance.
- Fixed .NET bridge workbook persistence so transient `__XLFLOW_MODE__`, related runtime/UI/debug defined names, and generated helper modules are removed before save, including after a timed-out session run or a macro that saves and then edits again, preventing manually opened workbooks from remaining in headless mode or retaining temporary harness code.
- Added non-configurable `VB015` lint/preflight validation for VBA logical lines with more than 24 line continuations, preventing opaque Excel import failures before `push` or `run` opens Excel.
- Added strict reusable UserForm spec validation based on the canonical contract, including unknown-field, type, fixed-value, control-property, parent-reference, and type/ProgID mismatch diagnostics before `form build` opens Excel.
- Fixed `form build` so a UserForm caption is persisted through the VBComponent property path and remains consistent with runtime forms; `form export-image` now captures the actual runtime caption instead of substituting the Designer value.

## v0.23.0

- Added `xlflow form migrate sidecar` for converting imported UserForms from `frm` code-source mode to sidecar code plus Designer specs, and added `xlflow init --userform-code-source sidecar` for opting imported workbooks into the modern UserForm layout.
- Fixed non-executing UserForm Designer snapshots to capture top-level form width and height from VBComponent properties when available.
- Fixed `.NET` UserForm Designer builds so top-level form width and height from sidecar specs are applied through VBComponent properties instead of staying at Excel's default size.

## v0.22.0

- Added configurable automatic backup retention through `[backup.retention]`, disabled by default, with workbook-scoped pruning after successful backup-producing `push` and `rollback` operations.
- Expanded bundled `XlflowAssert.bas` with strict equality, `Null` / `Empty`, numeric tolerance, string, regex, array, `Range.Value2`, and object identity assertions, plus typed assertion failure formatting for terminal and JSON output.
- Added parameterized VBA tests with `@TestCase(...)`, including named cases, per-case `id` / `qualified_name`, source/runtime discovery JSON, exact case filtering, scalar literal validation, and `invalid_test_case` diagnostics.
- Added `@Skip("reason")` and `@Todo("reason")` metadata for VBA tests, including discovery `status_hint`, non-executed `skipped` / `todo` results, separate CLI summary counts, and VS Code Testing API surfacing.
- Added editor annotation completion for xlflow `@Skip(...)` and `@Todo(...)` test metadata alongside `@ExpectedError(...)`.
- Added `@ExpectedError(number[, description[, source]])` metadata for VBA tests, including source/runtime discovery JSON, expected-error execution matching, `observed_error` JSON for expected-error results, and `invalid_test_metadata` diagnostics for malformed annotations.
- Added `xlflow edit sheet add` for live-session worksheet creation, including idempotent `--if-missing` usage and positioned `--before` / `--after` insertion.
- Added `xlflow edit formula` for live-session range formula edits, including R1C1 and A1 formula assignment, event control, optional target-range calculation, and structured edit summaries for AI-agent formula workflows.
- Added `xlflow backup prune` and `xlflow backup delete --backup <id>` for explicit backup cleanup, including dry-run previews, workbook-scoped deletion, invalid/legacy cleanup flags, safe managed-root deletion checks, and structured JSON summaries.
- Added VS Code extension rollback and backup pruning workflows, including backup Quick Pick selection, active-session safeguards, post-rollback pull/inspect actions, failed-push rollback offers, and localized command/menu entries.
- Added VS Code editor support for xlflow VBA documentation comments, including Quick Fix snippet generation from `'''`, Rubberduck `@Description` annotation completions in comments, doc-comment continuation on Enter, and highlighting for `'''` comments and Rubberduck description annotations.
- Added xlflow-style documentation comments to scaffolded `Xlflow*.bas` helper modules so their public APIs show useful Hover and Signature Help documentation.
- Added stable qualified VBA test identifiers (`id` / `qualified_name`) to `xlflow test list --json` and `xlflow test --json`, allowing duplicate procedure names across modules and qualified `xlflow test --filter Module.TestName` selection.
- Added isolated temporary workbook execution for `xlflow test`, including `--isolation none|module|test`, `--no-save`, and JSON `test_run` metadata. Plain non-session test runs now protect the configured workbook by running against `.xlflow/test-runs/<run-id>/` copies.
- Added `xlflow test --fail-fast`, `--max-failures`, and `--rerun-failed`, including `not_run` results for early termination, flaky pass reporting, attempt history, and execution metadata in JSON output.
- Improved backup handling so corrupted backup entries no longer block `backup list`, valid backup JSON now includes `size_bytes`, incomplete backup creation is cleaned up, and failed `.NET` pushes report successfully created pre-push backups.
- Expanded the scaffolded `SampleTests.bas` to demonstrate smoke tags, parameterized `@TestCase(...)`, `@ExpectedError(...)`, `@Todo(...)`, common assertions, and test execution commands without introducing failing sample tests.
- Added first-class `.xlsb` support for Excel COM/VBIDE-backed VBA workflows, including `new`, `init`, source pull/push, sessions, run/test/save, backup/rollback, and UserForm operations, while keeping `.xlsm` as the default project format.
- Added stable `workbook_format_unsupported` failures for `.xlsb` on direct OOXML/file-level features such as `formulas pull`, workbook inspect, workbook cell diff, and pure-Go `pack`.
- Added `.xlam` project creation to `xlflow new`, using Excel add-in file format `55` while keeping `.xlsm` as the default, and documented `.xlam` initialization through `xlflow init`.
- Fixed `.xlam` session reuse in the .NET bridge by resolving open add-in workbooks through direct filename lookup with full-path validation, and by making VBE Compile target activation tolerate add-in workbooks without normal visible workbook windows.

## v0.21.0

- Added conservative VBA keyword and known built-in identifier casing normalization to `xlflow fmt` and LSP document formatting, enabled by default with `[fmt].keyword_casing` and `[fmt].builtin_casing`.
- Added LSP diagnostics for high-signal analyze warnings, including `VBA201`, `VBA204`, `VBA208`, `VBA209`, and `VBA212`.
- Fixed VBA LSP document formatting so incomplete or syntactically invalid buffers are skipped without surfacing an internal parser error notification in VS Code.

## v0.20.0

- Added `xlflow session attach` to adopt an already-open configured workbook as an external xlflow session, and deprecated legacy `xlflow attach --active` validation-only usage.
- Added VS Code extension actions to connect to an already-open workbook from the session menu and open the configured workbook in Excel from the Project view.
- Added onboarding warnings for disabled Excel VBA object model access in `xlflow doctor` and related setup failures, including a real temporary-workbook VBProject probe, plus VS Code prompts when `.bas`, `.cls`, or `.frm` files are not associated with xlflow's `vba` language.
- Fixed `VB009` false positives for valid VBA strings such as the official VBA-JSON `JsonConverter.bas` escaped quote output (`json_Char = "\"""`).
- Fixed VBA LSP diagnostics so `Me` is recognized as the current instance in UserForm, workbook, sheet, class, and non-xlflow fallback modules, preventing false `VB029` undeclared warnings for sheet code such as `Me.Rows(...)`.

## v0.19.0

- Added `xlflow formulas pull` to extract worksheet formulas and defined names from `.xlsx` / `.xlsm` files into deterministic region-based JSONL snapshots without launching Excel, including standalone `--src` and `--out` options, region dependency indexes, and parse status summaries.
- Added `xlflow formulas inspect` to summarize formula snapshots, list sheet/range regions, locate a cell's formula region, expand supported R1C1 patterns, and emit agent-friendly JSON.
- Added `xlflow pull --formulas` to refresh the default `formulas/` snapshot after a successful VBA pull.
- Added formula snapshot guidance to the bundled `xlflow` Agent Skill so AI agents know when to inspect `formulas/` outputs before changing VBA or workbook layout.

## v0.18.0

- Improved human-readable CLI output with clearer sections, status markers, warnings/hints, and table-style summaries while preserving `--json` contracts.
- Added `xlflow run --push` to import edited VBA source into the configured workbook before running a macro.
- Improved `xlflow run` macro-failure diagnostics to suggest `xlflow push` / `xlflow run --push` when source files are newer than the saved workbook.
- Enabled `VB020` unused local variable warnings by default; projects can opt out with `[lint].disabled_rules = ["VB020"]`.
- Updated generated `xlflow.toml` files to show how to opt out of `VB020` and how to opt in to heavier project-wide rules such as `VB021` unused private procedures.
- Improved `VB020` unused local variable detection so write-only assignments no longer count as variable references.
- Fixed VBA LSP diagnostics so `VB020` unused local variable warnings appear in editor diagnostics for the current buffer.
- Added VS Code Quick Fix actions for xlflow diagnostics to insert line-level suppression comments.
- Fixed `.NET` `xlflow run --interactive` so native VBA UI such as `MsgBox` is left for the user to dismiss instead of being misreported as `macro_failed`.
- Fixed `.NET` Excel cleanup so direct `run`/`test` executions and `session stop` do not leave owned Excel processes behind after successful commands.

## v0.17.0

- Updated `xlflow doctor` to run project-independent diagnostics successfully even when `xlflow.toml` is missing, with warnings and setup hints instead of a config failure.
- Added `xlflow update check` for structured GitHub Release update checks, and wired the VS Code extension to notify users when the installed xlflow CLI is behind the latest release.
- Added `xlflow type db status/init/refresh/clean` for global generated TypeLib databases, with an initial Excel TypeLib importer feeding LSP type intelligence.
- Added AST-aware VBA operator spacing to `xlflow fmt`, enabled by default and configurable with `[fmt].operator_spacing`.
- Added AST-aware VBA declaration spacing to `xlflow fmt`, enabled by default and configurable with `[fmt].declaration_spacing`.
- Added conservative VBA LSP `prepareRename` and `rename` support for high-confidence local, private module, private procedure, and label symbols while refusing host, TypeLib, public API, UserForm/event, ambiguous, and unresolved targets.
- Updated `xlflow lsp` to load generated TypeLib databases when present while continuing to work with only the embedded built-in database.
- Added best-effort generated TypeLib DB initialization after successful `new` and `init`; failures are reported as warnings and do not fail project creation.
- Extended generated TypeLib databases with registry-derived ProgID mappings and `--library all`, improving LSP `CreateObject("...")` late-binding inference for installed Excel, Scripting, ADODB, MSForms, Office, and VBIDE libraries.
- Changed `xlflow type db refresh` to always regenerate the generated TypeLib database; `--force` remains accepted for compatibility but is no longer required.
- Updated `xlflow doctor` to report generated TypeLib DB status and suggest initialization or refresh when the global DB is missing or stale.
- Added best-effort generated TypeLib DB initialization when `xlflow lsp --stdio` or `xlflow lsp --check` starts and the global DB is missing or stale.
- Hardened generated TypeLib DB clean/refresh behavior to reject unsafe clean targets, avoid loading stale generated files outside the manifest, and keep best-effort multi-library imports from failing the whole run when one TypeLib cannot be imported.
- Updated the VS Code extension to avoid passing the default `.xlflow/lsp.log` path in non-xlflow workspaces, so syntax/LSP-only users do not get workspace log files unless they configure one.
- Updated the VS Code extension to hide LSP CodeLens run actions in non-xlflow workspaces.
- Fixed LSP ProgID completions so `CreateObject(`, `CreateObject "..."`, and `CreateObject(Class:=...)` contexts surface late-binding candidates instead of requiring an already-open string literal.
- Improved LSP ProgID completion ranking and details by prioritizing version-independent ProgIDs and labeling versioned ProgIDs such as `Excel.Application.16` or `Forms.CommandButton.1`.
- Improved VBA LSP intelligence for common Excel idioms, including With-relative call-chain signature help, collection default-member signatures such as `ListObjects(1)`, multi-line signature help across `_` continuations, named-argument completions such as `Destination:=`, and `WScript.Shell` / `VBScript.RegExp` `CreateObject` inference.
- Added high-confidence TypeLib-powered VBA LSP diagnostics for unknown concrete members, close-name suggestions, read-only/write-only property misuse, `Set` misuse, incompatible object assignments, and no-return method value use.
- Expanded VBA LSP top-level expression completions with `Set ... = New`, `Nothing`, object-producing built-ins such as `GetObject`, and common VBA `Is*`/type-inspection functions including signature help.
- Expanded the curated VBA standard library database with string, date/time, math, conversion, file-system, interaction, and financial functions plus common `vb*` constants for completion and signature help.
- Improved VBA formatting so explicit line-continuation tails are indented one additional level while preserving existing line-number alignment behavior.
- Added `VB032` lint/LSP/preflight validation for repeated `?` Debug.Print shorthand such as `?? "hoge"`, reporting it before Excel/VBE interaction.

## v0.16.1

- Fixed `VB029` false positives for module-level declarations inside `#If ... #Else ... #End If` conditional compilation blocks.
- Fixed `VB029` false positives for Excel member chains such as `Cells(ws.Rows.Count, "A").End(xlUp).Row`, where string or numeric arguments were mistaken for undeclared receiver identifiers.

## v0.16.0

- Removed the legacy PowerShell Excel bridge for v0.16.0. The only supported bridge modes are now `auto` and `dotnet`; `--bridge powershell`, `XLFLOW_EXCEL_BRIDGE=powershell`, and `[excel].bridge = "powershell"` now fail with bridge-mode/configuration errors instead of emitting a deprecation warning.
- Added source-only `xlflow module new` and `xlflow form new` scaffolding commands for standard modules, class modules, and sidecar UserForms.
- Added source-only `xlflow module remove` and `xlflow module rename` commands for safe standard-module, class-module, and UserForm source mutations before the next `xlflow push`.
- Added explicit unknown-command reporting so mistyped commands print a stderr error with help/suggestions, and `xlflow --json <unknown-command>` returns a structured `unknown_command` failure.
- Added LSP CodeLens support for runnable no-argument VBA `Sub` procedures, including `Run` and `Run Test` actions backed by in-memory editor buffers.
- Added `xlflow test list --json` for source-only VBA test discovery, enabling editor integrations to discover tests without parsing VBA in TypeScript or opening Excel.
- Clarified the CLI JSON contract for editor integrations, including single-document stdout output and inactive `session status --json` payloads with `session.active=false`.
- Added non-configurable `VB031` lint/preflight validation requiring standard `.bas` modules to include `Attribute VB_Name`.
- Fixed VBA LSP diagnostics so the built-in `Err` object is recognized as a global, preventing false `VBA029` undeclared warnings for calls such as `Err.Raise`.

## v0.15.0

- Added initial VBA LSP signature help and argument diagnostics for common project procedures and built-in VBA/COM members, including active parameter tracking and argument-count warnings.
- Improved VBA LSP signature help for parenless calls so typing a space after calls such as `dict.Add` can show argument hints and early argument-count diagnostics.
- Expanded built-in VBA/COM signature metadata for common VBA functions, Excel range/workbook methods, WorksheetFunction calls, Scripting.FileSystemObject/TextStream, and ADODB calls used by LSP signature help and argument diagnostics.
- Hardened LSP argument diagnostics to avoid declaration-line false positives, improved diagnostic method names for signatures with return types, and documented manual signature-help smoke checks for the VS Code dev client.
- Improved VBA LSP completion and diagnostics for E2E smoke scenarios, including namespace type completion such as `Excel.`, scoped inferred member completion for `dict.`, `amountRange.`, and `rs.`, built-in completions such as `True`, `CStr`, `Now`, `Debug.Print`, and `ByVal`, With-block signature help for `.Offset(`, and out-of-scope member receiver diagnostics.
- Fixed VBA LSP member completion inside nested call expressions such as `CStr(dict.` and after With-relative call chains such as `.Offset(1,0).`.
- Fixed VBA LSP module declaration snippets so `Option Explicit` is not suggested again on the next line after it already exists.
- Suppressed empty-prefix VBA LSP completions on blank module-level lines after existing content, preventing the completion list from reopening immediately after `Option Explicit`.
- Added LSP document formatting support so VS Code Format Document can call the same VBA formatter engine used by `xlflow fmt` and receive full-document `TextEdit` results for `.bas` and `.cls` buffers.
- Improved VBA hover display with member signatures, canonical receiver types, and source/provenance notes for declarations, inferred object types, built-in globals, UserForm controls, and built-in object model members.
- Added context-aware ProgID completion inside `CreateObject("...")` while suppressing unrelated completions in ordinary string literals.
- Added active `With` block receiver tracking for VBA hover and member completion, including nested `With .Member` blocks.
- Fixed VBA statement snippets so an already completed declaration such as `Option Explicit` is not immediately suggested again.
- Expanded VBA type inference so `Set target = <known expression>` propagates the right-hand expression type into hover and completion, including method-return chains such as `FileSystemObject.OpenTextFile`.
- Strengthened Excel table collection metadata and resolution tests for `ListObjects(...).ListColumns(...).DataBodyRange` chains, including Japanese table and column names.
- Added Excel PivotTable/PivotField and Shape/TextFrame metadata so LSP hover and completion can resolve chains such as `PivotTables(...).PivotFields(...).DataRange` and `Shapes(...).TextFrame.Characters.Text`.
- Added practical `Sheets(...)` inference so worksheet-like member chains such as `Sheets("Input").Range("A1")` and `ThisWorkbook.Sheets("Input").ListObjects(...)` resolve as worksheet operations without changing the underlying `Excel.Sheets` database type.
- Expanded practical chain coverage for `FileSystemObject.OpenTextFile`, `Application.Workbooks.Open`, workbook worksheet collections, table ranges, pivot fields, and shape text frames.
- Added context-aware completion for `Call ...` and `Set x = ...`, limiting suggestions to callable symbols and object-producing expressions respectively.
- Added reusable `xlflow lsp --stdio` VBA language server support with full-document synchronization, diagnostics, document/workspace symbols, definition lookup, references, hover, completion, and a practical built-in VBA/COM type database for common Excel, MSForms, Scripting, ADODB, VBIDE, Office, and VBA constant metadata.
- Fixed LSP document symbols so incomplete VBA declarations do not return empty symbol names that VS Code rejects while editing.
- Added `.` as an LSP completion trigger character so VS Code requests member completions such as `Range("A1").Font.Color` while typing.
- Added built-in `VBA.Collection` metadata so LSP hover and member completions resolve `Dim result As Collection` and `Set result = New Collection` correctly.
- Fixed LSP type inference to prefer the nearest in-scope VBA declaration before the cursor, avoiding stale same-name declarations such as `result As Boolean` overriding a local `result As Collection`.
- Added LSP completions for module-qualified VBA procedure calls such as `Utils.BuildName` after typing `Utils.`.
- Added module-level VBA declaration snippet completions such as `Public Sub`, `Public Function`, `Dim`, `Const`, `Type`, and `Enum` while typing at the top level of a module.
- Fixed LSP document symbols for empty or incomplete VBA files so VS Code does not reject module symbols whose selection range exceeds the source range, and added identifier trigger characters so declaration snippets appear while typing.
- Fixed module-level declaration snippets so multi-word prefixes such as `Public S` are completed by replacing the typed declaration prefix instead of disappearing after a modifier is typed.
- Fixed LSP completion visibility so `Private` declarations from other modules, including `Private Const`, are not suggested as cross-module candidates.
- Added a procedure-local `Dim` snippet completion and `VB029` diagnostics for undeclared assignment targets or loop control variables when `Option Explicit` is present.
- Added VBA LSP type-position completions for declarations such as `Dim ws As Workbook`, `Function Foo() As String`, and `Set dict = New Dictionary`, including built-in VBA types, COM type aliases, and project class modules.
- Tuned VBA LSP completion and editing behavior by limiting server-side trigger characters to `.`, keeping procedure-local completion candidates scoped to the current procedure, using UTF-16 symbol selection ranges, and debouncing diagnostics after document changes.
- Added an LSP workspace symbol cache for saved source files plus open-document overlays, reducing repeated full-project symbol indexing during completion, hover, definition, and workspace symbol requests.
- Improved VBA LSP definition and reference resolution for procedure-local variables and constants so same-name locals in other procedures are no longer returned for the current local scope.
- Added VBA parameter symbols so LSP definition and reference lookup can resolve procedure arguments within the current procedure scope.
- Improved VBA LSP hover for local symbols and parameters by reusing scoped definition lookup and avoiding type inference from later declarations.
- Added UserForm `.frm` control extraction for VBA LSP intelligence, enabling hover, completion, and definition support for controls such as `Me.txtName.Text` and `Me.Controls("txtName").Text` without opening Excel.
- Expanded the built-in Excel VBA/COM type database with common formatting and worksheet helper objects such as `Excel.Font`, `Excel.Interior`, `Excel.Borders`, `Excel.Validation`, `Excel.Hyperlinks`, `Excel.PageSetup`, `Excel.AutoFilter`, `Excel.Sort`, and `Excel.WorksheetFunction`.
- Expanded built-in Excel constant metadata for common formatting, border, alignment, page orientation, and sort constants, and included enum group information in constant hover output.
- Updated LSP diagnostics to reuse the same file-local VBA lint rules as `xlflow lint` for unsaved editor buffers, including stable `VB...` diagnostic codes and diagnostic clearing when issues are fixed.
- Hardened LSP file URI and range handling for Windows paths, UNC paths, escaped Japanese paths, and UTF-16 diagnostic positions after non-ASCII text.
- Fixed LSP workspace symbols so an open editor buffer hides stale filesystem symbols for the same module, preserving the in-memory document priority used by definition and reference features.
- Updated `tree-sitter-vba` to v0.8.1 and adapted call extraction and lint member-access checks to the new stable `receiver` / `member` / `arguments` AST fields.
- Deprecated the legacy PowerShell bridge for v0.15.0. Windows `auto` bridge mode now uses the `.NET` bridge without falling back to PowerShell; explicit `--bridge powershell`, `XLFLOW_EXCEL_BRIDGE=powershell`, and `[excel].bridge = "powershell"` remain available only as deprecated opt-ins and emit `powershell_bridge_deprecated`. PowerShell bridge removal is planned for v0.16.0.
- Fixed `xlflow analyze` false positives for `VBA209` array element and UDT-array member assignments, clarified `VBA204` fallthrough guidance, and recognized paired Push/Pop `Application` state restore helpers for `VBA203`.
- Extended experimental `xlflow pack` to update an existing UserForm's code-behind when the form already exists in the template (ADR-0012 Stage 2), honoring `[userform].code_source`: `frm` reads the code from the `.frm`, while `sidecar` (the default) reads it from `src/forms/code/<FormName>.bas` and merges it in memory without writing the source tree. The form's designer storage is carried through from the template byte-for-byte and `.frx` is not read, so layout is preserved but not authored; a `.frm` whose form is not in the template fails loudly with `pack_userform_generation_unsupported`, and a sidecar carrying `Attribute VB_*` headers or with no matching `.frm` fails with `pack_ambiguous_layout`. See `docs/specs/pack-command.md`.

## v0.14.1

- Improved `xlflow pack` protected-project detection: it now reads the CMG `ProjectProtectionState` bits (MS-OVBA §2.3.1.15) instead of a DPB password-length heuristic, so protected and unprotected projects are classified by the spec-defined signal rather than a corpus-calibrated threshold.
- Added `.NET doctor` diagnostics for the Windows systemprofile Desktop directories required by non-interactive Excel COM workbook automation. Missing directories now return `systemprofile_desktop_missing` with instructions to create both `System32` and `SysWOW64` Desktop paths, while permission-denied inspection results are reported as warnings instead of false missing-directory failures. `xlflow doctor` remains lightweight by default, while the new `--workbook` option opens the configured workbook and reports `workbook_openable`.

## v0.14.0

- Added inline VBA suppression comments for lint and analyze diagnostics, supporting `xlflow:disable-next-line <ID>` and `xlflow:disable-line <ID>` with stable IDs such as `VB002` and `VBA205`, plus warnings for unknown, unsupported, or unused suppressions. Preflight-blocking errors remain unsuppressible.
- Documented COM cleanup best practices for VBA tests that open external workbooks, including `Close` plus `Set ... = Nothing` to avoid file locks during test hooks.
- Fixed test and macro discovery so Unicode VBA procedure names such as Japanese `Test*` and `*_Test` names are recognized by both PowerShell and `.NET` bridges.
- Added experimental `xlflow pack`, a pure-Go, cross-platform command that builds an `.xlsm` artifact from the source tree plus a workbook template without Excel. It regenerates `xl/vbaProject.bin` from `.bas`/`.cls` sources and replaces only that single zip entry, leaving the rest of the workbook untouched. Gated behind `--experimental`; supports standard, class, and unambiguous document modules, carries existing UserForm designer streams through byte-for-byte, and performs no VBE compile or runtime validation (every run reports `pack.vbe_validation = "not_performed"` and a `vbe_validation_skipped` warning). Fails loudly on protected or signed projects, UserForm generation, ambiguous layouts, active sessions, and in-place overwrite of the template or configured workbook. See `docs/specs/pack-command.md` and ADR-0012.
- Updated `tree-sitter-vba` to v0.7.0 and removed `xlflow fmt` parser-workaround fallback for legacy numbered comments, colon-separated block lines, and valid line-continuation forms now handled by the grammar.
- Refactored `xlflow fmt` to use `tree-sitter-vba` structure-aware indentation for supported VBA blocks while preserving comments, strings, attributes, line continuations, line-number workflows, and `.frm` skip behavior.
- Added `[lint].disabled_rules` and `[analyze].disabled_rules` for disabling configurable source feedback rules by stable diagnostic ID, with compatibility warnings for legacy per-rule booleans.
- Refactored `xlflow lint` to use `tree-sitter-vba` AST-backed checks for core declaration, member-access, and local code-shape rules, including per-declarator implicit `Variant` diagnostics and parser recovery findings.
- Refactored `xlflow analyze` to use `tree-sitter-vba` AST-backed procedure context and added runtime-risk findings for `Range.Find` `Nothing` guards, object initialization, Application state restore, error-handler fallthrough, unqualified Excel object access, ByRef mismatch candidates, Dictionary/Collection guards, `ReDim Preserve`, object/array comparisons, function return paths, and expanded Excel member mismatches.
- Documented the full `xlflow lint` rule list on the command page, including `VB001` through `VB014` codes and severity levels.
- Added `xlflow inspect calls`, a source-only tree-sitter-vba call-site extractor for exported VBA files with caller context, argument summaries, source ranges, conservative project-symbol resolution, JSON output, and compact grouped text output.
- Added `xlflow inspect symbols`, a source-only tree-sitter-vba symbol indexer for exported `.bas`, `.cls`, and `.frm` VBA files with JSON and compact outline output.
- Updated `xlflow inspect symbols` for the tree-sitter-vba 0.6.0 declaration node shape changes, including split property and declare nodes.
- Added `VB028` source preflight blocking for bare `MsgBox` / `InputBox` calls when `XlflowUI.bas` is present, so `push` fails before Excel opens with guidance to use `XlflowUI` wrappers or explicit `VBA.Interaction.*` native dialogs.

## v0.13.1

- Fixed `xlflow form snapshot` so Designer snapshots no longer require executing an injected helper macro, avoiding Trust Center / Insider Beta Office failures that blocked temporary macro workbooks from running.
- Fixed `.NET` and PowerShell Designer inspection to recover concrete UserForm control types from COM metadata when `ProgId` is unavailable, so snapshots no longer persist generic `__ComObject` / `Control` types for standard MSForms controls.
- Fixed `.NET` Designer inspection for controls such as `TabStrip` that do not expose a `Controls` collection, preventing `DISP_E_UNKNOWNNAME` failures when snapshotting forms that contain a broad set of MSForms controls.

## v0.13.0

- Added WSL development support that delegates Excel-related commands to Windows `xlflow.exe`, translates Windows-mounted project paths, preserves command streams and exit codes, and extends `doctor` with WSL/Windows diagnostics. Linux x64 release archives are now published for the WSL frontend.
- Added a GitHub Pages-hosted WSL/Linux frontend installer at `https://harumiweb.github.io/xlflow/install.sh`, including one-line `curl | sh` install guidance and `--uninstall` support.
- Added `task wsl-install`, `task wsl-uninstall`, and `task uninstall` helpers for installing or removing local Go bin xlflow binaries during delegation testing.
- Added a GitHub Pages-hosted PowerShell installer at `https://harumiweb.github.io/xlflow/install.ps1`, including review-first safer install guidance, one-line quick install guidance, and `-Action uninstall` support that removes `%LOCALAPPDATA%\xlflow` and its user PATH entry.
- Hardened `xlflow-excel-bridge.exe` direct startup so no-arg, help, version, and invalid launches exit immediately with clear output, while `xlflow.exe` uses an explicit internal run flag before starting the actual bridge runtime.

## v0.12.2

- Fix .NET bridge VBA export decoding for non-ASCII pull sources

## v0.12.1

- Fixed `.NET` bridge stdin/stdout JSON transport on Windows to use explicit UTF-8 streams, preventing mojibake and invalid bridge JSON when project, workbook, or session paths contain Japanese or other non-ASCII characters.

## v0.12.0

- Fixed: Add support for detecting implicit variants inside user-defined types (UDTs) in linter
- Reduced the default `xlflow run --json` failure payload for AI-agent and normal VBA debugging loops. The default run JSON now promotes the best available `location` and `suggestion` to top-level fields and hides verbose workbook/bridge/runtime diagnostics, dialog snapshots, and location-capture metadata unless `--verbose` is specified.
- Fixed `.NET` bridge macro runs so Excel is stabilized and the STA message loop is pumped around `Application.Run`, improving reliability for formatting/layout operations such as `Range.Interior.Color`, `Range.Clear`, row height updates, and protection state reads. Fatal COM/RPC failures such as `0x800706BE` now return `excel_com_rpc_failure` with `h_result` and run diagnostics, and live sessions are marked poisoned instead of being silently reused.
- Added best-effort `.NET` VBE selection diagnostics for suppressed compile/runtime dialogs in `xlflow run --bridge dotnet --json` and compile failures in `xlflow push`, including selected component, procedure, source path, source-file line range, selected line text, and non-fatal capture-attempt metadata when Excel exposes it.
- Improved `.NET` dialog watcher button diagnostics and action selection by capturing Win32 button `access_key`, `control_id`, and `enabled` metadata. VBA runtime suppression now prefers accelerator keys such as `D` for Debug and `E` for End before localized text fallback, improving tolerance for non-English Excel/VBE UI.
- Fixed `.NET` VBE runtime location capture after `xlflow session start` so the first failing `run` no longer reports stale module header lines such as `Option Explicit` before VBE has moved the selection to the actual error line.
- Removed the legacy runtime-debug command surface completely. VBA-internal debugging is now documented around `XlflowDebug.Log`, `xlflow run --json`, structured diagnostics, and `Erl`/line-number workflows. Legacy `XlflowLog` / `XlflowSetTraceFile` usage is now treated as removed API surface in source analysis and preflight guidance.
- Added explicit VBA line-number operations to `xlflow fmt` via `--line-numbers preserve|add|remove|renumber`, including conservative ambiguity warnings for numeric-label control flow and structured JSON summary fields under `output.line_numbers`.
- Fixed `xlflow fmt --line-numbers add` so it no longer numbers `Select Case`, `Case` / `Case Else`, or `End Select` control lines, avoiding VBA compile failures when the first `Case` in a select block is numbered.
- Fixed `xlflow fmt --line-numbers add` so explicit VBA line-continuation statements only number their first physical line; continuation tail lines now stay unnumbered to avoid compile failures.
- Added a dedicated xlflow agent debugging reference at `internal/agentskill/templates/xlflow/references/debugging.md`, documenting the preferred workflow: inspect `run` diagnostics first, then use `fmt --line-numbers add` plus targeted `XlflowDebug.Log` only when the default error metadata is not enough.

## v0.11.0

- Added native `.NET` bridge support for the remaining Windows workbook commands: `xlflow new`, `session start|status|save|stop`, `attach --active`, `runner install|status|remove`, `list forms`, `ui button add|list|remove`, and `edit cell|range|rows|columns` with explicit `--bridge dotnet --json`.
- Packaged the `.NET` Excel bridge into Windows release ZIPs as `xlflow-excel-bridge.exe` beside `xlflow.exe`, while documenting AppLocker, WDAC, AV, and unsigned-executable caveats for managed Windows environments.
- Added native `.NET` bridge support for `xlflow test --bridge dotnet --json`, enabling VBA test discovery and execution through the .NET bridge. Supports `Test*`/`*_Test` naming, `@Tag("...")` annotations, `BeforeAll`/`AfterAll`/`BeforeEach`/`AfterEach` hooks, inconclusive detection (`vbObjectError + 516`), runtime injection, UI/debug stream helpers, and session-aware workflow. Auto mode keeps the existing PowerShell behavior; use `--bridge dotnet` explicitly to route through the .NET bridge.
- Enhanced `xlflow run --bridge dotnet` with MsgBox/InputBox/FileDialog response injection (`--msgbox`, `--inputbox`, `--filedialog`), UI stream pipe support (`--ui-stream`), and `__XLFLOW_DEBUG_PIPE__` injection for `XlflowDebug.Log` transport. Previously these options were rejected by the .NET bridge; they are now fully supported with the same behavior as the PowerShell bridge.
- Fixed `.NET` runtime injection cleanup so partial defined-name injection is rolled back and execution aborts when injection cannot be completed.
- Added native `.NET` bridge support for `xlflow macros --bridge dotnet --json` and `xlflow run Module1.Main --bridge dotnet --json`, enabling macro discovery and execution through the .NET Excel bridge without PowerShell. Supports typed arguments including finite invariant-culture `double` values, fully qualified macro names, save/no-save/save-as, timeout, session attachment, and structured error handling for `macro_failed`, `macro_not_found`, and `macro_disabled`. Auto mode keeps the existing PowerShell behavior for macros/run; use `--bridge dotnet` explicitly to route through the .NET bridge.
- Added a reusable `.NET` Excel/VBE dialog watcher that captures runtime, compile, MsgBox, InputBox, and FileDialog snapshots with Win32/UI Automation identity metadata. Runtime error dialogs are suppressed without requiring Excel focus, and unattended runs prefer End over Debug to avoid leaving VBE in break mode.
- Added native `.NET` bridge support for `xlflow pull --bridge dotnet --json` and `xlflow push --bridge dotnet --json`, enabling VBA component export/import through the .NET Excel bridge without PowerShell. Auto mode keeps the existing PowerShell behavior for pull/push; use `--bridge dotnet` explicitly to route through the .NET bridge.
- Added native `.NET` bridge support for runner-backed `xlflow inspect workbook|sheets|range --session --bridge dotnet --json` and `xlflow process list|cleanup --bridge dotnet --json`, including `--bridge auto` fallback from unsupported/runtime/protocol `.NET` failures back to PowerShell for supported commands.
- Added native `.NET` `xlflow doctor --bridge dotnet --json` diagnostics for runtime and Excel COM probing, plus documentation clarifying that top-level `bridge` metadata remains provider-specific between PowerShell and `.NET` bridges.
- Added structured COM error fields (`h_result`, `details`) to `xlflow doctor --bridge dotnet --json` error output. COM activation failures now include the HRESULT hex code and exception details alongside the error message.
- Added an Excel bridge provider abstraction in Go, moved PowerShell invocation behind `PowerShellProvider`, and added bridge selection via persistent `--bridge`, `XLFLOW_EXCEL_BRIDGE`, and `[excel].bridge` while keeping `auto` on the existing PowerShell behavior for now.
- Added `xlflow fmt` as a conservative, non-destructive VBA source formatter for `.bas` and `.cls` files. Supports `--write`, `--check`, `--diff`, `--json`, and `--stdin` modes. The formatter uses 4-space indentation, strips trailing whitespace, normalizes blank lines, preserves class module metadata, and is idempotent. Typical workflow: `fmt -> lint -> push -> run/test`.
- Refined the interactive `xlflow new` / `init` welcome screen with a new `Welcome to` heading, a command reference URL, and softer muted version/info text below the ASCII logo.
- Hardened the bundled TAKT orchestra, PR review, and issue bug workflows with explicit verification, audit-triage, and release-gate handling, broader loop monitoring around remediation and final audit, and clearer guidance to treat allowed untracked files and auto-staging state as non-blocking.
- Added `xlflow process list` to enumerate all local Excel processes with PID and open-workbook status.
- Added `xlflow process cleanup <pid>`, `xlflow process cleanup --auto`, and `xlflow process cleanup --all [--yes]` for safe and forceful Excel process termination. `--auto` targets only workbook-free processes; `--all` is a destructive force-stop of all local Excel instances with mandatory interactive confirmation or `--yes`.
- Fixed `XlflowDebug.bas` helper module to stop forwarding `Log`'s `ParamArray` into a secondary helper procedure, preventing VBA compile/runtime failures such as "Sub または Function が定義されていません" and "ParamArray の使い方が適切ではありません" in some hosts.
- Fixed `.NET` `xlflow run` compile-dialog handling so VBE compile errors that surface during macro invocation are suppressed, returned as structured `vba_compile_failed` / `compile_vba` diagnostics, and no longer block unattended workflows.
- Fixed `.NET` `xlflow push` so imported VBA is VBE-compiled before saving or updating source fingerprints, matching the legacy PowerShell bridge behavior and returning structured `vba_compile_failed` / `compile_vba` diagnostics for broken source.

## v0.10.0

- Fixed `xlflow run --diagnostic` compile watcher to return structured `vba_compile_failed` errors when the VBE compile command control cannot be found, instead of silently reclassifying the failure as `vbide_access_denied`.
- Improved runtime dialog capture for `xlflow run --diagnostic` so break-mode inspection prefers user-code lines over temporary `XlflowRun_*` helpers, and the runtime macro runner now executes in a disposable child PowerShell process so break-mode resets do not leave the parent CLI hung.
- Fixed `xlflow run --diagnostic` VBE compile preflight to locate `Compile VBAProject` from the VBE menu bar (`Id = 578`) instead of assuming the Debug toolbar contains it, and to treat a disabled compile command as "already compiled" rather than a hard failure.
- Fixed `xlflow ui button add` so it auto-reuses a matching live session workbook when `.xlflow/session.json` points at the configured workbook, preventing the Excel SaveAs dialog that previously appeared when a session was active.
- Extended `ui button add`, `ui button list`, and `ui button remove` to use the shared session-aware workbook open helper and explicit save/release cleanup, matching the behavior of `push`, `pull`, `run`, and other workbook-backed commands.
- Added `xlflow status` and `xlflow status --json` as a read-only project-level command that shows project, source, workbook, and session state in one output. Source freshness is a heuristic based on file mtimes; the command does not modify workbook files, source files, or `.xlflow/state`. `workbook_saved` is now derived from `save_required` instead of `dirty` to avoid contradictory results when the session probe reports `save_required=true` but `dirty` is unknown; baseline `session` payload now always includes `running`, `workbook_open`, and `metadata` for schema stability.
- Added `xlflow init --with-module` so imported projects can immediately receive bundled runtime helper modules and sync them back into the copied workbook.
- Added `xlflow module install [--push]` so existing xlflow projects can install bundled helper modules on demand without rerunning `new`.
- Removed `--keepalive` / `--keepalive-interval` from Excel COM-backed commands and the final `XLFLOW_DONE` marker; interactive stderr now uses spinner progress where available, while non-interactive runs fall back to line-oriented progress and streamed UI/debug stderr output suppresses separate progress frames.
- Added XlflowUI module with MsgBox and InputBox wrappers to handle user prompts.
- Extended XlflowUI with headless-safe file dialog wrappers for `Application.GetOpenFilename`, `Application.GetSaveAsFilename`, open `Application.FileDialog`, and folder picker flows, plus repeated `--filedialog <kind>:<dialog-id>=<value>` CLI responses for `run` and `test`.
- Added `--ui-stream` for `xlflow run` and `xlflow test`, streaming resolved headless `XlflowUI` dialog events to stderr in real time while preserving JSON stdout and returning final `ui.events` payloads plus human-readable `UI` summaries.
- Added scaffolded `XlflowDebug` helper support so explicit `XlflowDebug.Log` calls stream to stderr and final top-level `debug` payloads during `xlflow run` and `xlflow test` without a separate CLI flag, including direct and fast run paths.
- Updated run.ps1 and test.ps1 to accept MsgBoxResponsesJSON and InputResponsesJSON parameters.
- Added explanatory comments to scaffolded `XlflowRuntime.bas`, `XlflowUI.bas`, and `XlflowAssert.bas` so workbook authors can adopt the helper modules more easily.
- Added explicit live-session inspect mode for `inspect workbook`, `inspect sheets`, `inspect range`, `inspect used-range`, and `inspect cell` via `--session`, plus explicit `live_session` target metadata and saved-file warnings that point callers to live-session inspect when disk may be stale.
- Added runtime-aware execution mode injection for `run` and `test`, plus the scaffolded `XlflowRuntime` VBA helper for branching between interactive, headless, agent, CI, and test execution contexts.
- Enhanced `xlflow macros --json` output with `component_type`, `visibility`, `has_parameters`, `runnable`, `reason_not_runnable`, and `run_command` fields per macro so AI agents and users can choose the correct entrypoint without guessing.
- Added `default_entry` and `suggestions` fields to `xlflow macros --json` output, surfaced from `project.entry` in `xlflow.toml` and resolved against discovered runnable macros.
- Added `--runnable` flag to `xlflow macros` to filter the output to only directly runnable procedures.

## v0.9.0

- Added winget release publishing so GoReleaser can generate the `HarumiWeb.Xlflow` manifest and push it to the `harumiWeb/winget-pkgs` fork for upstream submission.
- Updated `xlflow new` to bootstrap the workbook/source sync automatically by pushing scaffolded VBA into the new workbook before the command reports success, and added placeholder `src/workbook/ThisWorkbook.bas` / `Sheet1.bas` files with `Option Explicit` for new projects.
- Updated `xlflow init` to bootstrap source sync automatically by pulling VBA from the copied workbook into `src/`.
- Added first-class workbook rollback support with `xlflow backup list` and `xlflow rollback`, including metadata-backed workbook-file backups under `.xlflow/backups/<backup-id>/`, automatic safety backups before restore, and session-aware guards that refuse rollback while the target workbook is open in an active xlflow session.
- Changed default `push` backups from component-export snapshots to rollback-capable workbook snapshots, and updated the CLI/docs surface, JSON output, and VitePress command/concept pages to reflect the new backup and recovery workflow.

## v0.8.1

- Fixed `xlflow inspect form <name> --designer --session` so normal designer inspection no longer takes the strict temporary-workbook path, reducing the sample `space-invader` session inspection from about one minute to a few seconds.
- Corrected PowerShell boolean parsing and case-insensitive variable handling around the `StrictDesigner` flag, preventing `"False"` string values from being treated as truthy.
- Hardened UserForm runtime cleanup guards in `inspect form` and `form export-image` so null runtime workbook state does not trigger unnecessary Excel COM cleanup and finalizer waits.

## v0.8.0

- Completed the UserForm feature set for issue #25 across phase 1 through phase 7, including explicit UserForm warnings in core workbook flows, `xlflow list forms`, `inspect form` for designer/runtime/both, `form snapshot`, and experimental runtime image export.
- Hardened `form export-image` for real Excel GUI behavior by repairing generic runtime captions from designer state, choosing the correct monitor-relative work area instead of forcing the primary screen, using DPI-aware capture sizing, and trimming capture artifacts so the exported PNG matches the visible UserForm more faithfully.
- Corrected UserForm build round-tripping so snapshot-derived width and height no longer grow on each `form build` cycle, preserving the persisted Designer dimensions from `snapshot` output.
- Updated the bundled docs, CLI contract, and agent guidance to reflect the UserForm discovery, inspection, snapshot, export, and warning workflow, including the experimental status of runtime image export.
- Strengthened PowerShell script coverage with behavior-oriented tests for the UserForm build and export helpers, replacing narrow string-presence checks where practical.

## v0.7.0

- Added `xlflow edit cell`, `edit range`, `edit rows`, and `edit columns` as minimal workbook-mutation helpers for AI-agent testing and visual tuning in a live Excel session.
- Added session-only workbook edit behavior for the new `edit` commands, including `--events keep|on|off` support for cell value and formula changes so `Worksheet_Change` flows can be exercised without generating temporary VBA.
- Commands now display explicit workbook state, including whether reading from saved file or live Excel session
- Added warnings when live session workbooks contain unsaved changes
- Extended workbook-backed JSON and human output with explicit `target` / `session` metadata across session-aware commands, plus top-level `edit` payloads for workbook mutation summaries.
- Updated the CLI contract, README files, ADR session policy note, and bundled xlflow skill guidance to cover the new edit workflow and session-state visibility.

## v0.6.0

- Added `xlflow export-image` to export worksheet ranges as PNG images for visual verification, including session-aware targeting, structured `target` / `output` metadata, and reliability fixes so hidden-workbook captures do not produce blank images or leak Excel processes.
- Added `--include-style` flag to `inspect range` and `inspect used-range` commands to display worksheet style metadata including cell fills, borders, merged cells, row heights, and column widths.
- Added Rubberduck-compatible folder-aware VBA sync so `xlflow pull` and `push` can round-trip nested source trees via `@Folder(...)`, recursive source discovery, duplicate module-name preflight, and nested `.frm`/`.frx` companion handling.
- Added `[vba]` configuration defaults for folder sync control, wired the settings through the Go/PowerShell bridge, and documented the new contract in the CLI spec, READMEs, and bundled xlflow skill.
- Fixed folder-sync path handling to stay compatible with Windows PowerShell 5.1 and hardened `pull` so it does not clear the existing exported source tree before the workbook opens successfully.
- Added `--no-update-check` and `XLFLOW_NO_UPDATE_CHECK=1` so interactive `new` and `init` can skip the GitHub Release lookup used by the scaffold welcome banner.
- Hardened GitHub Release packaging with stable `checksums.txt` SHA256 output and archive SBOM generation via GoReleaser.
- Extended the release workflow to install Syft and publish GitHub artifact attestations for release archives, checksums, and SBOM artifacts.
- Documented Windows-side release verification in both READMEs, including SHA256 checks, `gh attestation verify`, and the current non-goal of Authenticode signing.

## v0.5.0

- Added richer sample VBA projects, including the `world-news` NewsAPI example and the `stock-price` dashboard example, plus accompanying screenshots and README updates.
- Improved runtime error handling and diagnostics so CLI runs surface failures more clearly across the Go and PowerShell execution bridge.
- Refined release documentation and sample project metadata with formatting fixes and README polish, including Japanese README badge updates.

## v0.4.0

- Added `xlflow inspect` with workbook, sheet, range, used-range, and cell inspection for saved workbook snapshots.
- Added inspect-specific formatting and range limits so agents can read workbook structure and output without opening Excel.
- Updated the bundled xlflow agent skill and CLI contract docs to teach snapshot-first inspect workflows.

## v0.3.0

- Added automatic reuse of a matching live xlflow session workbook for workbook-backed commands when `--session` is omitted.
- Added structured save-state reporting so `push`, `run`, `session status`, and related commands can surface when a live session workbook differs from disk and needs `xlflow save`.
- Improved `run` with compile-first diagnostic mode, clearer direct-run restrictions, and fallback to `project.entry` when no macro argument is provided.
- Expanded the legacy runtime-debug lifecycle handling with helper injection and session-aware workbook reuse.
- Added a verbose `version` command that reports build metadata, script resolution, supported features, and executable details.
- Added update-checking and refreshed version/welcome messaging.
- Updated bundled PowerShell scripts, agent skill guidance, and JSON envelopes to match the new session-aware behavior.

## v0.2.0

- Bundled the PowerShell scripts used by xlflow for Excel session management, testing, tracing, and UI button manipulation.
- Added the initial session-aware command surface for opening, reusing, saving, and stopping Excel workbooks.
- Added run, pull, push, test, and UI button workflows built on the bundled PowerShell bridge.
