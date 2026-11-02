# Conditional-compilation analysis

Lint, analyze and LSP diagnostic parsing use a shared conditional-compilation
projection. Before parsing, provably inactive content is replaced by spaces
with exactly the same byte length. CR, LF and active source positions remain
unchanged. Resolved conditional wrappers are also replaced by spaces so a
selected declaration header can join its shared body. Unknown wrappers remain
available to structure and uncertainty checks. Original files, formatter input and workbook import source
are not rewritten by this projection.

## Constant environment

Project arguments can be declared in `xlflow.toml`:

```toml
[vba.conditional_constants]
DEBUG_FLAG = "False"
FeatureLevel = "2"
VBA7 = "True"
Win64 = "True"
Win32 = "True"
Mac = "False"
```

Keys are case-insensitive VBA identifiers; duplicate case variants are errors.
Reserved identifiers from MS-VBAL 3.3.5.2 are rejected, including `If`, `And`
and `Debug`. Use a name such as `DEBUG_FLAG` instead. The same check applies
to active module `#Const` declarations; an invalid name retains the original
source conservatively rather than selecting branches from an invalid binding.
VBA's reserved set does not include all VB.NET keywords: `Catch` remains valid.
Values are TOML strings containing independently evaluable VBA literal
expressions. Boolean, integral, floating-point, currency, string and Empty
values are accepted. Date, Null, Nothing, unresolved names and unsupported
expressions are configuration errors. The table is unordered: one entry
cannot reference another. No compiler arguments are read from or written to
the workbook by analysis. Match this table to the target VBA project and
restart the LSP after changing project configuration.

An active module-level `#Const` is evaluated in source order and overrides the
project value within that module. It does not leak into other modules.
Inactive `#Const` declarations do not affect later conditions. Conditional
assignments are merged at `#End If`; different possible values remain unknown.

Undefined ordinary identifiers evaluate as VBA Empty, which coerces to zero.
The target built-ins `VBA6`, `VBA7`, `Win16`, `Win32`, `Win64` and `Mac` remain
unknown unless explicitly configured. Analysis never guesses target constants
from the OS or architecture of the xlflow process.

## Evaluation and fallback

The explicit conditional evaluator reuses the shared constant-expression
parser. It supports Boolean and numeric conditions, arithmetic, comparisons,
Boolean/Empty numeric coercions and integral logical operators. `Not` binds
after comparisons and before `And`. These coercions do not change ordinary
declaration-constant evaluation.

Decimal Currency literals with at most four fractional digits retain exact
scaled integer values, including comparisons between adjacent large values.
Currency literal rounding, exponent notation, arithmetic and exponentiation
remain unknown rather than selecting a branch through floating-point rounding.

Unsupported coercions, runtime calls, string comparisons, signed-boundary
hexadecimal/octal literal widths, Single arithmetic rounding, mixed-precision
Single/Currency comparisons, wide integer comparison precision and fractional
bitwise operands remain
unknown. Both candidate branches of an unknown condition remain analyzable.
Malformed or unclosed conditional directive structures, and unsupported
continued directive headers, conservatively retain the original document so
syntax diagnostics cannot be hidden. Source-encoding validation also remains
applicable to the complete authored source, including inactive branches.

This change preserves conservative handling of unresolved conditional
declarations in resolution, CFG and dataflow. It does not turn an unknown
target into confirmed VBE compilation evidence. A retained candidate branch
with unsupported syntax may therefore still produce VB014 or analyze_failed.

VBA266 keeps its existing conservative lexical usedness policy: mentions in
the authored source, including conditional directives and inactive branches,
can suppress an unused-private-constant warning. Its declaration candidates
still come from the projected active source.

## Regression contracts

The tests cover inactive prose from Issue #966, active errors, nested and
alternate branches, case-insensitive project constants, module scope and
override behavior, uncertain assignment joins, malformed structures, CRLF and
Unicode byte positions, and unchanged raw source. Editor coverage also
protects changes to conditional constants in unsaved buffers and prevents
stale parse artifacts from surviving a branch switch.

The local VBE oracle supplies language observations separately from parser
recovery diagnostics. Oracle infrastructure failures never justify fixture
promotion or analyzer semantic changes.

## Performance acceptance

Revision preparation checks the exact snapshot and constant environment before
scanning for directives. Each snapshot records directive presence once. A raw
structural snapshot containing directives still requires projection; a
configured snapshot, including a caller-supplied masked source, retains its
prepared artifacts. Changed source, configuration and retired snapshots must
not reuse a previous projection. Disabled diagnostic entrypoints check their
enablement and TypeDB prerequisites before preparing the source, preserving
their existing cancellation and error boundaries.

VBA266 indexes authored lexical mentions in one pass for ordinary ASCII
candidate names. Match boundaries and the old regexp's non-overlapping match
selection remain unchanged; non-ASCII names and Unicode case-fold equivalents
retain the regexp fallback. The loop-invariant selector rule shares a lazy,
immutable module string-constant index between procedure lanes. Its lookup
retains source-order prefixes and procedure overrides. This optimization does
not repair the pre-existing extraction guard (`len(match) == 2` for a regexp
with two capture groups), which currently yields no extracted string constants;
enabling that proof requires a separate semantic fix and regression evidence.

These requirements apply to the Issue #966 projection and its batch/editor
integration. They are a local regression gate, not an absolute response-time
promise for arbitrary projects or a change to the repository-wide benchmark
helper's investigation thresholds. Correct source selection, uncertainty,
diagnostic contracts and authored-source delivery remain prerequisites.

### Budgets

Compare each workload separately against the same pinned pre-change commit.
Let `R` be the head median divided by the base median for that workload.

| Dimension                      | Acceptance requirement                                                                                       |
| ------------------------------ | ------------------------------------------------------------------------------------------------------------ |
| Elapsed time (`ns/op`)         | `R <= 1.10`, with the upper bound of a two-sided 95% confidence interval for `R` also `<= 1.10`              |
| Allocated bytes (`B/op`)       | `R <= 1.05`                                                                                                  |
| Allocation count (`allocs/op`) | `R <= 1.05`                                                                                                  |
| Deterministic work counters    | Stable across repetitions; no unexplained increase in traversals, evaluations, index builds or invalidations |

Do not average workloads together: a cold improvement cannot compensate for a
warm or editor regression. Lower allocations do not waive the time budget.
`B/op` measures cumulative allocation, not peak resident memory; no peak-memory
guarantee follows from this gate. Absolute millisecond targets require a
separately defined reference host and interactive request distribution.

A semantic change may legitimately change participation. Explain positive
counter deltas using source/procedure ownership, eligibility and related
counters; aggregate totals alone do not prove unchanged cost. In particular,
exchanging compact CFG walks for legacy walks requires a route-cost comparison
or profiling in addition to the unchanged sum. Unrelated procedures must not
acquire repeated work merely because conditional projection is enabled.

### Required workloads

- Both ROneCOne and std-vba, each in `cold`, `warm`, `local-edit` and
  `dependency-edit` modes, with identical rule configuration and corpus pins.
- Repeated editor hover, definition, completion and signature-help requests on
  one unchanged revision, for ordinary and large modules. Include sources
  without directives and sources with resolved and unresolved directives.
- An ordinary body edit and a conditional constant/directive edit. Measure
  these separately: a conditional edit is permitted to invalidate the entire
  affected document because branch selection can change distant declarations.
- Focused projection measurements at increasing source sizes and repeated
  unchanged-document preparation. Separate initial directive detection and
  projection from snapshot reuse; warm corpus analysis alone does not cover
  interactive preparation overhead.

Existing entry points include `BenchmarkRealWorldCorpus`,
`BenchmarkLSPInteractiveIndexOrdinary`, `BenchmarkLSPInteractiveIndexIssue756`,
`BenchmarkLSPIncrementalSingleCharacterEdit` and `BenchmarkProjectDependencyIncrementalBodyEdit`.
The existing editor fixtures do not establish coverage of every conditional
scenario above; add focused coverage before claiming that part of the gate.

Projection and directive-detection work must be owned by document revisions,
not repeated per procedure or per diagnostic rule. A cache hit must preserve
the existing interactive-index and procedure-catalog reuse contracts. Any
remaining full-source detection on repeated preparation must be measured and
accounted for; it is not free because no new CFG work is recorded. No-directive
documents must retain their existing incremental parsing eligibility.

### Measurement and verdict

1. Pin exact base/head commits and verify both checkouts. Use the same Windows
   host, Go/CGO toolchain, power configuration, environment, corpus and flags.
   Record these with the raw logs. An uncommitted run needs verified source
   hashes; a later commit name alone does not identify its compiled inputs.
2. A matched `-benchtime=1x -benchmem -count=3` run is screening evidence.
   Final acceptance requires at least ten observations per revision and
   workload, consistent with the retained-baseline procedure in
   `static-analysis-corpus.md`. Editor throughput measurements may use a longer
   fixed benchtime, identical on both revisions and recorded explicitly.
3. Run serially without concurrent tests, lint, builds or Excel automation.
   Balance order to reduce drift: five `base, head, head, base` blocks provide
   ten observations per side. Exclude only a documented infrastructure failure
   and repeat its whole block; never discard a valid slow sample.
4. Retain every sample, median and spread. For elapsed-time acceptance, compute
   the confidence interval by resampling whole order-balanced blocks, retaining
   the resampling method and seed. The comparison helper reports medians and
   suspicious changes; it does not implement this confidence gate.
5. Classify the result as **pass** only when all required workloads and budgets
   are satisfied. A memory-budget breach or confirmed unrelated repeated work
   is **fail**. Missing coverage, unstable counters, insufficient samples or
   a time interval crossing the limit is
   **unverified** and blocks a no-performance-regression claim. If the entire
   time interval exceeds `1.10`, that workload fails even before its cause is
   identified.

Verify corpus diagnostics and focused scalability/reuse tests after any
optimization. Preserve the public semantic boundary rather than changing
snapshots or retaining inactive declarations to recover benchmark time.
