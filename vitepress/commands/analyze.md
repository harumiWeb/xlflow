# xlflow analyze

Analyze VBA source for runtime-risk patterns without Excel COM.

## Usage

```bash
xlflow analyze
```

## Options and Arguments

| Option / argument | Description                          | Default |
| ----------------- | ------------------------------------ | ------- |
| `--json`          | Return structured analysis findings. | false   |

## Examples

```bash
xlflow analyze
xlflow analyze --json
```

## Notes

::: tip
Use `analyze` for fast source-level feedback before opening Excel.
:::

`VBA203` tracks each recognized `Application` state write through normal,
early-exit, error-handler, termination, and unknown CFG exits. A direct saved
value (including an exact saved-variable copy) restores the state only when it
reaches that exit on every path; cleanup labels are supported. It covers
`ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, `Calculation`, `StatusBar`,
`Cursor`, `Interactive`, `AskToUpdateLinks`, `AutomationSecurity`, and
`CutCopyMode`. The existing Push/Pop helper exception still prefers same-module
Pop/Restore helpers and accepts a uniquely resolved project-visible standard
module pair. Ambiguous, missing, external, or dynamically bound helper paths
are not treated as proof of restoration. This internal analysis does not add
fields to JSON output or report new caller-level diagnostics.

> [!IMPORTANT]
> Findings that block automation return a failure status and exit code `1`.

## Rules

The generated [static-analysis diagnostic catalog](../reference/diagnostics) is
the authoritative reference for rule metadata, including configuration,
default state, scope, precision, preflight behavior, and inline suppression.

| Code     | Severity              | Description                                                                                                                       |
| -------- | --------------------- | --------------------------------------------------------------------------------------------------------------------------------- |
| `VBA101` | warning               | Object variable assignment likely missing `Set`.                                                                                  |
| `VBA102` | warning               | Object-returning project function assignment likely missing `Set`.                                                                |
| `VBA103` | warning               | Object-returning function body likely missing `Set <FunctionName> = ...`.                                                         |
| `VBA104` | error                 | Known Excel object/member mismatch such as `Worksheet.DisplayGridlines`.                                                          |
| `VBA105` | error                 | Removed `XlflowLog` trace helper call.                                                                                            |
| `VBA106` | error                 | Removed `XlflowSetTraceFile` trace helper call.                                                                                   |
| `VBA201` | warning               | `Range.Find` result is dereferenced before a `Nothing` check.                                                                     |
| `VBA202` | warning               | Object variable may be used before an obvious `Set` assignment.                                                                   |
| `VBA203` | warning               | `Application` state is changed without an obvious restore path.                                                                   |
| `VBA204` | warning               | Normal execution can fall through into an error-handler label.                                                                    |
| `VBA205` | warning               | Ambiguous Excel workbook or worksheet scope depends on UI state or ordering.                                                      |
| `VBA206` | warning               | Runtime-safety warning for temporary, property/member, indirect, or uncertain ByRef arguments.                                    |
| `VBA207` | warning / information | `Dictionary` or `Collection` item access has no obvious existence guard.                                                          |
| `VBA208` | warning               | `ReDim Preserve` is used on a multi-dimensional array.                                                                            |
| `VBA209` | warning               | Object or array is compared with scalar equality.                                                                                 |
| `VBA210` | warning               | Function or Property Get may reach normal exit without a valid return assignment.                                                 |
| `VBA211` | error                 | Expanded known Excel object/member mismatch.                                                                                      |
| `VBA212` | warning               | `Nothing`/`IsArray` guard and matching nested access share an eager Boolean/selection expression.                                 |
| `VBA213` | warning               | Direct Dictionary iteration key is used as an object or value.                                                                    |
| `VBA214` | warning               | `On Error Resume Next` extends beyond a narrow compatibility probe.                                                               |
| `VBA215` | warning               | `Range.Find`/`Replace` omits saved Excel search settings.                                                                         |
| `VBA216` | error                 | A range expression mixes distinct explicit worksheet roots.                                                                       |
| `VBA217` | warning               | A last-row calculation has an implicit root or unstable boundary strategy.                                                        |
| `VBA218` | warning               | An Excel API failure contract is consumed without its required guard.                                                             |
| `VBA219` | warning               | A local Workbook or VBA file handle can exit without a matching Close.                                                            |
| `VBA220` | warning               | An Excel or UserForm event handler can re-enter itself or a related event.                                                        |
| `VBA221` | warning               | A local helper can leave an Application property changed for its caller.                                                          |
| `VBA222` | warning               | A public API exposes an inaccessible, ambiguous, or unresolved type.                                                              |
| `VBA223` | warning               | Likely hardcoded secret detected in VBA source.                                                                                   |
| `VBA224` | warning               | Conservative procedure-local analysis finds untrusted data at a sensitive API.                                                    |
| `VBA225` | warning               | Cell-by-cell Excel object-model work is repeated inside a non-trivial loop.                                                       |
| `VBA226` | warning               | A `Range.Value` result is used with an unsafe scalar or array shape assumption.                                                   |
| `VBA227` | warning               | Array allocation, lifecycle, dimension, bound, or object-element safety is not proven.                                            |
| `VBA228` | error                 | Definite ByRef type or array-shape mismatch rejected by the VBE; blocks source preflight.                                         |
| `VBA229` | error                 | Unresolved procedure-local `As <Type>` identifier; blocks source preflight.                                                       |
| `VBA230` | warning               | Dictionary `CompareMode` is changed after an entry was added.                                                                     |
| `VBA231` | warning               | Dictionary `Keys` or `Items` is repeatedly materialized in a loop.                                                                |
| `VBA232` | warning               | Dictionary key normalization is inconsistent.                                                                                     |
| `VBA233` | warning               | A late-bound Dictionary uses an undefined Scripting comparison constant.                                                          |
| `VBA234` | warning               | A Collection is mutated while the same object is being enumerated.                                                                |
| `VBA235` | warning               | A zero-based index is used directly with a one-based Collection.                                                                  |
| `VBA236` | warning               | A process-launch command may combine external input, an unquoted executable path, a command-line secret, or unobserved execution. |

Disable configurable analyzer rules with `[analyze].disabled_rules`:

```toml
[analyze]
disabled_rules = ["VBA205", "VBA211"]
```

Legacy per-rule booleans such as `forbid_unqualified_excel_objects = false` remain accepted for compatibility, but xlflow emits a deprecation warning. If both formats disagree, `disabled_rules` takes precedence and xlflow reports a conflict warning.

Use inline suppression comments for intentional local exceptions while keeping rules enabled globally:

```vb
' xlflow:disable-next-line VBA205
Range("A1").Value = 1

Cells(1, 1).Value = 2 ' xlflow:disable-line VBA205
```

Multiple IDs may be listed with spaces. Unknown IDs, unsupported preflight-blocking IDs, and suppressions that no longer match an analyzer diagnostic are reported as warnings.

`VBA205` reports active UI roots such as `ActiveWorkbook` and `Selection`, unqualified worksheet members (`Range`, `Cells`, `Rows`, and `Columns`), unqualified `Sheets` / `Worksheets`, numeric `Workbooks(...)` / `Windows(...)` indices, discarded `Workbooks.Open` results, and `ThisWorkbook` in add-in standard modules. It names the ambiguous root and suggests an explicit workbook, worksheet, range, or captured open result. In an add-in, sheet-collection guidance uses an explicit caller workbook rather than `ThisWorkbook`. Intentional interactive entrypoints can use the local suppression shown above.

`VBA210` checks every reachable path to a `Function` or `Property Get` normal exit, including `Exit Function`, `Exit Property`, error-handler paths that return normally, and shared cleanup labels. A dominating return assignment satisfies all paths; VBA's default-initialized return value does not. Known non-returning `Err.Raise` statements are treated as exceptional exits rather than normal fallthrough. Object returns require `Set`, known value returns require ordinary assignment or `Let`, and the diagnostic reason identifies a representative uncovered exit when practical. The rule is opt-in through `detect_function_return_path` and remains batch-only.

Rules `VBA201` through `VBA206`, `VBA208`, `VBA209`, `VBA211`, `VBA212`, `VBA214` through `VBA227`, and `VBA230` through `VBA236` are enabled by default. `VBA230` through `VBA236` are warning-level, non-blocking, inline-suppressible rules available in batch and real-time analysis. `VBA222` is a batch-only warning that checks public function/property return types, all public parameters, and custom event parameters against project visibility and the available TypeLib database. Standard modules and `VB_Exposed=True` classes/interfaces are public API surfaces; private or unexposed project types, ambiguous names, and unresolved external types are reported conservatively. Host-required event handlers are excluded. Suppress an intentional case with `xlflow:disable-line VBA222` or `xlflow:disable-next-line VBA222`, or add `VBA222` to `[analyze].disabled_rules`.
`VBA223` is a default-enabled, non-blocking, file-local, realtime warning. It uses structural credential patterns, ignores obvious placeholders where possible, and redacts source snippets with `[REDACTED]`.

`VBA224` is a conservative, procedure-local warning: it reports source, sink, and propagation path context, treats unsupported transformations as unknown, does not propagate taint across procedures, and does not block preflight. Literals and explicit constant/allowlist branches are accepted; `EncodeURL` is accepted only for HTTP URLs, while generic `Trim`, `CStr`, `Replace`, `IsNumeric`, and `Len` do not remove taint. Use `xlflow:disable-line VBA224`, `xlflow:disable-next-line VBA224`, or `[analyze].disabled_rules = ["VBA224"]` for an intentional flow. `VBA206` remains a configurable warning for literal temporaries, parenthesized, property/member, array-element, indirect, and otherwise uncertain ByRef forms. Literal arguments do not produce blocking `VBA228` errors because the VBE evaluates them as temporary values. `VBA228` owns only explicit statically incompatible bare value/object/array variables and `Long`/`LongPtr`/`LongLong` mismatches, including named arguments; it is always enabled, cannot be suppressed by `VBA206` settings or inline comments, and blocks source preflight. Array-valued Function return slots retain their array shape, local values shadow same-named procedures, and callee-module-qualified project types match the callee's unqualified type declaration; a different module qualification remains incompatible. `Object`, `Variant`, `Any`, unresolved, and late-bound types remain uncertain and do not produce `VBA228`. The legacy `detect_byref_argument_mismatch` key remains supported for `VBA206`. Rules `VBA207`, `VBA210`, and `VBA213` are opt-in through legacy `[analyze]` settings because they are more dataflow-sensitive. `VBA207` uses `warning` when absence is definite and `information` when existence is unknown. `VBA213` applies only when a known `Scripting.Dictionary` is iterated directly and the key variable is used as an object or value; ordinary key iteration remains valid. `VBA214` is warning-only and allows one compatibility probe followed by `On Error GoTo 0` (with optional `Err.Number` inspection and `Err.Clear`); scopes containing wider control flow, calls, or un-restored exits are reported without severity escalation or preflight blocking. `VBA215` requires explicit `Find` `LookIn`, `LookAt`, `SearchOrder`, and `MatchByte`, or `Replace` `LookAt`, `SearchOrder`, `MatchCase`, and `MatchByte`, because Excel can reuse saved Find/Replace dialog or macro settings when they are omitted. `VBA216` blocks preflight only when xlflow can prove that explicit range roots refer to different worksheets. `VBA217` guides last-row calculations that rely on the active sheet, `End(xlDown)`, unadjusted `UsedRange.Rows.Count`, or `CurrentRegion`; it does not block preflight. `VBA218` accepts `On Error GoTo <label>` for exception-raising APIs, or only a narrow `On Error Resume Next` probe that checks `Err` and immediately restores `On Error GoTo 0`; an unbounded `Resume Next` scope is not sufficient. `Variant/Error` APIs require `IsError` before consumption. `VBA219` tracks only captured local `Workbooks.Open` results and VBA `Open ... As #handle` calls. It accepts direct local aliases, error-handler cleanup labels, pre-Open file-number aliases, and ownership transfer only at an object-returning Function's normal exit; it intentionally does not assume that parameters, helper calls, or other COM resources are owned. Diagnostics `VBA101` through `VBA106` are always enabled.

Dictionary/Collection analysis recognizes early-bound and late-bound Dictionaries, local construction, direct aliases, additions/removals, `.Exists` branches, and uniquely resolved local helper effects. `VBA230` checks `CompareMode` ordering, `VBA231` catches repeated `.Keys`/`.Items` allocation in loops, `VBA232` checks explicit key normalization, and `VBA233` catches undefined Scripting comparison constants on late-bound Dictionaries. `VBA234` rejects mutation of the same Collection during `For Each`, while `VBA235` catches direct zero-based indexing. Safe `.Exists` branches, one-operation Collection probes restored immediately with `On Error GoTo 0`, cached key/item arrays, explicit `.Items` value iteration, one-based Collection loops, and `vbBinaryCompare`/`vbTextCompare`/`vbDatabaseCompare` are accepted.
`VBA225` is enabled in batch and real-time analysis and reports resolved cell-by-cell Excel reads, writes, formulas, formatting, lookups, worksheet-function calls, and helper effects inside non-trivial loops; nested-loop context is retained in the message but does not escalate its warning severity. Bulk range operations and statically provable loops of three or fewer iterations are exempt.

`VBA226` is enabled in batch and real-time analysis and tracks procedure-local `Range.Value` / `Value2` shapes. It reports one-dimensional or scalar assumptions for definite multi-cell ranges, dimensionless bounds, statically provable dimension/order/bounds mistakes, and incompatible known destination ranges. Multi-cell values are modeled as two-dimensional arrays; single-cell values are modeled as scalars. Dynamic, reassigned, and branch-merged shapes remain uncertain, so only unsafe consumption or a statically proven shape mismatch is reported. Use `values(row, column)`, dimension-specific bounds, and a dominating `IsArray` guard for dynamic values when appropriate.

`VBA227` is enabled in batch and real-time analysis and tracks array allocation through the CFG. Fixed arrays start allocated; dynamic arrays start unallocated; `ReDim`, `Erase`, array assignments, and proven project-local array returns update the state, while unknown `Variant` and external values remain conservative. In real-time analysis, array-return summaries are limited to the active document. It reports unsafe `LBound` / `UBound` and indexed access, invalid dimensions or known bounds, fixed-array `ReDim`, and unknown Variant array operations. `VBA208` remains the owner of `ReDim Preserve` findings, while object-array missing-`Set` findings remain owned by `VBA101` / `VBA102`. `Range.Value` / `Value2` shape cases remain owned by `VBA226`; use `xlflow:disable-line VBA227`, `xlflow:disable-next-line VBA227`, or `[analyze].disabled_rules = ["VBA227"]` for intentional exceptions.

`VBA236` owns process-launch safety for VBA `Shell`, `WScript.Shell.Run` /
`Exec`, `Shell.Application.ShellExecute`, and Win32 `ShellExecute` variants.
It separates executable path, ordinary arguments, interpreter command text,
URL/document target, window style, wait flag, and result observation when the
call shape permits. The additive context reports `injection`,
`process_launch`, paired with `tainted_command_text`, `unknown_origin`,
`credential_exposure`, `observability`, or `unquoted_executable_path` risk
kinds. Only known tainted input with a known
origin and role is a potential command-injection claim; unknown input is a
general process-launch warning. Constant command text avoids the
tainted-command-text risk. Quoting a trusted executable path addresses the
unquoted-path risk only; it does not sanitize ordinary arguments or interpreter
command text such as `cmd.exe /c` and PowerShell `-Command`. Independent
credential and observability risks remain visible. For Windows guidance, quote
the complete executable path, avoid interpolated `cmd.exe /c` and PowerShell
`-Command` text, prefer PowerShell `-File` with fixed arguments, and keep
secrets out of command lines. Use `xlflow:disable-line VBA236`,
`xlflow:disable-next-line VBA236`, or `[analyze].disabled_rules = ["VBA236"]`
for intentional exceptions. When retaining the legacy `VBA224` fallback,
disable or suppress both `VBA224` and `VBA236` for a complete process-launch
exception. Generic non-process source-to-sink findings remain owned by
`VBA224`.

`VBA229` is a realtime and batch compile-equivalent error for unresolved type identifiers in procedure-local `Dim` and `Static ... As <Type>` declarations. It uses the production built-in/host/TypeLib and project-symbol resolver, including enum, class, UserForm, and document-module types, points at the type identifier, cannot be suppressed, and blocks source preflight. When the generated TypeLib manifest is missing, malformed, empty, or otherwise incomplete, lookup misses are left unreported because none of those metadata states proves that a referenced type is absent. Parameters, return types, and module-level declarations are outside the v1 rule scope.

## JSON Output Example

Failed `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "failed",
  "command": "analyze",
  "error": {
    "code": "analyze_failed",
    "message": "1 analysis finding(s) found"
  },
  "analysis": [
    {
      "code": "VBA201",
      "severity": "warning",
      "file": "src/modules/Main.bas",
      "module": "Main",
      "procedure": "Run",
      "line": 12,
      "message": "Range.Find result found is dereferenced before a Nothing check."
    }
  ]
}
```

## Related

- [lint](./lint)
- [check](./check)
- [inspect-gui](./inspect-gui)

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow analyze` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

## Prerequisites

Check the project configuration and run `xlflow doctor --json` before workbook-backed operations. Source-only commands can run without Excel; commands that read or mutate a workbook require Windows Excel and VBIDE access.

## What this command reads and changes

The command reads the inputs and configuration described in its syntax and examples. Treat source files, the saved workbook, and a live session as separate states; add `--session` when the live workbook is authoritative. Any mutation is reversible only when a backup or explicit session save boundary exists.

## Effect on source-of-truth state

Use `xlflow status --json` before and after the command. A source edit normally requires `push`; a workbook edit normally requires `pull`; a dirty live session requires `save --session` or an intentional discard.

## Common workflows

Combine this command with the relevant [source/workbook/session workflow](../concepts/workbook-session-source), and use `--json` in scripts and agent loops.

## Common failures

Read the structured `error.code`, exit code, and recovery metadata instead of scraping terminal text. The [symptom-oriented troubleshooting guide](../help/troubleshooting) maps installation, execution, session, VS Code, and WSL failures to recovery steps.
