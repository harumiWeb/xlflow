# xlflow metrics

Calculate deterministic maintainability metrics for VBA procedures without
opening Excel.

## Usage

```bash
xlflow metrics
xlflow metrics --json
```

`metrics` is source-only. It scans the configured `[src]` roots and `tests`,
uses the procedure IR/CFG and uniquely resolved project-local calls, and does
not run `lint`, `analyze`, `check`, LSP diagnostics, preflight, or VBA. The only
command-local output option is the global `--json` flag.

In UserForm `sidecar` mode the sidecar code file is measured and generated
`.frm` code is skipped. In `frm` mode embedded UserForm code is measured once.
Build artifacts, `.frx` files, `.xlflow` state, and other generated files are
not inputs. Apply `[metrics].exclude` after normal source discovery when a
project path should be omitted; this policy is independent from
`[build].exclude` and `[analyze].disabled_rules`.

## Metrics

Every procedure reports these twelve integer fields:

| Metric                  | Meaning                                                                                                                    |
| ----------------------- | -------------------------------------------------------------------------------------------------------------------------- |
| `cyclomatic_complexity` | `1 + branch_count + loop_count`.                                                                                           |
| `max_nesting_depth`     | Maximum simultaneous `If`, `Select`, loop, and `With` nesting.                                                             |
| `statement_count`       | Normalized source-backed statements; colon-separated statements are separate and continuations are one.                    |
| `source_line_count`     | Physical lines from the procedure header through matching `End`, including blank/comment/continued lines.                  |
| `branch_count`          | `If`, `ElseIf`, and each `Case` other than `Case Else`.                                                                    |
| `loop_count`            | `For`, `For Each`, `While`/`Wend`, and each `Do`/`Loop`.                                                                   |
| `goto_count`            | Explicit `GoTo` only; `On Error GoTo` is excluded.                                                                         |
| `exit_point_count`      | Implicit tail exit, `Exit Sub`/`Function`/`Property`, and standalone `End`.                                                |
| `parameter_count`       | All parameters, including Property Let/Set `value` and `ParamArray`.                                                       |
| `byref_parameter_count` | Effective `ByRef`, including omitted `ByRef`; `ByVal` is excluded and `ParamArray` follows its effective passing mode.     |
| `local_variable_count`  | Each `Dim`/`Static`/`Const` declarator, excluding parameters, fields, and synthetic return slots.                          |
| `call_fan_out`          | Unique resolved project-local callees; repeated, ambiguous, unresolved, external, member, and built-in calls are excluded. |

Single-line and multiline syntax use the same structural rules. Properties are
separate `property_get`, `property_let`, and `property_set` procedures. See the
permanent counting and schema contract in
`docs/specs/vba-procedure-complexity-metrics.md`.

## Thresholds

Thresholds are disabled by default and configured separately from analyzer
diagnostics:

```toml
[metrics]
exclude = ["src/modules/Generated/**"]

[metrics.thresholds]
cyclomatic_complexity = 10
max_nesting_depth = 4
statement_count = 0 # zero disables this threshold
```

Omitted and zero values disable a threshold. A positive threshold is exceeded
only when `value > threshold`; equal values pass. Negative values, unknown
keys, non-integers, malformed tables, absolute paths, and escaping exclusion
patterns fail as configuration errors (exit `2`). An unmatched valid exclusion
glob is a warning.

When a threshold is exceeded, `metrics` emits one metrics-specific `MX001`
warning per procedure/metric pair, keeps the complete metrics array, sets
`error.code` to `metrics_threshold_exceeded`, and exits `1`. `MX001` is not an
analyzer rule, is not controlled by `[analyze].disabled_rules`, and cannot be
inline-suppressed. With no enabled threshold, the command returns only metrics
and exits `0`.

## JSON output

```json
{
  "status": "ok",
  "command": "metrics",
  "metrics": {
    "schema_version": 1,
    "procedures": [
      {
        "file": "src/modules/Main.bas",
        "module": "Main",
        "module_kind": "standard",
        "name": "Run",
        "kind": "sub",
        "declaration_range": {
          "startLine": 1,
          "startColumn": 1,
          "endLine": 18,
          "endColumn": 7,
          "startByte": 0,
          "endByte": 256
        },
        "cyclomatic_complexity": 3,
        "max_nesting_depth": 2,
        "statement_count": 12,
        "source_line_count": 18,
        "branch_count": 1,
        "loop_count": 1,
        "goto_count": 0,
        "exit_point_count": 2,
        "parameter_count": 1,
        "byref_parameter_count": 1,
        "local_variable_count": 2,
        "call_fan_out": 2
      }
    ]
  },
  "diagnostics": [],
  "warnings": [],
  "logs": []
}
```

`metrics.schema_version` is `1`. Additive fields may appear without a version
change; removing or redefining a field requires a new version. Procedures are
sorted by normalized file, declaration start position, module, name, and kind. Threshold
diagnostics are sorted by that procedure order and the fixed metric order. This
ordering, slash-separated paths, normalized exclusion list, and stable JSON
serialization make repeated runs deterministic.

## Related

- [JSON Output](../reference/json-output)
- [Configuration file](../reference/config-file)
- [analyze](./analyze)
- [check](./check)

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow metrics` when a report, baseline, CI threshold, or editor
integration needs procedure measurements independent of diagnostic findings.

## Prerequisites

Only source files and a valid project configuration are required. Excel, the
bridge, VBIDE access, and a workbook are not required.

## What this command reads and changes

The command reads configured source roots, `tests`, and `[metrics]` settings.
It is read-only and does not modify source, workbooks, or project state.

## Effect on source-of-truth state

None. Metrics are derived from the current source tree and are not persisted by
xlflow.

## Common workflows

Run `xlflow metrics --json` in CI or an agent loop, store the versioned
`metrics` object, and opt into thresholds only when the project wants a failing
gate. Use `xlflow analyze` separately for runtime-risk diagnostics.

## Common failures

Read `error.code` and the process exit code. Invalid configuration exits `2`;
`metrics_threshold_exceeded` exits `1` while retaining all raw metrics; an
unrecoverable parse or source read returns an operation error without partial
metrics. The [JSON Output](../reference/json-output) reference defines the
failure envelope.
