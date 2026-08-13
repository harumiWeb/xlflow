# VBA Procedure Complexity Metrics

This specification defines the source-only `xlflow metrics` command and its
procedure metric schema. ADR-0036 records why metrics have an independent
surface and why thresholds are not analyzer rules. The additive architectural
hotspot projection is defined by
`docs/specs/vba-procedure-and-module-hotspots.md` and ADR-0039.

## Scope and command

```text
xlflow [--json] metrics
```

`metrics` reads source and project configuration only. It never opens Excel,
uses the bridge, attaches to an Excel session, runs VBA, or invokes `lint`,
`analyze`, `check`, LSP diagnostics, or preflight findings. The command has no
command-local options in v1; `--json` is the persistent global output flag.
Human output is a stable procedure table containing the twelve core metrics and
the additive Boolean-control measurements. Machine
consumers must use the JSON contract below rather than parse the table.

The source scan includes files under configured `[src]` roots (`modules`,
`classes`, `forms`, and `workbook`) and `tests`. The normal source resolver
decides whether a file is a standard module, class, document module, or
UserForm. In `[userform].code_source = "sidecar"` mode, a matching
`src/forms/code/<FormName>.bas` is the authoritative UserForm procedure source
and the generated `.frm` code is skipped. In `frm` mode, embedded `.frm` code
is scanned once. `.frx`, build artifacts, `.xlflow` state, and other generated
files are never procedure inputs.

After source discovery, `[metrics].exclude` is applied to source files. These
patterns are project-root-relative `doublestar` globs, normalized to `/`,
deduplicated, and sorted. Absolute patterns and patterns containing `..` that
would escape the project root are invalid. An unmatched valid pattern produces
one `metrics_exclude_unmatched` warning; it is not an error. Metrics exclusions
are independent from `[build].exclude`, `[lint].disabled_rules`, and
`[analyze].disabled_rules`.

If a source cannot be read or parsed without recovery, the command fails closed
and does not publish partial procedure metrics. A failed source scan is not
converted to a threshold diagnostic.

## Metric schema

The successful result uses the normal xlflow envelope and adds:

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
        "call_fan_out": 2,
        "boolean_parameter_count": 0,
        "optional_boolean_parameter_count": 0,
        "vague_boolean_parameter_count": 0,
        "boolean_control_branch_count": 0,
        "boolean_controlled_statement_count": 0
      }
    ]
  },
  "diagnostics": [],
  "warnings": [],
  "logs": []
}
```

`file` is project-relative and uses `/` on every host. `module_kind` is the
resolved source module kind (`standard`, `class`, `document`, or `form`).
`kind` identifies the procedure declaration (`sub`, `function`,
`property_get`, `property_let`, or `property_set`). `declaration_range` is the
parser range with one-based physical `startLine`/`endLine` and
`startColumn`/`endColumn`, plus zero-based byte offsets, as exposed by the
shared AST range type. Procedure entries are
sorted by normalized `file`, declaration start position, `module`, `name`, and
`kind`; source enumeration order must not affect the result.

`schema_version` is `1` for this contract. New optional fields and additive
metrics do not change the version. Removing a field or changing the meaning or
units of an existing field requires a new version. Consumers must ignore
unknown fields and reject or quarantine unsupported schema versions rather than
guessing their meaning.

## Counting rules

The twelve core values and five Boolean-control measurements are non-negative
integers. The collector consumes normalized,
source-backed procedure IR and its conservative CFG. It does not infer facts
from diagnostic findings.

| Metric                               | Definition                                                                                                                                                                                                                                                                                |
| ------------------------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `cyclomatic_complexity`              | `1 + branch_count + loop_count`. Unknown CFG edges, Boolean `And`/`Or`, exceptions, and `GoTo` do not add complexity.                                                                                                                                                                     |
| `max_nesting_depth`                  | Maximum simultaneous nesting of `If`, `Select`, any loop, and `With`. `Else`, `ElseIf`, and `Case` stay within their parent structure. A single-line `If` contributes its control structure once.                                                                                         |
| `statement_count`                    | Count source-backed normalized IR statements. Colon-separated statements count separately; a continued physical statement counts once. Labels, declarations, control headers, and `Exit` statements count. The internal `do_condition` node of a `Do ... Loop` is not counted separately. |
| `source_line_count`                  | Physical line count from the procedure header through its matching `End Sub`, `End Function`, or `End Property`, inclusive. Blank lines, comments, and continuation lines count. A terminal newline does not create an extra line.                                                        |
| `branch_count`                       | One for every `If`, one for every `ElseIf`, and one for every `Case` other than `Case Else`. `Select Case`, `Else`, and `Case Else` themselves do not add a branch. Single-line and multiline forms use the same rule.                                                                    |
| `loop_count`                         | One for each `For`, `For Each`, `While`/`Wend`, and `Do`/`Loop` construct. A `Do While` or `Do Until` condition is part of that `Do`; it is not a second loop.                                                                                                                            |
| `goto_count`                         | Explicit `GoTo <label>` statements only. `On Error GoTo <label>`, `Resume`, `GoSub`, and numeric labels are excluded.                                                                                                                                                                     |
| `exit_point_count`                   | One implicit normal procedure-tail exit plus each explicit `Exit Sub`, `Exit Function`, or `Exit Property`, and each standalone `End`. `Exit For` and `Exit Do` are not procedure exits.                                                                                                  |
| `parameter_count`                    | Every declared procedure parameter, including `ParamArray` and the accessor `value` parameter of a Property Let/Set declaration.                                                                                                                                                          |
| `byref_parameter_count`              | Parameters whose effective passing mode is `ByRef`, including parameters where `ByRef` is omitted (VBA's default). `ByVal` is excluded; a `ParamArray` follows its effective `ByRef` passing mode.                                                                                        |
| `local_variable_count`               | Each local declarator in `Dim`, `Static`, and `Const`, including multiple declarators in one declaration. Parameters, module fields, and a synthesized Function/Property return slot are excluded.                                                                                        |
| `call_fan_out`                       | Number of unique, resolved project-local callee procedures referenced by the procedure. Repeated calls to one callee count once. Ambiguous, unresolved, external, member, and built-in calls are excluded.                                                                                |
| `boolean_parameter_count`            | Number of parameters whose effective type is exactly `Boolean`.                                                                                                                                                                                                                           |
| `optional_boolean_parameter_count`   | Number of Boolean parameters declared with `Optional`.                                                                                                                                                                                                                                    |
| `vague_boolean_parameter_count`      | Number of Boolean parameters named exactly `flag`, `mode`, or `option` (case-insensitive). Prefixes and compound names are excluded.                                                                                                                                                      |
| `boolean_control_branch_count`       | Number of `If`/`ElseIf` statements whose complete condition is one Boolean parameter, optionally wrapped in `Not` and/or parentheses. Aliases, compound `And`/`Or` expressions, and interprocedural flow are excluded.                                                                    |
| `boolean_controlled_statement_count` | Number of unique source-backed statements in descendants of directly controlled Boolean branches. Synthetic `do_condition` nodes and the branch statement itself are excluded.                                                                                                            |

The declaration, statement, and CFG facts must be normalized so equivalent
single-line and multiline syntax has identical structural metrics. Physical
source line count intentionally remains different when the source occupies a
different number of lines. CRLF and LF line endings represent the same physical
line count.

`Property Get`, `Property Let`, and `Property Set` are separate procedures and
are identified by `kind`; same-name accessors are not merged. Their parameters
and local declarations are counted independently. Labels do not increase
nesting or complexity; a label is only a statement when the normalized IR
represents it as a source-backed statement.

## Thresholds and `MX001`

Threshold configuration is separate from diagnostics:

```toml
[metrics]
exclude = []

[metrics.thresholds]
cyclomatic_complexity = 0
max_nesting_depth = 0
statement_count = 0
source_line_count = 0
branch_count = 0
loop_count = 0
goto_count = 0
exit_point_count = 0
parameter_count = 0
byref_parameter_count = 0
local_variable_count = 0
call_fan_out = 0
```

Every threshold key is an integer. Omitted keys and `0` disable that metric's
threshold. A positive threshold emits a finding only when `value > threshold`;
equal values are accepted. Negative values, non-integers, unknown threshold
keys, malformed tables, and invalid `exclude` patterns are configuration
errors (exit `2`).

When one or more enabled thresholds are exceeded, the command still returns
the complete `metrics` object and emits one warning diagnostic for each
procedure/metric pair:

```json
{
  "status": "failed",
  "command": "metrics",
  "error": {
    "code": "metrics_threshold_exceeded",
    "message": "1 procedure complexity threshold exceeded"
  },
  "metrics": {
    "schema_version": 1,
    "procedures": [],
    "hotspots": {
      "schema_version": 1,
      "score_model": "percentile_equal_weight_v1",
      "procedures": [],
      "modules": []
    }
  },
  "diagnostics": [
    {
      "code": "MX001",
      "severity": "warning",
      "file": "src/modules/Main.bas",
      "line": 1,
      "module": "Main",
      "procedure": "Run",
      "procedure_kind": "sub",
      "metric": "cyclomatic_complexity",
      "value": 8,
      "threshold": 5,
      "message": "Run cyclomatic_complexity exceeds threshold 5 (value 8)"
    }
  ],
  "warnings": [],
  "logs": []
}
```

The example abbreviates `procedures`; a real result never drops procedures
because a threshold is exceeded. `diagnostics` is sorted by procedure order and
the fixed metric order in the table above. `MX001` is metrics-specific: it is
not in the static-analysis rule registry, is not inline-suppressible, and is
not controlled by `[analyze].disabled_rules` or legacy analyzer booleans.

With no enabled threshold, a successful metrics run exits `0` and does not emit
`MX001`. With at least one threshold diagnostic, the command exits `1` and the
failure envelope retains all metrics for reporting. Source/configuration and
environment failures use the normal xlflow exit-code contract and do not emit
partial metrics.

## Determinism and consumers

The collector sorts normalized source paths, procedures, and diagnostics before
serialization. Exclusion patterns are normalized, deduplicated, and sorted;
valid unmatched patterns produce stable warnings. The same source bytes,
configuration, and tool version therefore produce identical metric structs and
JSON output regardless of filesystem enumeration order, path separator, or
line-ending convention.

Consumers should key a procedure by `(file, line, module, name, kind)` rather
than by name alone. They should preserve `schema_version`, tolerate additive
fields, and treat a missing or unsupported version as unavailable metrics.
Threshold diagnostics are suitable for CI gates, while raw values are suitable
for reports, baselines, and trend stores. Procedure/module hotspot ranking is
an additive v1 projection with its own schema and score model; see
`docs/specs/vba-procedure-and-module-hotspots.md`. Recommended thresholds,
history, and LSP projections remain outside v1.

## Related

- `docs/adr/ADR-0036-procedure-complexity-metrics.md`
- `docs/adr/ADR-0039-procedure-and-module-hotspots.md`
- `docs/specs/cli-contract.md`
- `vitepress/reference/json-output.md`
- `docs/specs/vba-analysis-ir.md`
- `docs/specs/vba-control-flow-graph.md`
- `docs/specs/vba-opaque-boolean-controls.md`
