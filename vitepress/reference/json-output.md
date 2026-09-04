# JSON Output

All commands accept the global `--json` flag.

```json
{
  "status": "ok",
  "command": "lint",
  "error": null,
  "logs": []
}
```

Failures set `status` to `failed` and populate `error`:

```json
{
  "status": "failed",
  "command": "run",
  "error": {
    "code": "macro_failed",
    "message": "Main Err 5: inputPath is required",
    "phase": "invoke_macro"
  },
  "logs": []
}
```

When another xlflow process is using the same workbook, workbook-bound commands
fail immediately with `error.code: "workbook_busy"`. `error.details` includes
`workbook`, `operation`, `resource_scope`, and `retryable`. It may also include
diagnostic `owner` metadata such as `pid`, stable command name, operation kind,
and start time. Owner metadata can be missing after a crash or during cleanup;
the error code, not the presence of `owner`, is the stable contention signal.

Retryable workbook commands may add global `--wait`. Waiting defaults to 30
seconds and can be changed with `--wait-timeout`. Timeout and cancellation use
`workbook_busy_timeout` and `workbook_busy_cancelled`; their details include
`wait_timeout`. JSON mode never writes wait progress to stderr or mixes it into
stdout, so the result remains one valid envelope.

Workbook recovery quarantine is separate from lock contention. After acquiring
the normal lock, an unsafe workbook command returns immediately with
`workbook_recovery_required` when a recovery marker exists:

```json
{
  "status": "failed",
  "command": "push",
  "error": {
    "code": "workbook_recovery_required",
    "message": "The workbook is in an uncertain Excel state after a previous operation. Explicit recovery is required before this command can run; --wait will not resolve it.",
    "phase": "coordination.recovery",
    "details": {
      "workbook": "C:\\projects\\sample\\sample.xlsm",
      "operation": "run",
      "reason": "vba_may_still_be_running",
      "recorded_at": "2026-07-16T09:30:00Z",
      "retryable": false,
      "wait_will_resolve": false,
      "recovery_actions": [
        "xlflow session stop --discard",
        "xlflow process cleanup 23456",
        "xlflow recovery clear",
        "xlflow recovery clear --force"
      ]
    }
  },
  "logs": []
}
```

`--wait` does not poll or bypass this state. A command whose own termination is
uncertain keeps its primary failure and adds top-level `recovery`:

```json
{
  "recovery": {
    "required": true,
    "published": true,
    "reason": "vba_may_still_be_running",
    "operation": "run",
    "recorded_at": "2026-07-16T09:30:00Z",
    "excel_pid": 23456,
    "worker_pid": 34567,
    "session": {
      "active": true,
      "owner": "managed"
    }
  }
}
```

PID and session fields are omitted when xlflow could not observe them.

Unknown commands are also structured when `--json` appears before the invalid command:

```json
{
  "status": "failed",
  "command": "xlflow",
  "error": {
    "code": "unknown_command",
    "message": "unknown command \"pussh\"",
    "suggestions": ["push"]
  },
  "logs": []
}
```

Command-specific fields are top-level fields such as `issues`, `analysis`,
`macro`, `macros`, `tests`, `diff`, `inspect`, `ui`, `debug`, `backups`,
`backup_prune`, `rollback`, `target`, `session`, `coordination`, `recovery`,
`warnings`, `hints`, `output`, `forms`, `edit`, `runner`, `version`, and
`capabilities`, and `build`.
`output` carries `fmt` result summaries, `export-image` output paths, and `form`
command artifacts.

## Procedure complexity metrics

`xlflow metrics --json` is a source-only command. It does not open Excel or run
`lint`, `analyze`, `check`, LSP diagnostics, or preflight. The result adds a
versioned `metrics` object. Procedure entries are project-relative, use `/`
path separators, and are sorted by file, declaration start position, module, name, and
kind. The twelve core metric fields and five additive Boolean-control fields are defined in
`docs/specs/vba-procedure-complexity-metrics.md`; the additive hotspot projection
is defined in `docs/specs/vba-procedure-and-module-hotspots.md`.

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
    ],
    "hotspots": {
      "schema_version": 1,
      "score_model": "percentile_equal_weight_v1",
      "procedures": [],
      "modules": []
    }
  },
  "diagnostics": [],
  "warnings": [],
  "logs": []
}
```

`metrics.schema_version` is `1`. Additive fields are allowed without changing
the version; removing a field or changing its meaning requires a new version.
Consumers must ignore unknown fields and should reject unsupported versions
instead of guessing their semantics. An unrecoverable source parse fails
without publishing partial procedures.

Thresholds are independent from analyzer diagnostics and are disabled unless a
positive value is configured under `[metrics.thresholds]`. `0` and omitted keys
disable a threshold; a positive threshold is exceeded only when `value >
threshold`. Each exceeded procedure/metric pair emits one warning diagnostic:

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

The `procedures` array in a real threshold result is complete; the abbreviated
example above only shortens it for readability. `diagnostics` are sorted by
procedure order and the fixed metric order. `MX001` is not a static-analysis
rule, cannot be inline-suppressed, and is unaffected by
`[analyze].disabled_rules`. No enabled thresholds means exit `0`; one or more
`MX001` entries retain the full metrics payload and use exit `1`.

### Hotspot ranking

`metrics.hotspots` is always present in a successful or threshold-failure
metrics result. It ranks procedure and module cohorts independently with the
versioned `percentile_equal_weight_v1` model. Each entity includes `rank`, a
`score` from `0` to `100`, `active_signal_count`, stable identity fields, and
both `raw_signals` and `normalized_signals`; selected entities additionally
include `selected_by` (`top_n` and/or `threshold`).

Each entity also exposes `uncertainty` counts for ambiguous, unresolved, and
dynamic calls. These counts are audit evidence only and do not contribute to
the composite score.

```json
{
  "schema_version": 1,
  "score_model": "percentile_equal_weight_v1",
  "procedures": [
    {
      "id": "src/modules/Main.bas|Main|Run|sub",
      "kind": "procedure",
      "file": "src/modules/Main.bas",
      "module": "Main",
      "module_kind": "standard",
      "name": "Run",
      "procedure_kind": "sub",
      "line": 1,
      "rank": 1,
      "score": 87.5,
      "score_model": "percentile_equal_weight_v1",
      "active_signal_count": 4,
      "raw_signals": {
        "complexity": 8,
        "call_fan_in": 3,
        "call_fan_out": 4,
        "affected_module_count": 2
      },
      "normalized_signals": {
        "complexity": 100,
        "call_fan_in": 75,
        "call_fan_out": 75,
        "affected_module_count": 100
      },
      "selected_by": { "top_n": true }
    }
  ],
  "modules": []
}
```

Raw signal names (including module-only `complexity_max` and
`public_procedure_count`) and counting boundaries are defined in
`docs/specs/vba-procedure-and-module-hotspots.md`. Average-rank percentiles
ignore constant/singleton signals, scores are equal-weight means rounded to
two decimals, and stable identity tie-breakers make ranking independent of
source enumeration. Top-N-only `MX002` diagnostics are informational;
threshold-selected entities are warning-level and use
`error.code = "metrics_hotspot_threshold_exceeded"` with exit `1`. The full
metrics and hotspot arrays remain in that failure envelope. `MX002` is not an
analyzer rule and cannot be inline-suppressed.

## Build manifest

`build --json` returns a versioned `build` manifest. Non-dry successful builds
also persist the same build evidence beside the artifact as
`<output>.build.json`.

```json
{
  "build": {
    "schema_version": 1,
    "command": "build",
    "backend": "excel",
    "base": "build/Book.xlsm",
    "output": "build/Release/Book.xlsm",
    "included_components": [
      { "name": "Main", "type": "standard", "source_path": "src/modules/Main.bas" }
    ],
    "excluded_components": [],
    "validation": {
      "source_applied": true,
      "vbe_compile": "passed",
      "workbook_saved": true,
      "workbook_closed": true
    },
    "publication": { "replaced_existing": false, "method": "atomic_create" },
    "manifest": { "path": "build/Release/Book.xlsm.build.json", "published": true, "error": null }
  }
}
```

`build.publication` describes artifact creation or replacement, while
`build.manifest.path` and `build.manifest.published` describe sidecar
publication. Dry-run uses `publication.method="not_run"` and
`manifest.published=false`. A `build_manifest_publish_failed` warning leaves
the workbook artifact successful and reports the manifest failure without
weakening the artifact replacement guarantee.

## Static-analysis rules

`xlflow rules --json` is project-independent and returns the canonical metadata
for the static-analysis rules embedded in the installed binary. Items are sorted
by `id`.

```json
{
  "status": "ok",
  "command": "rules",
  "rules": {
    "schema_version": 2,
    "items": [
      {
        "id": "VB001",
        "title": "Missing Option Explicit",
        "description": "The module does not declare Option Explicit.",
        "family": "lint",
        "category": "correctness",
        "evidence_class": "policy",
        "compile_equivalent": false,
        "default_severity": "warning",
        "supported_severities": ["warning"],
        "scope": "file-local",
        "precision": "high",
        "default_enabled": true,
        "configurable": true,
        "configuration_key": "require_option_explicit",
        "inline_suppressible": true,
        "preflight_blocking": false,
        "realtime": true,
        "fix_available": false,
        "documentation_url": "https://harumiweb.github.io/xlflow/reference/diagnostics#vb001"
      }
    ]
  }
}
```

Version 2 adds the `information` severity while retaining the catalog fields and
suppression semantics. New rules and fields may appear without changing
`schema_version`, and consumers must ignore fields they do not recognize. A
breaking field removal or meaning change requires a new version. An unavailable
command, malformed response, unsupported version, or unknown rule ID means the
metadata is unavailable; consumers must not infer suppressibility or other
safety policy. `VBA000`, `FRM...`, and `UFY...` diagnostics are intentionally
outside this registry.

## Capabilities

`xlflow capabilities --json` is project-independent and returns a versioned,
advisory view of command coordination metadata. It is generated by xlflow's
central policy registry, so integrations do not need a duplicated table of which
commands need a workbook, mutate it, execute VBA, use the UserForm Designer, or
can run concurrently.

```json
{
  "status": "ok",
  "command": "capabilities",
  "capabilities": {
    "capability_version": 1,
    "commands": {
      "push": {
        "cli_paths": ["push"],
        "resource_scope": "workbook",
        "operation_kind": "mutate",
        "parallel_safe": false,
        "retryable_when_busy": true,
        "default_wait_policy": "fail",
        "recovery_behavior": "block",
        "requires_excel": true
      }
    }
  }
}
```

Version 1 keeps command IDs, `cli_paths`, and all listed field meanings stable.
New commands and fields may be added, so consumers must ignore entries or fields
they do not recognize. Older xlflow versions may not support the command; an
unavailable, malformed, failed, or unsupported-version response means the
metadata is unavailable, not that a command is safe. The metadata is advisory:
the CLI lock remains authoritative when other processes start operations.

Top-level `status` and `session status` add `coordination` without removing
existing fields. Successful observation always distinguishes OS lock ownership
from quarantine with `busy` and `recovery_required`. A busy object may add
`resource_scope`, `operation_kind`, `command`, `pid`, and `started_at`. A
recovery object adds `reason`, previous `operation`, `recorded_at`, and optional
`excel_pid` / session ownership. Both states may briefly be true. The value is observational
command-start state. Probe failure omits the field and adds warning
`coordination_status_unavailable`.

While recovery is required, status avoids unsafe workbook COM access. Session
fields may include `dirty: null`, `source_of_truth: "uncertain"`, and
`discard_required: true`.

`backup_prune` is returned by `xlflow backup prune` and may also appear on successful `push` or `rollback` when automatic retention deleted entries, skipped invalid or legacy entries, or encountered a pruning failure. Automatic results include `"automatic": true`; pruning failures are warnings and do not change the successful workbook operation status.

`xlflow run --json` uses a compact failure payload by default. This keeps the fields that are usually relevant for fixing user VBA code and hides xlflow-internal diagnostic detail unless `--verbose` is specified.

Default failure example:

```json
{
  "status": "failed",
  "command": "run",
  "error": {
    "code": "macro_failed",
    "message": "実行時エラー '9':\n\nインデックスが有効範囲にありません。",
    "number": 9,
    "phase": "invoke_macro"
  },
  "macro": {
    "name": "Main.Run",
    "duration_ms": 1115
  },
  "location": {
    "source_path": "src/modules/QRCode.bas",
    "component": "QRCode",
    "component_type": "module",
    "procedure": "AddErrorCorrection",
    "line": 449,
    "end_line": 449,
    "text": "        dividend(i + j + 1) = dividend(i + j + 1) Xor genCoef",
    "confidence": "high",
    "method": "vbe.selection"
  },
  "session": {
    "active": true,
    "mode": "explicit",
    "dirty": true,
    "save_required": true,
    "source_of_truth": "live_workbook",
    "workbook_path": "C:\\dev\\test\\QRCode\\build\\Book.xlsm"
  },
  "target": {
    "kind": "live_session",
    "path": "C:\\dev\\test\\QRCode\\build\\Book.xlsm"
  },
  "warnings": [
    {
      "code": "save_required",
      "message": "The live workbook is newer than disk. Run xlflow save --session to persist workbook changes."
    }
  ],
  "suggestion": "Inspect src/modules/QRCode.bas:449 in AddErrorCorrection. Add targeted XlflowDebug.Log calls around the failing block and rerun."
}
```

Use `xlflow run --json --verbose` when you need xlflow-internal diagnostics such as full `run_diagnostic`, workbook/bridge/runtime metadata, dialog snapshots, or location-capture attempt details for bug reports or dialog-watcher debugging.

`run --ui-stream` and `test --ui-stream` may add a top-level `ui` object when `XlflowUI` dialogs are resolved. The stable field is `ui.events`, where each event may contain keys such as `kind`, `dialog_id`, `response_source`, `resolved_result`, `resolved_value`, `redacted`, and `error`.

Example:

```json
{
  "status": "ok",
  "command": "run",
  "ui": {
    "events": [
      {
        "kind": "msgbox",
        "dialog_id": "confirm-save",
        "response_source": "default",
        "resolved_result": "yes",
        "redacted": false
      },
      {
        "kind": "inputbox",
        "dialog_id": "customer-name",
        "response_source": "default",
        "resolved_value": "[redacted]",
        "redacted": true
      }
    ]
  }
}
```

When `--ui-stream` is enabled, xlflow also writes realtime `XlflowUI` summaries to stderr. Those streamed lines are not part of stdout JSON.

`fmt --json` returns `output` with `changed`, `unchanged`, `skipped`, and `total` summary fields. `fmt --stdin --json` returns the same envelope shape instead of formatted text; the formatted source body is not included in the JSON output.

Example:

```json
{
  "status": "ok",
  "command": "fmt",
  "output": {
    "changed": 2,
    "unchanged": 5,
    "skipped": 1,
    "total": 8
  },
  "logs": []
}
```

Source: `docs/specs/cli-contract.md`.
