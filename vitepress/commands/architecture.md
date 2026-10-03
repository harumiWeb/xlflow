# xlflow architecture

Get one source-only overview of a VBA project's modules, entry points,
dependencies, cycles, unreachable internal code, state/effects, and hotspots.

## Usage

```bash
xlflow architecture
xlflow architecture --json
xlflow architecture --module Billing
xlflow architecture --path src/modules --json
```

## When to use this command

Start here when reviewing or editing an unfamiliar VBA project. The command
combines existing semantic facts into one snapshot. Use `analyze` for
runtime-risk diagnostics, `metrics` for optional maintainability gates,
`graph dependencies` for dependency details, and `impact` for a procedure's
caller/callee blast radius.

## Prerequisites

Source files and a valid `xlflow.toml` are required. Excel, COM, the bridge,
VBIDE access, and a workbook are not required. Optional generated TypeLib
metadata may improve resolution; missing metadata keeps uncertain calls
explicit and does not require opening Excel.

## What this command reads and changes

Reads configured source roots, conventional `tests/`, and relevant UserForm
Designer metadata. Sidecar mode measures the sidecar code once. It does not
modify files or project state. Metrics exclusions, threshold selectors, and
lint/analyzer diagnostic switches do not filter or gate this report.

## Effect on source-of-truth state

None. The report describes the captured source snapshot and is not persisted
by xlflow.

## Scope and uncertainty

`--module` is a case-insensitive exact-name display filter. `--path` selects a
source file/directory. Together they select the intersection. The entire
project is resolved first, so a filtered view preserves reachability,
fan-in/out, and global hotspot ranks. Cross-boundary dependencies and related
cycle members remain explicit context; selected counts exclude boundary nodes.

Confirmed project dependencies remain separate from ambiguous, unresolved,
external, member, built-in, and dynamic relationships. A possible root or
callback is not a confirmed static call. High hotspot scores are review leads.
Hotspot Excel-reference signals retain the metrics model; direct side-effect
evidence has its own canonical effect classification.

## Human-readable overview

The overview starts with module/procedure counts, then lists entry points,
confirmed/uncertain call counts, cyclic components, the top procedure/module
hotspots, unreachable internal procedures, state/effects, and uncertainty.
Each large list displays up to ten entries and its remaining count. Use JSON
for complete evidence.

For a project containing `Main.Run`, `Main.Unused`, and `Worker.Execute`, the
beginning of the overview looks like this:

```text
OK xlflow architecture

Project Architecture:
2 modules / 3 procedures

Entry points:
  Main.Run [confirmed] project.entry (resolved)
```

Later sections explain dependency counts, ranks, unreachable internal procedures,
and state/effect evidence. Each entry point retains its discovery reason, so the
same procedure can appear with both a configured-entry and a public-macro reason.

## JSON output

The standard envelope contains `command: "architecture"` and a versioned
`architecture` object. Its sections include `summary`, `project_summary`,
`scope`, `modules`, `procedures`, `entry_points`, `dependencies`, `cycles`,
`unreachable`, `hotspots`, mutable-state/effect evidence, and `uncertainty`.

For example, this excerpt from `architecture --module Main --json` identifies
the selected view and its source module:

```json
{
  "schema_version": 1,
  "scope": { "path": "", "module": "Main" },
  "modules": [
    {
      "id": "module|main|src/modules/Main.bas",
      "name": "Main",
      "kind": "standard",
      "file": "src/modules/Main.bas",
      "procedure_count": 2
    }
  ]
}
```

The complete object also includes whole-project counts and evidence.
`dependency_boundary_node_ids` marks retained endpoints outside the selected
view. `dynamic_references` preserves callback expressions and known targets.

`architecture.schema_version` is 1. Additive fields are compatible; removing
or changing field meanings requires a new version. Nested hotspots retain
`schema_version: 1` and `score_model: "percentile_equal_weight_v1"`. Paths use
`/`; IDs and ordering follow canonical projections. Scores/ranks describe the
entire cohort, including when the display is filtered.

## Common workflows

```bash
# Orient an agent before selecting files to edit.
xlflow architecture --json
# Focus on one module while retaining project context.
xlflow architecture --module Billing --json
# Inspect details after identifying a highly connected procedure.
xlflow impact Billing.ProcessInvoices
xlflow inspect calls --from Billing.ProcessInvoices
xlflow graph dependencies --module Billing
```

For CI reporting, persist the JSON report and its schema/model versions.
Cycles, state, uncertainty, and scores do not make architecture exit nonzero.
Configure metrics thresholds separately when a failing gate is desired.

## Common failures

Invalid configuration, unknown module, or invalid path exits 2. Source parse
recovery exits 1. Read/operation/cancellation failures exit 3. Failed envelopes
omit partial architecture data. The report's uncertainty section covers
incomplete resolution in otherwise successfully captured source.

## Related

- [JSON Output](../reference/json-output)
- [metrics](./metrics)
- [analyze](./analyze)
- [graph dependencies](./graph-dependencies)
- [impact](./impact)
