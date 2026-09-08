# xlflow rules

List the static-analysis rules known to the installed xlflow binary. The command
is source-only: it does not load project configuration, open a workbook, or
start Excel.

## Usage

```bash
xlflow rules
xlflow rules --json
```

The human-readable table is sorted by diagnostic ID and shows the family,
severity, scope, default-enabled state, and title. Use JSON when an editor or
automation needs the complete metadata contract.

## JSON output

```json
{
  "status": "ok",
  "command": "rules",
  "rules": {
    "schema_version": 2,
    "items": [
      {
        "id": "VB001",
        "family": "lint",
        "evidence_class": "policy",
        "compile_equivalent": false,
        "default_severity": "warning",
        "supported_severities": ["warning"],
        "surfaces": ["lint", "lsp"],
        "scope": "file-local",
        "default_enabled": true,
        "configurable": true,
        "configuration_key": "require_option_explicit",
        "inline_suppressible": true,
        "preflight_blocking": false,
        "documentation_url": "https://harumiweb.github.io/xlflow/reference/diagnostics#vb001"
      }
    ]
  }
}
```

The example is abbreviated. See [JSON Output](../reference/json-output) for the
full schema and compatibility rules, and the generated [static-analysis
diagnostic catalog](../reference/diagnostics) for every rule.

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow rules` to discover stable diagnostic IDs, default behavior,
configuration bindings, suppression eligibility, or documentation links from
the exact xlflow version an integration launches.

## Prerequisites

Only the xlflow executable is required. The command works outside an xlflow
project and without Excel, VBIDE access, or a workbook.

## What this command reads and changes

The command reads the rule registry embedded in the executable. It does not
read project source or configuration and does not write files.

## Effect on source-of-truth state

None. `rules` is a read-only, parallel-safe metadata operation.

## Common workflows

- Inspect a rule before adding it to `[lint].disabled_rules` or
  `[analyze].disabled_rules`.
- Cache `rules --json` by resolved xlflow executable identity in an editor
  integration. The identity must include binary file attributes or an equivalent
  fingerprint so an in-place executable replacement invalidates stale metadata.
- Follow `documentation_url` from a diagnostic code to its catalog entry.

## Common failures

An older xlflow binary may not implement the command. Integrations must treat a
missing command, failed process, malformed response, unsupported
`schema_version`, or unknown diagnostic ID as unavailable metadata. In
particular, do not infer that an unknown rule supports inline suppression.
