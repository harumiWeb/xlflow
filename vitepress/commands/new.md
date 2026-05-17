# xlflow new

Create a new xlflow project and macro-enabled workbook.

## Usage

```bash
xlflow new [workbook] [--with-skill] [--agent <provider>] [--no-update-check]
```

## When to use

Use this command when its target state is the next step in the source-to-workbook workflow. Prefer `--json` for automation and AI agents.

## Example

```bash
xlflow new Book.xlsm --with-skill --agent codex
```

## Output notes

JSON output uses the xlflow envelope with `status`, `command`, `error`, and command-specific top-level fields. Workbook-backed commands may also include `target`, `session`, `warnings`, and `hints`.

## Common failures

- CLI or config mistakes return exit code `2`.
- Validation, lint, macro, GUI-boundary, or test failures return exit code `1`.
- Excel, COM, VBIDE, PowerShell, or bridge failures return exit code `3`.

## Related

- [JSON output](../reference/json-output)
- [Exit codes](../reference/exit-codes)
- [Troubleshooting](../reference/troubleshooting)
