# xlflow form

Manage UserForms through snapshot, build, and image export workflows.

## Usage

```bash
xlflow form snapshot <name> --out <path.yaml>`nxlflow form build <spec.yaml> [--overwrite]`nxlflow form export-image <name> --out <path.png>
```

## When to use

Use this command when its target state is the next step in the source-to-workbook workflow. Prefer `--json` for automation and AI agents.

## Example

```bash
xlflow form snapshot UserForm1 --out src/forms/specs/UserForm1.yaml --json
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
