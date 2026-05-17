# xlflow export-image

Export a worksheet range as a PNG image through Excel COM.

## Usage

```bash
xlflow export-image [workbook] --sheet <name> --range <A1:B2> [--out <path.png>] [--overwrite] [--session]
```

## When to use

Use this command when its target state is the next step in the source-to-workbook workflow. Prefer `--json` for automation and AI agents.

## Example

```bash
xlflow export-image --sheet QR --range A1:AE31 --out artifacts/qr.png --overwrite --json
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
