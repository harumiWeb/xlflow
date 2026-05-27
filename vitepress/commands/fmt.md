# xlflow fmt

Format VBA source files with a conservative, non-destructive formatter.

## Usage

```bash
xlflow fmt [--write | --check | --diff] [--json] [--stdin] [<path>...]
```

## Options and Arguments

| Option / argument | Description                                              | Default      |
| ----------------- | -------------------------------------------------------- | ------------ |
| `<path>...`       | Files or directories to format. Defaults to project src. | project src  |
| `--write`         | Write formatted source back to files.                   | false        |
| `--check`         | Check formatting without modifying files.               | false        |
| `--diff`          | Show unified diff of formatting changes.                | false        |
| `--stdin`         | Read VBA source from stdin, write formatted to stdout.  | false        |
| `--json`          | Return structured machine-readable output.              | false        |

At most one of `--write`, `--check`, or `--diff` may be used. When none is set, `fmt` runs in inspect mode and reports which files would be changed.

## Examples

```bash
# Show which files need formatting (inspect mode)
xlflow fmt

# Write formatted source back to files
xlflow fmt --write

# Check formatting in CI (non-zero exit if unformatted)
xlflow fmt --check

# Show unified diff of changes
xlflow fmt --diff

# Format a specific file or directory
xlflow fmt --write src/modules/Main.bas
xlflow fmt --write src/modules/

# Pipe through stdin
cat MyModule.bas | xlflow fmt --stdin

# Machine-readable output
xlflow fmt --json
xlflow fmt --check --json

# Pipe through stdin with JSON envelope
cat MyModule.bas | xlflow fmt --stdin --json
```

## Notes

> [!IMPORTANT]
> `fmt` is source-only and does not open Excel COM. It targets `.bas` and `.cls` files under the configured project source directories plus `tests/`.

> [!WARNING]
> `.frm` files are skipped by default. The formatter preserves class module metadata (`Attribute VB_*`, `VERSION`, `BEGIN`/`END` blocks) verbatim.

> [!NOTE]
> `--stdin --json` writes the JSON envelope to stdout instead of formatted text. The envelope includes `output.changed` / `output.unchanged` summary fields but does not include the formatted source body.

## JSON Output

Successful `--json` output uses the xlflow envelope with command-specific fields.

```json
{
  "status": "ok",
  "command": "fmt",
  "target": {
    "kind": "source",
    "path": "src/modules, src/classes, src/workbook, tests",
    "description": "source files"
  },
  "output": {
    "mode": "inspect",
    "changed": 2,
    "unchanged": 5,
    "skipped": 1,
    "total": 8,
    "changed_paths": ["src/modules/Main.bas", "src/modules/Utils.bas"],
    "skipped_paths": ["src/forms/UserForm1.frm"],
    "skipped_reasons": [
      {"path": "src/forms/UserForm1.frm", "reason": "unsupported extension: .frm"}
    ]
  }
}
```

## Exit Codes

| Code | Meaning                           |
| ---: | --------------------------------- |
|  `0` | Success or no changes (diff mode) |
|  `1` | `--check` found unformatted files |
|  `2` | Invalid argument combination      |
|  `3` | File system or read/write failure |

## Related

- [lint](./lint)
- [check](./check)
- [error codes](../reference/error-codes)
- [exit codes](../reference/exit-codes)
- [JSON output](../reference/json-output)
