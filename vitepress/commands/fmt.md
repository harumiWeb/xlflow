# xlflow fmt

Format VBA source files with a conservative, structure-aware formatter backed by `tree-sitter-vba`.

## Usage

```bash
xlflow fmt [--write | --check | --diff] [--line-numbers preserve|add|remove|renumber] [--json] [--stdin] [<path>...]
```

## Options and Arguments

| Option / argument | Description                                              | Default     |
| ----------------- | -------------------------------------------------------- | ----------- |
| `<path>...`       | Files or directories to format. Defaults to project src. | project src |
| `--write`         | Write formatted source back to files.                    | false       |
| `--check`         | Check formatting without modifying files.                | false       |
| `--diff`          | Show unified diff of formatting changes.                 | false       |
| `--line-numbers`  | Line-number policy for VBA statements.                   | `preserve`  |
| `--stdin`         | Read VBA source from stdin, write formatted to stdout.   | false       |
| `--json`          | Return structured machine-readable output.               | false       |

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

# Preview diagnostic VBA line numbers without writing
xlflow fmt --line-numbers add

# Apply diagnostic VBA line numbers
xlflow fmt --line-numbers add --write

# Remove diagnostic VBA line numbers
xlflow fmt --line-numbers remove --write

# Normalize existing line numbers
xlflow fmt --line-numbers renumber --write

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

## Configuration

Operator and declaration spacing are enabled by default:

```toml
[fmt]
operator_spacing = true
declaration_spacing = true
```

Set `operator_spacing = false` to keep expression/operator spacing unchanged while still applying the other formatter passes.
Set `declaration_spacing = false` to keep declaration whitespace unchanged while still applying indentation, line-number handling, and other enabled formatter passes.

## Notes

> [!IMPORTANT]
> `fmt` is source-only and does not open Excel COM. It targets `.bas` and `.cls` files under the configured project source directories plus `tests/`.
> It uses parser-backed block structure for indentation and only applies minimal text edits.
> [!NOTE]
> By default, `fmt` normalizes safe binary operator spacing such as `x=1+2` to `x = 1 + 2` and `Range("A"&i).Value=x+1` to `Range("A" & i).Value = x + 1`.
> [!NOTE]
> Operator spacing preserves named arguments such as `Filename:="C:\a.xlsx"`, type-declaration suffixes such as `Dim n&`, strings, comments, attributes, preprocessor directives, and explicit line-continuation statements. Ambiguous cases are skipped instead of rewritten.
> [!NOTE]
> Declaration spacing normalizes safe declarations such as `Dim   wb   As   Workbook` to `Dim wb As Workbook`, `Dim a As Long,b As String` to `Dim a As Long, b As String`, and `Private Function   Add(a As Long,b As Long)As Long` to `Private Function Add(a As Long, b As Long) As Long`.
> [!NOTE]
> Declaration spacing preserves type-declaration suffixes, comments, strings, attributes, preprocessor directives, `Declare` statements, fixed-length string declarations, and explicit line-continuation statements. Unsupported declaration shapes are skipped.
> [!WARNING]
> `.frm` files are skipped by default. The formatter preserves class module metadata (`Attribute VB_*`, `VERSION`, `BEGIN`/`END` blocks) verbatim.
> [!WARNING]
> Files with VBA parser errors are skipped in file-based formatting so broken source is not rewritten. With `--stdin`, parser errors return `fmt_failed`.
> [!NOTE]
> Plain `xlflow fmt` uses `--line-numbers preserve`. It does not add line numbers automatically, but it preserves existing ones where possible.
> [!NOTE]
> `--line-numbers add` skips `Select Case`, `Case` / `Case Else`, and `End Select` control lines. Only executable statements inside the case bodies receive diagnostic line numbers.
> [!NOTE]
> `--line-numbers add` also numbers only the first physical line of an explicit VBA line-continuation statement. Continuation tail lines stay unnumbered to avoid compile errors.
> [!NOTE]
> `--stdin --json` writes the JSON envelope to stdout instead of formatted text. The envelope includes `output.changed` / `output.unchanged` summary fields but does not include the formatted source body. `--stdin` cannot be combined with `--line-numbers`.
> [!NOTE]
> `fmt --line-numbers ...` still follows the normal `fmt` contract. Without `--write`, it only reports what would change. Use `--write` to persist the numbered or de-numbered source.

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
    "line_numbers": {
      "mode": "add",
      "applied": false,
      "files_to_change": 2,
      "lines_to_add": 24,
      "lines_to_remove": 0,
      "lines_to_renumber": 0,
      "warnings": []
    },
    "changed_paths": ["src/modules/Main.bas", "src/modules/Utils.bas"],
    "skipped_paths": ["src/forms/UserForm1.frm"],
    "skipped_reasons": [
      { "path": "src/forms/UserForm1.frm", "reason": "unsupported extension: .frm" }
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
