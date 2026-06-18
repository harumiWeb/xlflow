# xlflow lint

Lint VBA source files for agent-hostile and compile-dialog-prone patterns.

## Usage

```bash
xlflow lint
```

## Options and Arguments

| Option / argument | Description                    | Default |
| ----------------- | ------------------------------ | ------- |
| `--json`          | Return structured lint issues. | false   |

## Examples

```bash
xlflow lint
xlflow lint --json
```

## Notes

> [!IMPORTANT]
> Syntax-safety checks are always enabled for patterns that could surface as modal VBE compile dialogs.

::: tip
Use `lint --json` in agent loops before `push` to catch source problems while Excel is still closed.
:::

## Rules

| Code    | Severity | Description                                                                                                                |
| ------- | -------- | -------------------------------------------------------------------------------------------------------------------------- |
| `VB001` | error    | Missing `Option Explicit`.                                                                                                 |
| `VB002` | warning  | `Select` member access such as `Range("A1").Select`.                                                                       |
| `VB003` | warning  | `Activate` member access such as `ActiveCell.Activate`.                                                                    |
| `VB004` | warning  | Broad `On Error Resume Next`.                                                                                              |
| `VB005` | warning  | Possible implicit `Variant`, including individual untyped declarators in one `Dim` statement.                              |
| `VB006` | warning  | Module-level `Public` variable.                                                                                            |
| `VB007` | warning  | Automation-hostile GUI boundary such as raw dialogs, file pickers, UserForms, message pumps, or external process launches. |
| `VB008` | error    | Typographic quote character that can trigger VBE compile dialogs.                                                          |
| `VB009` | error    | Likely C-style quote escape in a VBA string literal.                                                                       |
| `VB010` | error    | Unterminated `Sub`, `Function`, or `Property` procedure.                                                                   |
| `VB011` | error    | Unexpected `End Sub`, `End Function`, or `End Property`.                                                                   |
| `VB012` | error    | Mismatched procedure end statement.                                                                                        |
| `VB013` | error    | Missing whitespace before a line-continuation underscore.                                                                  |
| `VB014` | error    | `tree-sitter-vba` parser recovery found syntax errors or missing syntax nodes.                                             |

Core declaration, member-access, and error-handling checks are AST-backed. They ignore comments and strings, distinguish module-level declarations from procedure-local declarations, and report individual declarators such as `a` in `Dim a, b As Long`.

## JSON Output Example

Failed `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "failed",
  "command": "lint",
  "error": {
    "code": "lint_failed",
    "message": "1 lint issue(s) found"
  },
  "logs": [],
  "issues": [
    {
      "code": "VB005",
      "severity": "warning",
      "file": "src/modules/Main.bas",
      "line": 7,
      "column": 7,
      "message": "Declare an explicit type with As <Type>."
    }
  ]
}
```

## Related

- [analyze](./analyze)
- [check](./check)
- [error codes](../reference/error-codes)
