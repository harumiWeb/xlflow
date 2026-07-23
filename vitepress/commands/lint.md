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
| `VB015` | error    | A VBA logical line uses more than 24 line-continuation characters.                                                         |
| `VB018` | warning  | Local declarations or parameters shadow module-level names, procedure names, or same-scope declarations.                   |
| `VB019` | warning  | Multiple declarators mix typed and untyped names; in VBA each name needs its own `As <Type>`.                              |
| `VB020` | warning  | Procedure-local variable is declared but never referenced.                                                                 |
| `VB021` | warning  | Private procedure is not called from parsed source.                                                                        |
| `VB022` | warning  | Confusing parenthesized call syntax such as `Foo (bar)`.                                                                   |
| `VB023` | warning  | `For Each` control variable is undeclared or obviously incompatible.                                                       |
| `VB026` | warning  | `Resume` is used outside a likely error-handler context.                                                                   |
| `VB027` | warning  | Nested `With` blocks use implicit Excel members whose target can be ambiguous.                                             |
| `VB028` | error    | Bare `MsgBox` or `InputBox` appears while `XlflowUI.bas` is present; use `XlflowUI` or explicit `VBA.Interaction`.         |
| `VB029` | error    | `Option Explicit` is present and an assignment target or loop control variable is not declared.                            |
| `VB031` | error    | Standard `.bas` module is missing `Attribute VB_Name`.                                                                     |
| `VB032` | error    | Repeated `?` Debug.Print shorthand such as `?? "hoge"`.                                                                    |
| `VB044` | warning  | Configured local procedure-name string constant does not match its enclosing procedure name.                               |

Core declaration, member-access, error-handling, and procedure-scope checks are AST-backed. They ignore comments and strings, distinguish module-level declarations from procedure-local declarations, and report individual declarators such as `a` in `Dim a, b As Long`.

Disable configurable lint rules with `[lint].disabled_rules`:

```toml
[lint]
disabled_rules = ["VB002", "VB006"]
```

Legacy per-rule booleans such as `forbid_select = false` remain accepted for compatibility, but xlflow emits a deprecation warning. If both formats disagree, `disabled_rules` takes precedence and xlflow reports a conflict warning.

Use inline suppression comments for intentional local exceptions while keeping rules enabled globally:

```vb
' xlflow:disable-next-line VB002
Range("A1").Select

Range("A2").Select ' xlflow:disable-line VB002
```

Multiple IDs may be listed with spaces. Unknown IDs, unsupported preflight-blocking IDs, and suppressions that no longer match a lint diagnostic are reported as warnings.

Safety diagnostics `VB008` through `VB015`, `VB028`, `VB029`, `VB031`, and `VB032` are always enabled and cannot be suppressed inline because they prevent VBE compile dialogs before `push` or `run` opens Excel.

Rules `VB019`, `VB020`, `VB022`, `VB023`, and `VB026` are enabled by default. Disable `VB020` with `disabled_rules = ["VB020"]` when a project intentionally keeps scratch locals. Heavier project-wide rules such as `detect_unused_private_procedures = true` (`VB021`) stay conservative opt-ins; new `xlflow.toml` files include commented examples. Use [`analyze`](./analyze) for semantic runtime-risk checks such as unqualified Excel access, error-handler fallthrough, Application state leaks, `Range.Find` `Nothing` guards, and object `Nothing` guards combined with dereferences in non-short-circuit boolean expressions.

To keep runtime-error diagnostics useful after procedure renames, opt into `VB044` with a local constant convention:

```toml
[lint.procedure_name_constant]
enabled = true
constant_name = "PROCEDURE_NAME"
```

The rule checks existing direct string literals only; it never inserts a missing constant or rewrites source during `xlflow lint`. It supports `Sub`, `Function`, and all `Property` procedures in standard, class, document, and UserForm modules. The LSP offers a Quick Fix that updates only the mismatched string literal.

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

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow lint` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

## Prerequisites

Check the project configuration and run `xlflow doctor --json` before workbook-backed operations. Source-only commands can run without Excel; commands that read or mutate a workbook require Windows Excel and VBIDE access.

## What this command reads and changes

The command reads the inputs and configuration described in its syntax and examples. Treat source files, the saved workbook, and a live session as separate states; add `--session` when the live workbook is authoritative. Any mutation is reversible only when a backup or explicit session save boundary exists.

## Effect on source-of-truth state

Use `xlflow status --json` before and after the command. A source edit normally requires `push`; a workbook edit normally requires `pull`; a dirty live session requires `save --session` or an intentional discard.

## Common workflows

Combine this command with the relevant [source/workbook/session workflow](../concepts/workbook-session-source), and use `--json` in scripts and agent loops.

## Common failures

Read the structured `error.code`, exit code, and recovery metadata instead of scraping terminal text. The [symptom-oriented troubleshooting guide](../help/troubleshooting) maps installation, execution, session, VS Code, and WSL failures to recovery steps.
