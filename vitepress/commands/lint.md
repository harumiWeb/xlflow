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

The generated [static-analysis diagnostic catalog](../reference/diagnostics) is
the authoritative reference for rule metadata, including configuration,
default state, scope, precision, preflight behavior, and inline suppression.
The summary below explains the lint findings in workflow terms.

| Code    | Severity | Description                                                                                                                          |
| ------- | -------- | ------------------------------------------------------------------------------------------------------------------------------------ |
| `VB001` | warning  | Missing `Option Explicit`.                                                                                                           |
| `VB002` | warning  | `Select` member access such as `Range("A1").Select`.                                                                                 |
| `VB003` | warning  | `Activate` member access such as `ActiveCell.Activate`.                                                                              |
| `VB004` | warning  | Broad `On Error Resume Next`.                                                                                                        |
| `VB005` | warning  | Possible implicit `Variant`, including individual untyped declarators in one `Dim` statement.                                        |
| `VB006` | warning  | Module-level `Public` variable.                                                                                                      |
| `VB007` | warning  | Automation-hostile GUI boundary such as raw dialogs, file pickers, UserForms, message pumps, or external process launches.           |
| `VB008` | error    | Typographic quote character that can trigger VBE compile dialogs.                                                                    |
| `VB009` | error    | Likely C-style quote escape in a VBA string literal.                                                                                 |
| `VB010` | error    | Unterminated `Sub`, `Function`, or `Property` procedure.                                                                             |
| `VB011` | error    | Unexpected `End Sub`, `End Function`, or `End Property`.                                                                             |
| `VB012` | error    | Mismatched procedure end statement.                                                                                                  |
| `VB013` | error    | Missing whitespace before a line-continuation underscore.                                                                            |
| `VB014` | error    | `tree-sitter-vba` recovered with an `ERROR` or `MISSING` node; this is a parser-compatibility signal, not proof that VBA is invalid. |
| `VB015` | error    | A VBA logical line uses more than 24 line-continuation characters.                                                                   |
| `VB018` | warning  | Local declarations or parameters shadow module-level names, procedure names, or same-scope declarations.                             |
| `VB019` | warning  | Multiple declarators mix typed and untyped names; in VBA each name needs its own `As <Type>`.                                        |
| `VB020` | warning  | Procedure-local variable is declared but never referenced.                                                                           |
| `VB021` | warning  | Private procedure is unreachable from known project roots; dynamic callbacks are treated conservatively.                             |
| `VB022` | warning  | Confusing parenthesized call syntax such as `Foo (bar)`.                                                                             |
| `VB023` | warning  | `For Each` control variable is undeclared or obviously incompatible.                                                                 |
| `VB026` | warning  | `Resume` is used outside a likely error-handler context.                                                                             |
| `VB027` | warning  | Nested `With` blocks use implicit Excel members whose target can be ambiguous.                                                       |
| `VB028` | error    | Bare `MsgBox` or `InputBox` appears while `XlflowUI.bas` is present; use `XlflowUI` or explicit `VBA.Interaction`.                   |
| `VB029` | error    | `Option Explicit` is present and an assignment target or loop control variable is not declared.                                      |
| `VB031` | error    | Standard `.bas` module is missing `Attribute VB_Name`.                                                                               |
| `VB032` | error    | Repeated `?` Debug.Print shorthand such as `?? "hoge"`.                                                                              |
| `VB037` | error    | Definite scalar assignment incorrectly uses the `Set` keyword; blocks source preflight.                                              |
| `VB044` | warning  | Configured local procedure-name string constant does not match its enclosing procedure name.                                         |
| `VB045` | error    | Deterministic argument-count or named-argument binding error; blocks source preflight.                                               |
| `VB046` | error    | Duplicate declaration in the same module, procedure, Enum, or Type scope; blocks source preflight.                                   |
| `VB047` | error    | Declaration appears in an invalid module/procedure position; blocks source preflight.                                                |
| `VB048` | error    | Invalid procedure parameter declaration; blocks source preflight.                                                                    |
| `VB049` | error    | Inconsistent Property Get/Let/Set accessor contract; blocks source preflight.                                                        |
| `VB050` | error    | Declaration is invalid for the canonical module kind or has an invalid WithEvents/public-member shape; blocks source preflight.      |
| `VB051` | error    | `Me` is used in a standard module; blocks source preflight.                                                                          |

Core declaration, member-access, error-handling, and procedure-scope checks are AST-backed. They ignore comments and strings, distinguish module-level declarations from procedure-local declarations, and report individual declarators such as `a` in `Dim a, b As Long`. `VB029` also resolves public declarations from standard modules across the project, so a valid project-level assignment is not reported as undeclared.

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

Safety diagnostics `VB008` through `VB015`, `VB028`, `VB029`, `VB031`, `VB032`, `VB037`, and `VB045` through `VB051` are always enabled and cannot be suppressed inline because they prevent VBE compile dialogs before `push` or `run` opens Excel.

`VB030` remains a warning for inferred or otherwise uncertain argument compatibility. `VB045` is reserved for deterministic argument binding errors confirmed by the VBE contract.

`VB046` compares declaration names case-insensitively within their containing
module, procedure, Enum, or user-defined Type. Property accessor groups are
allowed when valid; repeated accessors and Property/non-Property collisions are
errors. Repeated `Option Explicit`, `Option Base`, `Option Compare`, and
`Option Private Module` directives are also errors. `VB047` reports `Option`
statements or declarations in positions VBA does not permit while allowing
procedure-local `Dim`, `Static`, and `Const`; class-module `Implements` clauses
may precede `Option` directives. Conditional-compilation branches
are compared only when their coexistence can be proven.

`VB048` validates procedure parameter declarations against compile-time VBA
contracts, including optional/required ordering, `ParamArray`, array shape,
resolved UDT passing, parameter limits, and definitely invalid optional defaults.
Unknown or ambiguous type-dependent checks remain fail-open. `VB049` validates
the shared index/value parameter contract and return/value types of Property
Get/Let/Set accessors; unresolved type compatibility remains fail-open. Both
rules are unsuppressible errors and block source preflight. `VB050` validates
Event/Friend/Implements/WithEvents and object-module public-member placement
using canonical module metadata; unknown external type information remains
fail-open. `VB051` reports only an AST `Me` token in a standard module.

`VB014` is fail-closed for `push` and `run`, but parser recovery alone does not prove that Excel will reject the VBA. Its JSON issue may include `parser_node` (`ERROR` or `MISSING`), `parser_token`, and a short source-line `context`; when xlflow can confidently match an unclosed multiline block, it also includes `block_kind`, `expected_closer`, `opening_line`, and `opening_column`. In that case the diagnostic location marks where the closer is expected and the message identifies the opener, for example `Possible missing 'End If' for multiline If block opened at line 8.` When a parent closer is aligned exactly with its outer opener and all skipped nested openers are indented further, that parent closer is highlighted for the inner missing terminator. Conditional compilation and other ambiguous structures keep the generic recovery diagnostic. Inspect that context and validate the source in the target host before changing otherwise-valid VBA merely to satisfy parser compatibility.

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

For `VB014`, the optional recovery metadata identifies the first concrete recovery node:

```json
{
  "code": "VB014",
  "severity": "error",
  "file": "src/modules/Main.bas",
  "line": 7,
  "column": 12,
  "kind": "parser_recovery",
  "parser_node": "MISSING",
  "parser_token": "End Sub",
  "context": "Public Sub Main()",
  "message": "VBA parser recovery detected; inspect the reported source context before pushing to Excel."
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
