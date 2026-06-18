# Linter

The linter catches automation-hostile VBA and syntax patterns that can block agents or surface modal dialogs.

Declaration, member-access, and error-handling checks are backed by `tree-sitter-vba`, so findings can distinguish comments and strings from executable code, module-level declarations from procedure-local declarations, and individual declarators in one `Dim` statement.

Rules include:

- Missing `Option Explicit`
- `Select` and `Activate`
- Broad `On Error Resume Next`
- Implicit `Variant` risks
- Public module fields
- GUI boundaries such as file pickers, modal dialogs, UserForms, message pumps, and external process launches
- Syntax-safety findings that prevent VBE compile dialogs before `push` or `run`

See [Error Codes](../reference/error-codes) for stable lint codes.
