# xlflow lsp

Start the reusable VBA language server for editor integrations.

## Usage

```bash
xlflow lsp --stdio
xlflow lsp --stdio --log-file .xlflow/lsp.log
xlflow lsp --stdio --performance-log
xlflow lsp --check
xlflow lsp --version
```

## Options and Arguments

| Option / argument   | Description                                                 | Default |
| ------------------- | ----------------------------------------------------------- | ------- |
| `--stdio`           | Run the LSP server over standard input/output JSON-RPC.     | false   |
| `--check`           | Validate server dependencies without starting JSON-RPC.     | false   |
| `--version`         | Print LSP server build metadata.                            | false   |
| `--log-file`        | Write server logs to this file instead of standard error.   | stderr  |
| `--performance-log` | Log structured performance measurements for LSP operations. | false   |

Exactly one of `--stdio`, `--check`, or `--version` is required.

## Notes

`xlflow lsp --stdio` reserves stdout exclusively for LSP messages. Normal logs go to stderr unless `--log-file` is provided.

Performance logging is opt-in. With `--performance-log`, each measured operation writes one structured log line to stderr or the configured `--log-file`. Records include the operation, document URI or path where applicable, version or generation, input bytes and lines, elapsed time, result count, and outcome. Diagnostics records also identify discarded obsolete results, and `diagnostics/stage` records identify the Fast/Full phase, stage duration, result count, workspace-resolution view count, and procedure IR/CFG build or reuse counts. `textDocument/didOpen` measures only the synchronous lifecycle handler; the scheduled analysis is measured separately. `workspaceSymbols/overlay` records include the overlay generation, elapsed time, outcome, and whether an obsolete result was discarded.

`xlflow lsp --check` works even before a project has an `xlflow.toml`; it validates the parser and built-in type database using default configuration.

When the global generated TypeLib database is missing or stale, `xlflow lsp --stdio` and `xlflow lsp --check` attempt best-effort generation before loading the database. Generation failures are reported on stderr and do not prevent the LSP from starting with the embedded built-in database.

The server advertises incremental document synchronization and applies ordered UTF-16 ranged edits, including CRLF and supplementary Unicode source. For ranged edits it clones and edits the previous tree-sitter tree with UTF-8 byte coordinates before parsing the new source; the previously published tree remains immutable for in-flight requests. It retains whole-document replacements as a compatibility fallback. Invalid coordinates or version ordering preserve the last valid buffer until a full replacement resynchronizes it. With performance logging enabled, each change records whether parsing was incremental, a full fallback, or retained and includes the fallback reason when applicable. It also supports diagnostics, semantic tokens, document symbols, workspace symbols, definition lookup, references, Call Hierarchy, prepare rename, rename, hover, completion, signature help, document formatting, and CodeLens. Open editor buffers are authoritative over saved filesystem content until the editor sends `textDocument/didClose`.

Call Hierarchy supports direct project-local calls to `Sub`, `Function`, and unambiguously resolved Property accessors across standard, class, document, and UserForm modules. Incoming and outgoing results use the current workspace overlay and aggregate repeated call sites. To avoid false navigation, ambiguous, unresolved, external, built-in-like, dynamic member, and module-level calls are omitted; private procedures resolve only within their module.

Rename support is conservative and scope-aware. It works for high-confidence VBA source symbols such as local variables, parameters, procedure-local constants, private module-level variables/constants, same-module private procedures, and same-procedure labels used by `GoTo`, `GoSub`, `Resume`, or `On Error GoTo`. It refuses host object model members, TypeLib/external members, public project-wide APIs, module files, `Attribute VB_Name`, UserForm controls, event handlers, property groups, ambiguous names, and unresolved identifiers.

Diagnostics reuse xlflow's file-local VBA lint rules against the current in-memory editor buffer and publish stable `VB...` codes with `source="xlflow"`. Opening a document schedules workspace-overlay and Full diagnostics analysis, but the `didOpen` handler returns without waiting. After edits, changed procedures receive Fast diagnostics after the 300 ms debounce; unchanged procedure diagnostics are safely moved with their source, and a two-second idle Full pass replaces the Fast result. Interprocedural findings remain Full-only. Changes are coalesced so a document never has more than one active analysis worker, and only the newest matching generation and highest completed phase can publish. While an overlay is pending, workspace symbols and Call Hierarchy conservatively omit stale indexed data; current open-document declarations are still available to diagnostic and document-local completion, hover, definition, and signature resolution. Superseding edits and close cancel in-progress analysis without publishing partial results. Slow documents do not prevent other documents from being analyzed. Closing a document invalidates its generation before restoring the saved workspace entry and publishes a final empty diagnostics result; an older analysis cannot publish afterward. Project-wide and filesystem-only lint checks remain available through `xlflow lint`.

`VBA237` error-suppression findings require a complete project summary and are
published only by Full diagnostics. While an open-document overlay is pending,
the LSP withholds `VBA237` rather than using the saved file or an older overlay.
When a procedure outcome or confirmed call edge changes, only open documents in
the old-or-new transitive caller set are reanalyzed. Dependency-generation
checks discard and reschedule Full results if that project evidence changes
while analysis is running, so removed edges and unsaved edits cannot leave
stale failure chains visible.

Direct YAML files under the configured `src.forms/specs` directory are also validated against the shared UserForm contract while editing. Syntax failures use `UFY001`; contract errors and support-level warnings use stable `UFV001`–`UFV014` codes with `source="xlflow"`. The server points unknown or unsupported properties at their keys, invalid values and references at their values, and missing or structural problems at the nearest affected YAML node. Files in that directory are treated as UserForm candidates even with an invalid `kind`, so `kind` itself can be diagnosed; unrelated YAML files remain ignored. JSON specifications are detected but do not yet provide editor diagnostics.

UserForm YAML completion and Hover use that same contract. Completion offers root, form, and control fields; built-in control types and matching ProgIDs; fixed values and booleans; and eligible same-document `parentId` references. Control properties are filtered by the selected type, and completion remains useful while a normal YAML edit is temporarily incomplete. Snapshot-oriented fields appear only after their name is typed. Document and control-entry snippets are also available; JSON specifications do not yet provide completion.

Hovering a UserForm YAML field, built-in control type, or built-in ProgID shows its expected value type, required status, applicability, support level, and relevant limitations. `best-effort` geometry and observed-only list state must be confirmed after Designer build; snapshot fields such as `observed`, `warnings`, and `unsupported` are capture metadata rather than normal guaranteed build inputs. Custom ProgIDs receive common structural validation only: type-specific property checks and Designer compatibility depend on the installed control.

For an enabled `VB044` procedure-name constant mismatch, diagnostics are published for the current editor buffer on open and after edits. `textDocument/codeAction` offers a `quickfix` that replaces only the direct string literal with the enclosing procedure name. This does not add missing constants or change procedure rename behavior.

The LSP publishes `VB045` for deterministic missing, excessive, duplicate, unknown, or misordered arguments when the target signature is known from project symbols or the built-in VBA/COM database. `VB030` remains the warning-level projection for inferred or uncertain argument compatibility. For project-local procedures, `VBA228` reports definite ByRef type or array-shape mismatches rejected by the VBE; temporary, parenthesized, property/member, and indirect forms remain `VBA206` warnings. `VBA229` reports unresolved procedure-local `As <Type>` identifiers using the same production resolver configuration as batch analysis. `VB045`, `VBA228`, and `VBA229` are unsuppressible compile-equivalent diagnostics and are also part of batch analysis/source preflight.

The LSP also publishes `VB046` for case-insensitive duplicate declarations in a
module, procedure, Enum, or user-defined Type scope, and `VB047` for declarations
in invalid module/procedure positions. Valid Property accessor groups and
procedure-local `Dim`, `Static`, and `Const` declarations remain accepted, and
class-module `Implements` clauses may precede `Option` directives.
Repeated `Option Explicit`, `Option Base`, `Option Compare`, and `Option Private
Module` directives are duplicate declarations. Both diagnostics use the same
unsuppressible, preflight-blocking contract as batch `lint`, and
conditional-compilation branches are compared only when their coexistence is
provable.

The LSP publishes `VB048` for compile-equivalent procedure parameter
declaration errors, including invalid array or `ParamArray` forms, optional
parameter ordering/defaults, unsupported resolved UDT passing, and excessive
parameter counts. Structural violations remain visible when type resolution is
unknown or ambiguous, while type-dependent checks fail open. `VB049` validates
Property Get/Let/Set accessor contracts, including setter value parameters,
indexed-parameter alignment, and compatible Get return/setter value types. A
Property signature edit recomputes the whole same-name accessor group so stale
cross-accessor diagnostics are not retained. Both rules are unsuppressible,
preflight-blocking errors shared with batch `lint`.

The LSP also publishes `VB050` for declarations that are invalid in the
canonical standard/class/document/UserForm context, including invalid
`WithEvents` shapes and object-module public members. `VB051` is published for
an AST `Me` token in a standard module. Both are unsuppressible,
preflight-blocking errors shared with batch `lint`; unresolved type metadata
keeps type-dependent `WithEvents` checks fail-open.

The LSP publishes `VB052` for project-local calls whose canonical resolver
proves a missing or known non-callable target, `VB053` for a bare Enum member
with multiple complete visible candidates and no lexical winner, and `VB054`
for an undeclared `RaiseEvent` target in the containing object module. These
compile-equivalent errors share the resolver snapshot and exact candidate and
visibility rules used by batch `lint`/`analyze`, Call Hierarchy, impact, and
effect propagation. Full diagnostics withhold them while the workspace index,
overlay, parser, TypeLib database, or conditional-compilation model is
incomplete; Fast diagnostics and dynamic/external/late-bound calls remain
quiet. They are unsuppressible and do not offer suppression Quick Fixes.

The LSP also publishes `VB059` for call forms that VBE rejects: explicit
`Call` statements without argument parentheses, standalone empty or
multi-argument parenthesized calls, Function calls used in expressions without
required parentheses, and syntactically identifiable invalid `Call` targets.
Legal parenthesized ByRef/ByVal argument idioms remain quiet, as does
`VB022`'s valid-but-confusing maintainability warning. `VB059` is an
unsuppressible, preflight-blocking error.

The built-in VBA/COM database includes practical Excel, MSForms, Scripting, ADODB, VBIDE, Office, and VBA constant metadata for hover, completion, and basic type inference.

Semantic tokens are provided by the Go language server with full-document `textDocument/semanticTokens/full` responses and `textDocument/semanticTokens/full/delta` updates. They classify VBA declarations, parameters, variables, built-in types, globals, constants, member expressions, comments, strings, numbers, operators, and keywords. Full responses carry opaque result IDs; the server retains the four most recent results for each open document and returns a delta only when its JSON payload is smaller than a full response. Unknown, expired, cross-document, or closed-document IDs fall back to a full response. Range semantic token requests are not advertised.

CodeLens is provided through `textDocument/codeLens` with `resolveProvider=false`. The server returns `$(play) Run` for runnable no-argument `Sub` procedures and `$(beaker) Run Test` for no-argument test procedures named `Test*`, `test*`, or `*_Test`; VS Code renders those `$(...)` prefixes as codicons. Function, Property, Declare, and argument-bearing procedures are intentionally excluded. UserForm event-style procedures are hidden by default unless the editor passes `initializationOptions.codeLens.userFormEvents=true`.

For UserForms, xlflow reads tracked `.frm` files and extracts design-time controls for code intelligence. Controls such as `txtName` in `Me.txtName.Text` and `Me.Controls("txtName").Text` can participate in hover, completion, and definition lookup without opening Excel.

## VS Code Client

The VS Code extension should launch xlflow as a thin language client:

```ts
serverOptions = {
  command: "xlflow",
  args: ["lsp", "--stdio"],
};
```

For the MVP, xlflow is resolved from `PATH`.

Set `xlflow.lsp.performanceLogging` to `true` to have the VS Code extension pass `--performance-log`. The setting defaults to `false`.

## Related

- [lint](./lint)
- [inspect symbols](./inspect)
- [JSON Output](../reference/json-output)

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow lsp` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

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
