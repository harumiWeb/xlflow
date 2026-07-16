# Changelog

All notable changes to xlflow will be documented in this file.

## Unreleased

- Added Windows cross-process workbook coordination with crash-released
  `LockFileEx` ownership, immediate structured `workbook_busy` diagnostics, and
  guarded atomic owner metadata for conflicting CLI processes and
  WSL-delegated commands.
- Added opt-in `--wait` and `--wait-timeout` workbook coordination with a finite
  30-second default, Ctrl+C cancellation, and structured timeout/cancel errors.
- Added an observational top-level `coordination` section to `session status`,
  including current workbook busy state and guarded public owner metadata when
  available without changing existing session fields or failure behavior.
- Added persistent workbook recovery quarantine for uncertain Excel/VBA
  termination, including atomic per-workbook markers, fail-closed
  `workbook_recovery_required` diagnostics, recovery-aware status/session/process
  output, managed `session stop --discard`, and verified or force
  `recovery clear` workflows that remain distinct from the operating-system
  workbook lock.
- Documented and exhaustively tested UserForm/Designer coordination so migration,
  snapshot, build/apply, image export, form inspection/listing, pull, and push all
  converge on the same canonical workbook lock before Excel or VBIDE starts.
- Changed new-project scaffolds and `module install` to place bundled `Xlflow*.bas` helpers under `src/modules/Xlflow/`; existing root-level helper files are not moved, and `module install` now refuses those legacy collisions rather than creating duplicate VBA components.
- Fixed .NET `push` to stop before importing when Excel cannot remove an existing VBA component and to reject Excel-renamed imports, preventing duplicate modules with a `1` suffix; every removal/import failure poisons a partial session replacement, managed sessions discard the unsafe unsaved state on stop, and external sessions receive owner-correct recovery guidance.
- Fixed .NET bridge workbook persistence so transient `__XLFLOW_MODE__`, related runtime/UI/debug defined names, and generated helper modules are removed before save, including after a timed-out session run or a macro that saves and then edits again, preventing manually opened workbooks from remaining in headless mode or retaining temporary harness code.
- Added non-configurable `VB015` lint/preflight validation for VBA logical lines with more than 24 line continuations, preventing opaque Excel import failures before `push` or `run` opens Excel.
- Added strict reusable UserForm spec validation based on the canonical contract, including unknown-field, type, fixed-value, control-property, parent-reference, and type/ProgID mismatch diagnostics before `form build` opens Excel.
- Fixed `form build` so a UserForm caption is persisted through the VBComponent property path and remains consistent with runtime forms; `form export-image` now captures the actual runtime caption instead of substituting the Designer value.

## v0.23.0

- Added `xlflow form migrate sidecar` for converting imported UserForms from `frm` code-source mode to sidecar code plus Designer specs, and added `xlflow init --userform-code-source sidecar` for opting imported workbooks into the modern UserForm layout.
- Fixed non-executing UserForm Designer snapshots to capture top-level form width and height from VBComponent properties when available.
- Fixed `.NET` UserForm Designer builds so top-level form width and height from sidecar specs are applied through VBComponent properties instead of staying at Excel's default size.

## v0.22.0

- Added configurable automatic backup retention through `[backup.retention]`, disabled by default, with workbook-scoped pruning after successful backup-producing `push` and `rollback` operations.
- Expanded bundled `XlflowAssert.bas` with strict equality, `Null` / `Empty`, numeric tolerance, string, regex, array, `Range.Value2`, and object identity assertions, plus typed assertion failure formatting for terminal and JSON output.
- Added parameterized VBA tests with `@TestCase(...)`, including named cases, per-case `id` / `qualified_name`, source/runtime discovery JSON, exact case filtering, scalar literal validation, and `invalid_test_case` diagnostics.
- Added `@Skip("reason")` and `@Todo("reason")` metadata for VBA tests, including discovery `status_hint`, non-executed `skipped` / `todo` results, separate CLI summary counts, and VS Code Testing API surfacing.
- Added editor annotation completion for xlflow `@Skip(...)` and `@Todo(...)` test metadata alongside `@ExpectedError(...)`.
- Added `@ExpectedError(number[, description[, source]])` metadata for VBA tests, including source/runtime discovery JSON, expected-error execution matching, `observed_error` JSON for expected-error results, and `invalid_test_metadata` diagnostics for malformed annotations.
- Added `xlflow edit sheet add` for live-session worksheet creation, including idempotent `--if-missing` usage and positioned `--before` / `--after` insertion.
- Added `xlflow edit formula` for live-session range formula edits, including R1C1 and A1 formula assignment, event control, optional target-range calculation, and structured edit summaries for AI-agent formula workflows.
- Added `xlflow backup prune` and `xlflow backup delete --backup <id>` for explicit backup cleanup, including dry-run previews, workbook-scoped deletion, invalid/legacy cleanup flags, safe managed-root deletion checks, and structured JSON summaries.
- Added VS Code extension rollback and backup pruning workflows, including backup Quick Pick selection, active-session safeguards, post-rollback pull/inspect actions, failed-push rollback offers, and localized command/menu entries.
- Added VS Code editor support for xlflow VBA documentation comments, including Quick Fix snippet generation from `'''`, Rubberduck `@Description` annotation completions in comments, doc-comment continuation on Enter, and highlighting for `'''` comments and Rubberduck description annotations.
- Added xlflow-style documentation comments to scaffolded `Xlflow*.bas` helper modules so their public APIs show useful Hover and Signature Help documentation.
- Added stable qualified VBA test identifiers (`id` / `qualified_name`) to `xlflow test list --json` and `xlflow test --json`, allowing duplicate procedure names across modules and qualified `xlflow test --filter Module.TestName` selection.
- Added isolated temporary workbook execution for `xlflow test`, including `--isolation none|module|test`, `--no-save`, and JSON `test_run` metadata. Plain non-session test runs now protect the configured workbook by running against `.xlflow/test-runs/<run-id>/` copies.
- Added `xlflow test --fail-fast`, `--max-failures`, and `--rerun-failed`, including `not_run` results for early termination, flaky pass reporting, attempt history, and execution metadata in JSON output.
- Improved backup handling so corrupted backup entries no longer block `backup list`, valid backup JSON now includes `size_bytes`, incomplete backup creation is cleaned up, and failed `.NET` pushes report successfully created pre-push backups.
- Expanded the scaffolded `SampleTests.bas` to demonstrate smoke tags, parameterized `@TestCase(...)`, `@ExpectedError(...)`, `@Todo(...)`, common assertions, and test execution commands without introducing failing sample tests.
- Added first-class `.xlsb` support for Excel COM/VBIDE-backed VBA workflows, including `new`, `init`, source pull/push, sessions, run/test/save, backup/rollback, and UserForm operations, while keeping `.xlsm` as the default project format.
- Added stable `workbook_format_unsupported` failures for `.xlsb` on direct OOXML/file-level features such as `formulas pull`, workbook inspect, workbook cell diff, and pure-Go `pack`.
- Added `.xlam` project creation to `xlflow new`, using Excel add-in file format `55` while keeping `.xlsm` as the default, and documented `.xlam` initialization through `xlflow init`.
- Fixed `.xlam` session reuse in the .NET bridge by resolving open add-in workbooks through direct filename lookup with full-path validation, and by making VBE Compile target activation tolerate add-in workbooks without normal visible workbook windows.

## v0.21.0

- Added conservative VBA keyword and known built-in identifier casing normalization to `xlflow fmt` and LSP document formatting, enabled by default with `[fmt].keyword_casing` and `[fmt].builtin_casing`.
- Added LSP diagnostics for high-signal analyze warnings, including `VBA201`, `VBA204`, `VBA208`, `VBA209`, and `VBA212`.
- Fixed VBA LSP document formatting so incomplete or syntactically invalid buffers are skipped without surfacing an internal parser error notification in VS Code.

## v0.20.0

- Added `xlflow session attach` to adopt an already-open configured workbook as an external xlflow session, and deprecated legacy `xlflow attach --active` validation-only usage.
- Added VS Code extension actions to connect to an already-open workbook from the session menu and open the configured workbook in Excel from the Project view.
- Added onboarding warnings for disabled Excel VBA object model access in `xlflow doctor` and related setup failures, including a real temporary-workbook VBProject probe, plus VS Code prompts when `.bas`, `.cls`, or `.frm` files are not associated with xlflow's `vba` language.
- Fixed `VB009` false positives for valid VBA strings such as the official VBA-JSON `JsonConverter.bas` escaped quote output (`json_Char = "\"""`).
- Fixed VBA LSP diagnostics so `Me` is recognized as the current instance in UserForm, workbook, sheet, class, and non-xlflow fallback modules, preventing false `VB029` undeclared warnings for sheet code such as `Me.Rows(...)`.

## v0.19.0

- Added `xlflow formulas pull` to extract worksheet formulas and defined names from `.xlsx` / `.xlsm` files into deterministic region-based JSONL snapshots without launching Excel, including standalone `--src` and `--out` options, region dependency indexes, and parse status summaries.
- Added `xlflow formulas inspect` to summarize formula snapshots, list sheet/range regions, locate a cell's formula region, expand supported R1C1 patterns, and emit agent-friendly JSON.
- Added `xlflow pull --formulas` to refresh the default `formulas/` snapshot after a successful VBA pull.
- Added formula snapshot guidance to the bundled `xlflow` Agent Skill so AI agents know when to inspect `formulas/` outputs before changing VBA or workbook layout.

## v0.18.0

- Improved human-readable CLI output with clearer sections, status markers, warnings/hints, and table-style summaries while preserving `--json` contracts.
- Added `xlflow run --push` to import edited VBA source into the configured workbook before running a macro.
- Improved `xlflow run` macro-failure diagnostics to suggest `xlflow push` / `xlflow run --push` when source files are newer than the saved workbook.
- Enabled `VB020` unused local variable warnings by default; projects can opt out with `[lint].disabled_rules = ["VB020"]`.
- Updated generated `xlflow.toml` files to show how to opt out of `VB020` and how to opt in to heavier project-wide rules such as `VB021` unused private procedures.
- Improved `VB020` unused local variable detection so write-only assignments no longer count as variable references.
- Fixed VBA LSP diagnostics so `VB020` unused local variable warnings appear in editor diagnostics for the current buffer.
- Added VS Code Quick Fix actions for xlflow diagnostics to insert line-level suppression comments.
- Fixed `.NET` `xlflow run --interactive` so native VBA UI such as `MsgBox` is left for the user to dismiss instead of being misreported as `macro_failed`.
- Fixed `.NET` Excel cleanup so direct `run`/`test` executions and `session stop` do not leave owned Excel processes behind after successful commands.

## v0.17.0

- Updated `xlflow doctor` to run project-independent diagnostics successfully even when `xlflow.toml` is missing, with warnings and setup hints instead of a config failure.
- Added `xlflow update check` for structured GitHub Release update checks, and wired the VS Code extension to notify users when the installed xlflow CLI is behind the latest release.
- Added `xlflow type db status/init/refresh/clean` for global generated TypeLib databases, with an initial Excel TypeLib importer feeding LSP type intelligence.
- Added AST-aware VBA operator spacing to `xlflow fmt`, enabled by default and configurable with `[fmt].operator_spacing`.
- Added AST-aware VBA declaration spacing to `xlflow fmt`, enabled by default and configurable with `[fmt].declaration_spacing`.
- Added conservative VBA LSP `prepareRename` and `rename` support for high-confidence local, private module, private procedure, and label symbols while refusing host, TypeLib, public API, UserForm/event, ambiguous, and unresolved targets.
- Updated `xlflow lsp` to load generated TypeLib databases when present while continuing to work with only the embedded built-in database.
- Added best-effort generated TypeLib DB initialization after successful `new` and `init`; failures are reported as warnings and do not fail project creation.
- Extended generated TypeLib databases with registry-derived ProgID mappings and `--library all`, improving LSP `CreateObject("...")` late-binding inference for installed Excel, Scripting, ADODB, MSForms, Office, and VBIDE libraries.
- Changed `xlflow type db refresh` to always regenerate the generated TypeLib database; `--force` remains accepted for compatibility but is no longer required.
- Updated `xlflow doctor` to report generated TypeLib DB status and suggest initialization or refresh when the global DB is missing or stale.
- Added best-effort generated TypeLib DB initialization when `xlflow lsp --stdio` or `xlflow lsp --check` starts and the global DB is missing or stale.
- Hardened generated TypeLib DB clean/refresh behavior to reject unsafe clean targets, avoid loading stale generated files outside the manifest, and keep best-effort multi-library imports from failing the whole run when one TypeLib cannot be imported.
- Updated the VS Code extension to avoid passing the default `.xlflow/lsp.log` path in non-xlflow workspaces, so syntax/LSP-only users do not get workspace log files unless they configure one.
- Updated the VS Code extension to hide LSP CodeLens run actions in non-xlflow workspaces.
- Fixed LSP ProgID completions so `CreateObject(`, `CreateObject "..."`, and `CreateObject(Class:=...)` contexts surface late-binding candidates instead of requiring an already-open string literal.
- Improved LSP ProgID completion ranking and details by prioritizing version-independent ProgIDs and labeling versioned ProgIDs such as `Excel.Application.16` or `Forms.CommandButton.1`.
- Improved VBA LSP intelligence for common Excel idioms, including With-relative call-chain signature help, collection default-member signatures such as `ListObjects(1)`, multi-line signature help across `_` continuations, named-argument completions such as `Destination:=`, and `WScript.Shell` / `VBScript.RegExp` `CreateObject` inference.
- Added high-confidence TypeLib-powered VBA LSP diagnostics for unknown concrete members, close-name suggestions, read-only/write-only property misuse, `Set` misuse, incompatible object assignments, and no-return method value use.
- Expanded VBA LSP top-level expression completions with `Set ... = New`, `Nothing`, object-producing built-ins such as `GetObject`, and common VBA `Is*`/type-inspection functions including signature help.
- Expanded the curated VBA standard library database with string, date/time, math, conversion, file-system, interaction, and financial functions plus common `vb*` constants for completion and signature help.
- Improved VBA formatting so explicit line-continuation tails are indented one additional level while preserving existing line-number alignment behavior.
- Added `VB032` lint/LSP/preflight validation for repeated `?` Debug.Print shorthand such as `?? "hoge"`, reporting it before Excel/VBE interaction.

## v0.16.1

- Fixed `VB029` false positives for module-level declarations inside `#If ... #Else ... #End If` conditional compilation blocks.
- Fixed `VB029` false positives for Excel member chains such as `Cells(ws.Rows.Count, "A").End(xlUp).Row`, where string or numeric arguments were mistaken for undeclared receiver identifiers.

## v0.16.0

- Removed the legacy PowerShell Excel bridge for v0.16.0. The only supported bridge modes are now `auto` and `dotnet`; `--bridge powershell`, `XLFLOW_EXCEL_BRIDGE=powershell`, and `[excel].bridge = "powershell"` now fail with bridge-mode/configuration errors instead of emitting a deprecation warning.
- Added source-only `xlflow module new` and `xlflow form new` scaffolding commands for standard modules, class modules, and sidecar UserForms.
- Added source-only `xlflow module remove` and `xlflow module rename` commands for safe standard-module, class-module, and UserForm source mutations before the next `xlflow push`.
- Added explicit unknown-command reporting so mistyped commands print a stderr error with help/suggestions, and `xlflow --json <unknown-command>` returns a structured `unknown_command` failure.
- Added LSP CodeLens support for runnable no-argument VBA `Sub` procedures, including `Run` and `Run Test` actions backed by in-memory editor buffers.
- Added `xlflow test list --json` for source-only VBA test discovery, enabling editor integrations to discover tests without parsing VBA in TypeScript or opening Excel.
- Clarified the CLI JSON contract for editor integrations, including single-document stdout output and inactive `session status --json` payloads with `session.active=false`.
- Added non-configurable `VB031` lint/preflight validation requiring standard `.bas` modules to include `Attribute VB_Name`.
- Fixed VBA LSP diagnostics so the built-in `Err` object is recognized as a global, preventing false `VBA029` undeclared warnings for calls such as `Err.Raise`.

## v0.15.0

- Added initial VBA LSP signature help and argument diagnostics for common project procedures and built-in VBA/COM members, including active parameter tracking and argument-count warnings.
- Improved VBA LSP signature help for parenless calls so typing a space after calls such as `dict.Add` can show argument hints and early argument-count diagnostics.
- Expanded built-in VBA/COM signature metadata for common VBA functions, Excel range/workbook methods, WorksheetFunction calls, Scripting.FileSystemObject/TextStream, and ADODB calls used by LSP signature help and argument diagnostics.
- Hardened LSP argument diagnostics to avoid declaration-line false positives, improved diagnostic method names for signatures with return types, and documented manual signature-help smoke checks for the VS Code dev client.
- Improved VBA LSP completion and diagnostics for E2E smoke scenarios, including namespace type completion such as `Excel.`, scoped inferred member completion for `dict.`, `amountRange.`, and `rs.`, built-in completions such as `True`, `CStr`, `Now`, `Debug.Print`, and `ByVal`, With-block signature help for `.Offset(`, and out-of-scope member receiver diagnostics.
- Fixed VBA LSP member completion inside nested call expressions such as `CStr(dict.` and after With-relative call chains such as `.Offset(1,0).`.
- Fixed VBA LSP module declaration snippets so `Option Explicit` is not suggested again on the next line after it already exists.
- Suppressed empty-prefix VBA LSP completions on blank module-level lines after existing content, preventing the completion list from reopening immediately after `Option Explicit`.
- Added LSP document formatting support so VS Code Format Document can call the same VBA formatter engine used by `xlflow fmt` and receive full-document `TextEdit` results for `.bas` and `.cls` buffers.
- Improved VBA hover display with member signatures, canonical receiver types, and source/provenance notes for declarations, inferred object types, built-in globals, UserForm controls, and built-in object model members.
- Added context-aware ProgID completion inside `CreateObject("...")` while suppressing unrelated completions in ordinary string literals.
- Added active `With` block receiver tracking for VBA hover and member completion, including nested `With .Member` blocks.
- Fixed VBA statement snippets so an already completed declaration such as `Option Explicit` is not immediately suggested again.
- Expanded VBA type inference so `Set target = <known expression>` propagates the right-hand expression type into hover and completion, including method-return chains such as `FileSystemObject.OpenTextFile`.
- Strengthened Excel table collection metadata and resolution tests for `ListObjects(...).ListColumns(...).DataBodyRange` chains, including Japanese table and column names.
- Added Excel PivotTable/PivotField and Shape/TextFrame metadata so LSP hover and completion can resolve chains such as `PivotTables(...).PivotFields(...).DataRange` and `Shapes(...).TextFrame.Characters.Text`.
- Added practical `Sheets(...)` inference so worksheet-like member chains such as `Sheets("Input").Range("A1")` and `ThisWorkbook.Sheets("Input").ListObjects(...)` resolve as worksheet operations without changing the underlying `Excel.Sheets` database type.
- Expanded practical chain coverage for `FileSystemObject.OpenTextFile`, `Application.Workbooks.Open`, workbook worksheet collections, table ranges, pivot fields, and shape text frames.
- Added context-aware completion for `Call ...` and `Set x = ...`, limiting suggestions to callable symbols and object-producing expressions respectively.
- Added reusable `xlflow lsp --stdio` VBA language server support with full-document synchronization, diagnostics, document/workspace symbols, definition lookup, references, hover, completion, and a practical built-in VBA/COM type database for common Excel, MSForms, Scripting, ADODB, VBIDE, Office, and VBA constant metadata.
- Fixed LSP document symbols so incomplete VBA declarations do not return empty symbol names that VS Code rejects while editing.
- Added `.` as an LSP completion trigger character so VS Code requests member completions such as `Range("A1").Font.Color` while typing.
- Added built-in `VBA.Collection` metadata so LSP hover and member completions resolve `Dim result As Collection` and `Set result = New Collection` correctly.
- Fixed LSP type inference to prefer the nearest in-scope VBA declaration before the cursor, avoiding stale same-name declarations such as `result As Boolean` overriding a local `result As Collection`.
- Added LSP completions for module-qualified VBA procedure calls such as `Utils.BuildName` after typing `Utils.`.
- Added module-level VBA declaration snippet completions such as `Public Sub`, `Public Function`, `Dim`, `Const`, `Type`, and `Enum` while typing at the top level of a module.
- Fixed LSP document symbols for empty or incomplete VBA files so VS Code does not reject module symbols whose selection range exceeds the source range, and added identifier trigger characters so declaration snippets appear while typing.
- Fixed module-level declaration snippets so multi-word prefixes such as `Public S` are completed by replacing the typed declaration prefix instead of disappearing after a modifier is typed.
- Fixed LSP completion visibility so `Private` declarations from other modules, including `Private Const`, are not suggested as cross-module candidates.
- Added a procedure-local `Dim` snippet completion and `VB029` diagnostics for undeclared assignment targets or loop control variables when `Option Explicit` is present.
- Added VBA LSP type-position completions for declarations such as `Dim ws As Workbook`, `Function Foo() As String`, and `Set dict = New Dictionary`, including built-in VBA types, COM type aliases, and project class modules.
- Tuned VBA LSP completion and editing behavior by limiting server-side trigger characters to `.`, keeping procedure-local completion candidates scoped to the current procedure, using UTF-16 symbol selection ranges, and debouncing diagnostics after document changes.
- Added an LSP workspace symbol cache for saved source files plus open-document overlays, reducing repeated full-project symbol indexing during completion, hover, definition, and workspace symbol requests.
- Improved VBA LSP definition and reference resolution for procedure-local variables and constants so same-name locals in other procedures are no longer returned for the current local scope.
- Added VBA parameter symbols so LSP definition and reference lookup can resolve procedure arguments within the current procedure scope.
- Improved VBA LSP hover for local symbols and parameters by reusing scoped definition lookup and avoiding type inference from later declarations.
- Added UserForm `.frm` control extraction for VBA LSP intelligence, enabling hover, completion, and definition support for controls such as `Me.txtName.Text` and `Me.Controls("txtName").Text` without opening Excel.
- Expanded the built-in Excel VBA/COM type database with common formatting and worksheet helper objects such as `Excel.Font`, `Excel.Interior`, `Excel.Borders`, `Excel.Validation`, `Excel.Hyperlinks`, `Excel.PageSetup`, `Excel.AutoFilter`, `Excel.Sort`, and `Excel.WorksheetFunction`.
- Expanded built-in Excel constant metadata for common formatting, border, alignment, page orientation, and sort constants, and included enum group information in constant hover output.
- Updated LSP diagnostics to reuse the same file-local VBA lint rules as `xlflow lint` for unsaved editor buffers, including stable `VB...` diagnostic codes and diagnostic clearing when issues are fixed.
- Hardened LSP file URI and range handling for Windows paths, UNC paths, escaped Japanese paths, and UTF-16 diagnostic positions after non-ASCII text.
- Fixed LSP workspace symbols so an open editor buffer hides stale filesystem symbols for the same module, preserving the in-memory document priority used by definition and reference features.
- Updated `tree-sitter-vba` to v0.8.1 and adapted call extraction and lint member-access checks to the new stable `receiver` / `member` / `arguments` AST fields.
- Deprecated the legacy PowerShell bridge for v0.15.0. Windows `auto` bridge mode now uses the `.NET` bridge without falling back to PowerShell; explicit `--bridge powershell`, `XLFLOW_EXCEL_BRIDGE=powershell`, and `[excel].bridge = "powershell"` remain available only as deprecated opt-ins and emit `powershell_bridge_deprecated`. PowerShell bridge removal is planned for v0.16.0.
- Fixed `xlflow analyze` false positives for `VBA209` array element and UDT-array member assignments, clarified `VBA204` fallthrough guidance, and recognized paired Push/Pop `Application` state restore helpers for `VBA203`.
- Extended experimental `xlflow pack` to update an existing UserForm's code-behind when the form already exists in the template (ADR-0012 Stage 2), honoring `[userform].code_source`: `frm` reads the code from the `.frm`, while `sidecar` (the default) reads it from `src/forms/code/<FormName>.bas` and merges it in memory without writing the source tree. The form's designer storage is carried through from the template byte-for-byte and `.frx` is not read, so layout is preserved but not authored; a `.frm` whose form is not in the template fails loudly with `pack_userform_generation_unsupported`, and a sidecar carrying `Attribute VB_*` headers or with no matching `.frm` fails with `pack_ambiguous_layout`. See `docs/specs/pack-command.md`.

## v0.14.1

- Improved `xlflow pack` protected-project detection: it now reads the CMG `ProjectProtectionState` bits (MS-OVBA §2.3.1.15) instead of a DPB password-length heuristic, so protected and unprotected projects are classified by the spec-defined signal rather than a corpus-calibrated threshold.
- Added `.NET doctor` diagnostics for the Windows systemprofile Desktop directories required by non-interactive Excel COM workbook automation. Missing directories now return `systemprofile_desktop_missing` with instructions to create both `System32` and `SysWOW64` Desktop paths, while permission-denied inspection results are reported as warnings instead of false missing-directory failures. `xlflow doctor` remains lightweight by default, while the new `--workbook` option opens the configured workbook and reports `workbook_openable`.

## v0.14.0

- Added inline VBA suppression comments for lint and analyze diagnostics, supporting `xlflow:disable-next-line <ID>` and `xlflow:disable-line <ID>` with stable IDs such as `VB002` and `VBA205`, plus warnings for unknown, unsupported, or unused suppressions. Preflight-blocking errors remain unsuppressible.
- Documented COM cleanup best practices for VBA tests that open external workbooks, including `Close` plus `Set ... = Nothing` to avoid file locks during test hooks.
- Fixed test and macro discovery so Unicode VBA procedure names such as Japanese `Test*` and `*_Test` names are recognized by both PowerShell and `.NET` bridges.
- Added experimental `xlflow pack`, a pure-Go, cross-platform command that builds an `.xlsm` artifact from the source tree plus a workbook template without Excel. It regenerates `xl/vbaProject.bin` from `.bas`/`.cls` sources and replaces only that single zip entry, leaving the rest of the workbook untouched. Gated behind `--experimental`; supports standard, class, and unambiguous document modules, carries existing UserForm designer streams through byte-for-byte, and performs no VBE compile or runtime validation (every run reports `pack.vbe_validation = "not_performed"` and a `vbe_validation_skipped` warning). Fails loudly on protected or signed projects, UserForm generation, ambiguous layouts, active sessions, and in-place overwrite of the template or configured workbook. See `docs/specs/pack-command.md` and ADR-0012.
- Updated `tree-sitter-vba` to v0.7.0 and removed `xlflow fmt` parser-workaround fallback for legacy numbered comments, colon-separated block lines, and valid line-continuation forms now handled by the grammar.
- Refactored `xlflow fmt` to use `tree-sitter-vba` structure-aware indentation for supported VBA blocks while preserving comments, strings, attributes, line continuations, line-number workflows, and `.frm` skip behavior.
- Added `[lint].disabled_rules` and `[analyze].disabled_rules` for disabling configurable source feedback rules by stable diagnostic ID, with compatibility warnings for legacy per-rule booleans.
- Refactored `xlflow lint` to use `tree-sitter-vba` AST-backed checks for core declaration, member-access, and local code-shape rules, including per-declarator implicit `Variant` diagnostics and parser recovery findings.
- Refactored `xlflow analyze` to use `tree-sitter-vba` AST-backed procedure context and added runtime-risk findings for `Range.Find` `Nothing` guards, object initialization, Application state restore, error-handler fallthrough, unqualified Excel object access, ByRef mismatch candidates, Dictionary/Collection guards, `ReDim Preserve`, object/array comparisons, function return paths, and expanded Excel member mismatches.
- Documented the full `xlflow lint` rule list on the command page, including `VB001` through `VB014` codes and severity levels.
- Added `xlflow inspect calls`, a source-only tree-sitter-vba call-site extractor for exported VBA files with caller context, argument summaries, source ranges, conservative project-symbol resolution, JSON output, and compact grouped text output.
- Added `xlflow inspect symbols`, a source-only tree-sitter-vba symbol indexer for exported `.bas`, `.cls`, and `.frm` VBA files with JSON and compact outline output.
- Updated `xlflow inspect symbols` for the tree-sitter-vba 0.6.0 declaration node shape changes, including split property and declare nodes.
- Added `VB028` source preflight blocking for bare `MsgBox` / `InputBox` calls when `XlflowUI.bas` is present, so `push` fails before Excel opens with guidance to use `XlflowUI` wrappers or explicit `VBA.Interaction.*` native dialogs.

## v0.13.1

- Fixed `xlflow form snapshot` so Designer snapshots no longer require executing an injected helper macro, avoiding Trust Center / Insider Beta Office failures that blocked temporary macro workbooks from running.
- Fixed `.NET` and PowerShell Designer inspection to recover concrete UserForm control types from COM metadata when `ProgId` is unavailable, so snapshots no longer persist generic `__ComObject` / `Control` types for standard MSForms controls.
- Fixed `.NET` Designer inspection for controls such as `TabStrip` that do not expose a `Controls` collection, preventing `DISP_E_UNKNOWNNAME` failures when snapshotting forms that contain a broad set of MSForms controls.

## v0.13.0

- Added WSL development support that delegates Excel-related commands to Windows `xlflow.exe`, translates Windows-mounted project paths, preserves command streams and exit codes, and extends `doctor` with WSL/Windows diagnostics. Linux x64 release archives are now published for the WSL frontend.
- Added a GitHub Pages-hosted WSL/Linux frontend installer at `https://harumiweb.github.io/xlflow/install.sh`, including one-line `curl | sh` install guidance and `--uninstall` support.
- Added `task wsl-install`, `task wsl-uninstall`, and `task uninstall` helpers for installing or removing local Go bin xlflow binaries during delegation testing.
- Added a GitHub Pages-hosted PowerShell installer at `https://harumiweb.github.io/xlflow/install.ps1`, including review-first safer install guidance, one-line quick install guidance, and `-Action uninstall` support that removes `%LOCALAPPDATA%\xlflow` and its user PATH entry.
- Hardened `xlflow-excel-bridge.exe` direct startup so no-arg, help, version, and invalid launches exit immediately with clear output, while `xlflow.exe` uses an explicit internal run flag before starting the actual bridge runtime.

## v0.12.2

- Fix .NET bridge VBA export decoding for non-ASCII pull sources

## v0.12.1

- Fixed `.NET` bridge stdin/stdout JSON transport on Windows to use explicit UTF-8 streams, preventing mojibake and invalid bridge JSON when project, workbook, or session paths contain Japanese or other non-ASCII characters.

## v0.12.0

- Fixed: Add support for detecting implicit variants inside user-defined types (UDTs) in linter
- Reduced the default `xlflow run --json` failure payload for AI-agent and normal VBA debugging loops. The default run JSON now promotes the best available `location` and `suggestion` to top-level fields and hides verbose workbook/bridge/runtime diagnostics, dialog snapshots, and location-capture metadata unless `--verbose` is specified.
- Fixed `.NET` bridge macro runs so Excel is stabilized and the STA message loop is pumped around `Application.Run`, improving reliability for formatting/layout operations such as `Range.Interior.Color`, `Range.Clear`, row height updates, and protection state reads. Fatal COM/RPC failures such as `0x800706BE` now return `excel_com_rpc_failure` with `h_result` and run diagnostics, and live sessions are marked poisoned instead of being silently reused.
- Added best-effort `.NET` VBE selection diagnostics for suppressed compile/runtime dialogs in `xlflow run --bridge dotnet --json` and compile failures in `xlflow push`, including selected component, procedure, source path, source-file line range, selected line text, and non-fatal capture-attempt metadata when Excel exposes it.
- Improved `.NET` dialog watcher button diagnostics and action selection by capturing Win32 button `access_key`, `control_id`, and `enabled` metadata. VBA runtime suppression now prefers accelerator keys such as `D` for Debug and `E` for End before localized text fallback, improving tolerance for non-English Excel/VBE UI.
- Fixed `.NET` VBE runtime location capture after `xlflow session start` so the first failing `run` no longer reports stale module header lines such as `Option Explicit` before VBE has moved the selection to the actual error line.
- Removed the legacy runtime-debug command surface completely. VBA-internal debugging is now documented around `XlflowDebug.Log`, `xlflow run --json`, structured diagnostics, and `Erl`/line-number workflows. Legacy `XlflowLog` / `XlflowSetTraceFile` usage is now treated as removed API surface in source analysis and preflight guidance.
- Added explicit VBA line-number operations to `xlflow fmt` via `--line-numbers preserve|add|remove|renumber`, including conservative ambiguity warnings for numeric-label control flow and structured JSON summary fields under `output.line_numbers`.
- Fixed `xlflow fmt --line-numbers add` so it no longer numbers `Select Case`, `Case` / `Case Else`, or `End Select` control lines, avoiding VBA compile failures when the first `Case` in a select block is numbered.
- Fixed `xlflow fmt --line-numbers add` so explicit VBA line-continuation statements only number their first physical line; continuation tail lines now stay unnumbered to avoid compile failures.
- Added a dedicated xlflow agent debugging reference at `internal/agentskill/templates/xlflow/references/debugging.md`, documenting the preferred workflow: inspect `run` diagnostics first, then use `fmt --line-numbers add` plus targeted `XlflowDebug.Log` only when the default error metadata is not enough.

## v0.11.0

- Added native `.NET` bridge support for the remaining Windows workbook commands: `xlflow new`, `session start|status|save|stop`, `attach --active`, `runner install|status|remove`, `list forms`, `ui button add|list|remove`, and `edit cell|range|rows|columns` with explicit `--bridge dotnet --json`.
- Packaged the `.NET` Excel bridge into Windows release ZIPs as `xlflow-excel-bridge.exe` beside `xlflow.exe`, while documenting AppLocker, WDAC, AV, and unsigned-executable caveats for managed Windows environments.
- Added native `.NET` bridge support for `xlflow test --bridge dotnet --json`, enabling VBA test discovery and execution through the .NET bridge. Supports `Test*`/`*_Test` naming, `@Tag("...")` annotations, `BeforeAll`/`AfterAll`/`BeforeEach`/`AfterEach` hooks, inconclusive detection (`vbObjectError + 516`), runtime injection, UI/debug stream helpers, and session-aware workflow. Auto mode keeps the existing PowerShell behavior; use `--bridge dotnet` explicitly to route through the .NET bridge.
- Enhanced `xlflow run --bridge dotnet` with MsgBox/InputBox/FileDialog response injection (`--msgbox`, `--inputbox`, `--filedialog`), UI stream pipe support (`--ui-stream`), and `__XLFLOW_DEBUG_PIPE__` injection for `XlflowDebug.Log` transport. Previously these options were rejected by the .NET bridge; they are now fully supported with the same behavior as the PowerShell bridge.
- Fixed `.NET` runtime injection cleanup so partial defined-name injection is rolled back and execution aborts when injection cannot be completed.
- Added native `.NET` bridge support for `xlflow macros --bridge dotnet --json` and `xlflow run Module1.Main --bridge dotnet --json`, enabling macro discovery and execution through the .NET Excel bridge without PowerShell. Supports typed arguments including finite invariant-culture `double` values, fully qualified macro names, save/no-save/save-as, timeout, session attachment, and structured error handling for `macro_failed`, `macro_not_found`, and `macro_disabled`. Auto mode keeps the existing PowerShell behavior for macros/run; use `--bridge dotnet` explicitly to route through the .NET bridge.
- Added a reusable `.NET` Excel/VBE dialog watcher that captures runtime, compile, MsgBox, InputBox, and FileDialog snapshots with Win32/UI Automation identity metadata. Runtime error dialogs are suppressed without requiring Excel focus, and unattended runs prefer End over Debug to avoid leaving VBE in break mode.
- Added native `.NET` bridge support for `xlflow pull --bridge dotnet --json` and `xlflow push --bridge dotnet --json`, enabling VBA component export/import through the .NET Excel bridge without PowerShell. Auto mode keeps the existing PowerShell behavior for pull/push; use `--bridge dotnet` explicitly to route through the .NET bridge.
- Added native `.NET` bridge support for runner-backed `xlflow inspect workbook|sheets|range --session --bridge dotnet --json` and `xlflow process list|cleanup --bridge dotnet --json`, including `--bridge auto` fallback from unsupported/runtime/protocol `.NET` failures back to PowerShell for supported commands.
- Added native `.NET` `xlflow doctor --bridge dotnet --json` diagnostics for runtime and Excel COM probing, plus documentation clarifying that top-level `bridge` metadata remains provider-specific between PowerShell and `.NET` bridges.
- Added structured COM error fields (`h_result`, `details`) to `xlflow doctor --bridge dotnet --json` error output. COM activation failures now include the HRESULT hex code and exception details alongside the error message.
- Added an Excel bridge provider abstraction in Go, moved PowerShell invocation behind `PowerShellProvider`, and added bridge selection via persistent `--bridge`, `XLFLOW_EXCEL_BRIDGE`, and `[excel].bridge` while keeping `auto` on the existing PowerShell behavior for now.
- Added `xlflow fmt` as a conservative, non-destructive VBA source formatter for `.bas` and `.cls` files. Supports `--write`, `--check`, `--diff`, `--json`, and `--stdin` modes. The formatter uses 4-space indentation, strips trailing whitespace, normalizes blank lines, preserves class module metadata, and is idempotent. Typical workflow: `fmt -> lint -> push -> run/test`.
- Refined the interactive `xlflow new` / `init` welcome screen with a new `Welcome to` heading, a command reference URL, and softer muted version/info text below the ASCII logo.
- Hardened the bundled TAKT orchestra, PR review, and issue bug workflows with explicit verification, audit-triage, and release-gate handling, broader loop monitoring around remediation and final audit, and clearer guidance to treat allowed untracked files and auto-staging state as non-blocking.
- Added `xlflow process list` to enumerate all local Excel processes with PID and open-workbook status.
- Added `xlflow process cleanup <pid>`, `xlflow process cleanup --auto`, and `xlflow process cleanup --all [--yes]` for safe and forceful Excel process termination. `--auto` targets only workbook-free processes; `--all` is a destructive force-stop of all local Excel instances with mandatory interactive confirmation or `--yes`.
- Fixed `XlflowDebug.bas` helper module to stop forwarding `Log`'s `ParamArray` into a secondary helper procedure, preventing VBA compile/runtime failures such as "Sub または Function が定義されていません" and "ParamArray の使い方が適切ではありません" in some hosts.
- Fixed `.NET` `xlflow run` compile-dialog handling so VBE compile errors that surface during macro invocation are suppressed, returned as structured `vba_compile_failed` / `compile_vba` diagnostics, and no longer block unattended workflows.
- Fixed `.NET` `xlflow push` so imported VBA is VBE-compiled before saving or updating source fingerprints, matching the legacy PowerShell bridge behavior and returning structured `vba_compile_failed` / `compile_vba` diagnostics for broken source.

## v0.10.0

- Fixed `xlflow run --diagnostic` compile watcher to return structured `vba_compile_failed` errors when the VBE compile command control cannot be found, instead of silently reclassifying the failure as `vbide_access_denied`.
- Improved runtime dialog capture for `xlflow run --diagnostic` so break-mode inspection prefers user-code lines over temporary `XlflowRun_*` helpers, and the runtime macro runner now executes in a disposable child PowerShell process so break-mode resets do not leave the parent CLI hung.
- Fixed `xlflow run --diagnostic` VBE compile preflight to locate `Compile VBAProject` from the VBE menu bar (`Id = 578`) instead of assuming the Debug toolbar contains it, and to treat a disabled compile command as "already compiled" rather than a hard failure.
- Fixed `xlflow ui button add` so it auto-reuses a matching live session workbook when `.xlflow/session.json` points at the configured workbook, preventing the Excel SaveAs dialog that previously appeared when a session was active.
- Extended `ui button add`, `ui button list`, and `ui button remove` to use the shared session-aware workbook open helper and explicit save/release cleanup, matching the behavior of `push`, `pull`, `run`, and other workbook-backed commands.
- Added `xlflow status` and `xlflow status --json` as a read-only project-level command that shows project, source, workbook, and session state in one output. Source freshness is a heuristic based on file mtimes; the command does not modify workbook files, source files, or `.xlflow/state`. `workbook_saved` is now derived from `save_required` instead of `dirty` to avoid contradictory results when the session probe reports `save_required=true` but `dirty` is unknown; baseline `session` payload now always includes `running`, `workbook_open`, and `metadata` for schema stability.
- Added `xlflow init --with-module` so imported projects can immediately receive bundled runtime helper modules and sync them back into the copied workbook.
- Added `xlflow module install [--push]` so existing xlflow projects can install bundled helper modules on demand without rerunning `new`.
- Removed `--keepalive` / `--keepalive-interval` from Excel COM-backed commands and the final `XLFLOW_DONE` marker; interactive stderr now uses spinner progress where available, while non-interactive runs fall back to line-oriented progress and streamed UI/debug stderr output suppresses separate progress frames.
- Added XlflowUI module with MsgBox and InputBox wrappers to handle user prompts.
- Extended XlflowUI with headless-safe file dialog wrappers for `Application.GetOpenFilename`, `Application.GetSaveAsFilename`, open `Application.FileDialog`, and folder picker flows, plus repeated `--filedialog <kind>:<dialog-id>=<value>` CLI responses for `run` and `test`.
- Added `--ui-stream` for `xlflow run` and `xlflow test`, streaming resolved headless `XlflowUI` dialog events to stderr in real time while preserving JSON stdout and returning final `ui.events` payloads plus human-readable `UI` summaries.
- Added scaffolded `XlflowDebug` helper support so explicit `XlflowDebug.Log` calls stream to stderr and final top-level `debug` payloads during `xlflow run` and `xlflow test` without a separate CLI flag, including direct and fast run paths.
- Updated run.ps1 and test.ps1 to accept MsgBoxResponsesJSON and InputResponsesJSON parameters.
- Added explanatory comments to scaffolded `XlflowRuntime.bas`, `XlflowUI.bas`, and `XlflowAssert.bas` so workbook authors can adopt the helper modules more easily.
- Added explicit live-session inspect mode for `inspect workbook`, `inspect sheets`, `inspect range`, `inspect used-range`, and `inspect cell` via `--session`, plus explicit `live_session` target metadata and saved-file warnings that point callers to live-session inspect when disk may be stale.
- Added runtime-aware execution mode injection for `run` and `test`, plus the scaffolded `XlflowRuntime` VBA helper for branching between interactive, headless, agent, CI, and test execution contexts.
- Enhanced `xlflow macros --json` output with `component_type`, `visibility`, `has_parameters`, `runnable`, `reason_not_runnable`, and `run_command` fields per macro so AI agents and users can choose the correct entrypoint without guessing.
- Added `default_entry` and `suggestions` fields to `xlflow macros --json` output, surfaced from `project.entry` in `xlflow.toml` and resolved against discovered runnable macros.
- Added `--runnable` flag to `xlflow macros` to filter the output to only directly runnable procedures.

## v0.9.0

- Added winget release publishing so GoReleaser can generate the `HarumiWeb.Xlflow` manifest and push it to the `harumiWeb/winget-pkgs` fork for upstream submission.
- Updated `xlflow new` to bootstrap the workbook/source sync automatically by pushing scaffolded VBA into the new workbook before the command reports success, and added placeholder `src/workbook/ThisWorkbook.bas` / `Sheet1.bas` files with `Option Explicit` for new projects.
- Updated `xlflow init` to bootstrap source sync automatically by pulling VBA from the copied workbook into `src/`.
- Added first-class workbook rollback support with `xlflow backup list` and `xlflow rollback`, including metadata-backed workbook-file backups under `.xlflow/backups/<backup-id>/`, automatic safety backups before restore, and session-aware guards that refuse rollback while the target workbook is open in an active xlflow session.
- Changed default `push` backups from component-export snapshots to rollback-capable workbook snapshots, and updated the CLI/docs surface, JSON output, and VitePress command/concept pages to reflect the new backup and recovery workflow.

## v0.8.1

- Fixed `xlflow inspect form <name> --designer --session` so normal designer inspection no longer takes the strict temporary-workbook path, reducing the sample `space-invader` session inspection from about one minute to a few seconds.
- Corrected PowerShell boolean parsing and case-insensitive variable handling around the `StrictDesigner` flag, preventing `"False"` string values from being treated as truthy.
- Hardened UserForm runtime cleanup guards in `inspect form` and `form export-image` so null runtime workbook state does not trigger unnecessary Excel COM cleanup and finalizer waits.

## v0.8.0

- Completed the UserForm feature set for issue #25 across phase 1 through phase 7, including explicit UserForm warnings in core workbook flows, `xlflow list forms`, `inspect form` for designer/runtime/both, `form snapshot`, and experimental runtime image export.
- Hardened `form export-image` for real Excel GUI behavior by repairing generic runtime captions from designer state, choosing the correct monitor-relative work area instead of forcing the primary screen, using DPI-aware capture sizing, and trimming capture artifacts so the exported PNG matches the visible UserForm more faithfully.
- Corrected UserForm build round-tripping so snapshot-derived width and height no longer grow on each `form build` cycle, preserving the persisted Designer dimensions from `snapshot` output.
- Updated the bundled docs, CLI contract, and agent guidance to reflect the UserForm discovery, inspection, snapshot, export, and warning workflow, including the experimental status of runtime image export.
- Strengthened PowerShell script coverage with behavior-oriented tests for the UserForm build and export helpers, replacing narrow string-presence checks where practical.

## v0.7.0

- Added `xlflow edit cell`, `edit range`, `edit rows`, and `edit columns` as minimal workbook-mutation helpers for AI-agent testing and visual tuning in a live Excel session.
- Added session-only workbook edit behavior for the new `edit` commands, including `--events keep|on|off` support for cell value and formula changes so `Worksheet_Change` flows can be exercised without generating temporary VBA.
- Commands now display explicit workbook state, including whether reading from saved file or live Excel session
- Added warnings when live session workbooks contain unsaved changes
- Extended workbook-backed JSON and human output with explicit `target` / `session` metadata across session-aware commands, plus top-level `edit` payloads for workbook mutation summaries.
- Updated the CLI contract, README files, ADR session policy note, and bundled xlflow skill guidance to cover the new edit workflow and session-state visibility.

## v0.6.0

- Added `xlflow export-image` to export worksheet ranges as PNG images for visual verification, including session-aware targeting, structured `target` / `output` metadata, and reliability fixes so hidden-workbook captures do not produce blank images or leak Excel processes.
- Added `--include-style` flag to `inspect range` and `inspect used-range` commands to display worksheet style metadata including cell fills, borders, merged cells, row heights, and column widths.
- Added Rubberduck-compatible folder-aware VBA sync so `xlflow pull` and `push` can round-trip nested source trees via `@Folder(...)`, recursive source discovery, duplicate module-name preflight, and nested `.frm`/`.frx` companion handling.
- Added `[vba]` configuration defaults for folder sync control, wired the settings through the Go/PowerShell bridge, and documented the new contract in the CLI spec, READMEs, and bundled xlflow skill.
- Fixed folder-sync path handling to stay compatible with Windows PowerShell 5.1 and hardened `pull` so it does not clear the existing exported source tree before the workbook opens successfully.
- Added `--no-update-check` and `XLFLOW_NO_UPDATE_CHECK=1` so interactive `new` and `init` can skip the GitHub Release lookup used by the scaffold welcome banner.
- Hardened GitHub Release packaging with stable `checksums.txt` SHA256 output and archive SBOM generation via GoReleaser.
- Extended the release workflow to install Syft and publish GitHub artifact attestations for release archives, checksums, and SBOM artifacts.
- Documented Windows-side release verification in both READMEs, including SHA256 checks, `gh attestation verify`, and the current non-goal of Authenticode signing.

## v0.5.0

- Added richer sample VBA projects, including the `world-news` NewsAPI example and the `stock-price` dashboard example, plus accompanying screenshots and README updates.
- Improved runtime error handling and diagnostics so CLI runs surface failures more clearly across the Go and PowerShell execution bridge.
- Refined release documentation and sample project metadata with formatting fixes and README polish, including Japanese README badge updates.

## v0.4.0

- Added `xlflow inspect` with workbook, sheet, range, used-range, and cell inspection for saved workbook snapshots.
- Added inspect-specific formatting and range limits so agents can read workbook structure and output without opening Excel.
- Updated the bundled xlflow agent skill and CLI contract docs to teach snapshot-first inspect workflows.

## v0.3.0

- Added automatic reuse of a matching live xlflow session workbook for workbook-backed commands when `--session` is omitted.
- Added structured save-state reporting so `push`, `run`, `session status`, and related commands can surface when a live session workbook differs from disk and needs `xlflow save`.
- Improved `run` with compile-first diagnostic mode, clearer direct-run restrictions, and fallback to `project.entry` when no macro argument is provided.
- Expanded the legacy runtime-debug lifecycle handling with helper injection and session-aware workbook reuse.
- Added a verbose `version` command that reports build metadata, script resolution, supported features, and executable details.
- Added update-checking and refreshed version/welcome messaging.
- Updated bundled PowerShell scripts, agent skill guidance, and JSON envelopes to match the new session-aware behavior.

## v0.2.0

- Bundled the PowerShell scripts used by xlflow for Excel session management, testing, tracing, and UI button manipulation.
- Added the initial session-aware command surface for opening, reusing, saving, and stopping Excel workbooks.
- Added run, pull, push, test, and UI button workflows built on the bundled PowerShell bridge.
