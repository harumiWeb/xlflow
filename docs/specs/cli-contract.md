# xlflow CLI Contract

## Scope

This spec defines the MVP command, configuration, JSON output, and exit-code contracts for xlflow.

xlflow is a Windows-first Go CLI that treats Excel VBA projects as source-controlled code. Excel operations use the `.NET` Excel bridge by default on Windows, with the PowerShell bridge retained as an explicit legacy fallback. Source-only commands such as `lint` should remain testable without Excel installed.

## Commands

```text
xlflow [--json] [--bridge <auto|powershell|dotnet>] new [workbook] [--with-skill] [--agent <provider>] [--no-update-check]
xlflow [--json] [--bridge <auto|powershell|dotnet>] init <workbook> [--with-module] [--with-skill] [--agent <provider>] [--no-update-check]
xlflow [--json] [--bridge <auto|powershell|dotnet>] doctor
xlflow [--json] [--bridge <auto|powershell|dotnet>] attach --active
xlflow [--json] backup list
xlflow [--json] list forms [--session]
xlflow [--json] inspect form <name> [--runtime|--designer|--both] [--initializer <method>] [--session] [--format text|json|markdown]
xlflow [--json] form snapshot <name> --out <path.json|path.yaml|path.yml> [--session]
xlflow [--json] form build <spec.json|spec.yaml|spec.yml> [--overwrite] [--session] [--no-save]
xlflow [--json] form export-image <name> --out <path.png> [--initializer <method>] [--overwrite] [--session]
xlflow [--json] pull [--session]
xlflow [--json] push [--backup always|never] [--fast] [--changed-only] [--session] [--no-save]
xlflow [--json] rollback (--latest | --backup <backup-id>)
xlflow [--json] session start
xlflow [--json] session status
xlflow [--json] session stop
xlflow [--json] save [--session]
xlflow [--json] runner install
xlflow [--json] runner remove
xlflow [--json] runner status
xlflow [--json] run [macro] [--input <workbook>] [--arg <type:value>]... [--msgbox <dialog-id=result>]... [--inputbox <dialog-id=value>]... [--filedialog <kind>:<dialog-id>=<value>]... [--ui-stream] [--save | --save-as <path>] [--headless | --interactive] [--direct] [--fast] [--diagnostic] [--session] [--timeout <duration>]
xlflow [--json] export-image [workbook] --sheet <name> --range <A1:B2> [--out <path> | --output-dir <dir>] [--name <filename>] [--format png] [--overwrite] [--session]
xlflow [--json] edit cell [workbook] --sheet <name> --cell <A1> (--value <text> | --formula <formula> | --fill <#RGB|#RRGGBB>) --session [--events keep|on|off]
xlflow [--json] edit range [workbook] --sheet <name> --range <A1:B2> (--fill <#RGB|#RRGGBB> | --clear contents|formats|all) --session
xlflow [--json] edit rows [workbook] --sheet <name> --rows <1:31> --height <points> --session
xlflow [--json] edit columns [workbook] --sheet <name> --columns <A:AE> --width <chars> --session
xlflow [--json] macros [--session]
xlflow [--json] ui button add --sheet <name> --cell <A1> --text <caption> --macro <module.proc> [--id <id>] [--width <points>] [--height <points>] [--create-sheet] [--verify-macro] [--session]
xlflow [--json] ui button list [--sheet <name>] [--session]
xlflow [--json] ui button remove --id <id> [--sheet <name>] [--session]
xlflow [--json] test [--filter <name>] [--module <name>] [--tag <tag>] [--msgbox <dialog-id=result>]... [--inputbox <dialog-id=value>]... [--filedialog <kind>:<dialog-id>=<value>]... [--ui-stream] [--session]
xlflow [--json] diff <before-workbook> <after-workbook> [--vba-before <dir>] [--vba-after <dir>]
xlflow [--json] inspect workbook [--session] [--format text|json|markdown]
xlflow [--json] inspect sheets [--session] [--format text|json|markdown]
xlflow [--json] inspect range [<sheet!A1:B2>] [--sheet <name> --address <A1:B2>] [--max-rows <n>] [--max-cols <n>] [--include-style] [--session] [--format text|json|markdown]
xlflow [--json] inspect used-range [<sheet>] [--sheet <name>] [--max-rows <n>] [--max-cols <n>] [--include-style] [--session] [--format text|json|markdown]
xlflow [--json] inspect cell [<sheet!A1>] [--sheet <name> --address <A1>] [--session] [--format text|json|markdown]
xlflow [--json] inspect-gui
xlflow [--json] lint
xlflow [--json] fmt [--write|--check|--diff] [--line-numbers <preserve|add|remove|renumber>] [--stdin] [<path>...]
xlflow [--json] analyze
xlflow [--json] check
xlflow [--json] generate test <module-name>
xlflow [--json] module install [--push]
xlflow [--json] process list
xlflow [--json] process cleanup <pid>
xlflow [--json] process cleanup --auto
xlflow [--json] process cleanup --all [--yes]
xlflow [--json] skill install [--agent <provider> | --target <dir>] [--force]
xlflow [--json] version [--verbose]
```

`--json` is a persistent global flag and can be used with every command, including `new` and `init`.

`--bridge` is also a persistent global flag for Excel bridge-backed commands. Supported values are `auto`, `powershell`, and `dotnet`. Resolution order is `--bridge`, then `XLFLOW_EXCEL_BRIDGE`, then `[excel].bridge`, then the default `auto`. On Windows, `auto` prefers the `.NET` bridge and falls back to PowerShell only when `.NET` is unavailable, incompatible, or unsupported for the requested command. Explicit `--bridge dotnet` is strict and does not implicitly fall back to PowerShell. Explicit `--bridge powershell` always uses the legacy PowerShell bridge.

When `--json` is not set, output is optimized for humans rather than machines. Interactive terminals may use Bubble Tea/Lipgloss presentation, color, and progress spinners for Excel COM-backed commands. Non-interactive output, such as CI logs and pipes, stays static and text-oriented while preserving the same command result information. Machine consumers must use `--json` instead of parsing human output.

Excel COM-backed commands report progress on stderr. Interactive stderr terminals may show Bubble Tea/Lipgloss spinner output, while `--json` or non-interactive runs fall back to line-oriented stderr progress so stdout remains a single final human result or JSON envelope. Commands that stream UI or debug events to stderr may suppress separate progress output. Agents should wait for process exit instead of parsing interim stderr progress text.

`run --ui-stream` and `test --ui-stream` add a second stderr-only channel for headless `XlflowUI` activity. When enabled, xlflow streams resolved `XlflowUI.MsgBox`, `XlflowUI.InputBox`, and `XlflowUI` file dialog wrapper events to stderr in real time as lines such as `xlflow: ui kind=msgbox id=confirm-save source=default result=yes` or `xlflow: ui kind=file-open id=input-files source=scripted value=C:\temp\a.txt | C:\temp\b.txt`. This stream never writes to stdout, so `--json` stdout remains valid. When `--ui-stream` is enabled, final command results also include the structured `ui.events` payload. InputBox values are redacted by default in both the streamed stderr lines and the final event payload.

`run` and `test` also enable terminal streaming for explicit `XlflowDebug.Log(...)` calls by default. New scaffolded projects include `src/modules/XlflowDebug.bas`, which mirrors its message to the normal VBA Immediate Window and, during xlflow execution, also writes realtime stderr lines such as `xlflow: debug source=XlflowDebug.Log mode=headless message=starting run` without requiring an extra CLI flag. This stream never writes to stdout, so `--json` stdout remains valid. Final command results may include top-level `debug.events`, plus `debug.count` and `debug.truncated` when xlflow retained only the most recent debug lines.

Excel COM-backed commands also include top-level `bridge` metadata identifying the xlflow Excel bridge process. The fields vary by bridge provider. The PowerShell bridge returns `host` (process name, e.g. `pwsh.exe`), `edition` (e.g. `Core` or `Desktop`), and `version` (PowerShell version). The .NET bridge returns `name`, `version`, `protocol_version`, `runtime`, `architecture`, and may also include `commit`. If workbook VBA launches its own external PowerShell process, that workbook-side host may differ and must be inspected separately.

`doctor --json` adds a top-level `diagnostics` object. Provider-specific `bridge` metadata and `diagnostics` serve different purposes: `bridge` identifies the bridge process itself, while `diagnostics` describes the probed Excel/runtime environment. `diagnostics.requested_bridge` records the requested mode, `diagnostics.selected_bridge` records the resolved provider, `diagnostics.fallback` reports whether `auto` fell back to PowerShell, and `diagnostics.legacy` reports whether the final provider is the legacy PowerShell bridge. With `.NET`, successful output also includes `diagnostics.protocol_version`, nested `runtime`, and nested `excel` probe fields. With PowerShell, callers must not assume the same nested shape unless the provider contract explicitly documents it for that bridge.

`new` creates a fresh macro-enabled workbook under `build/`, scaffolds the same project layout as `init`, and then automatically `push`es the scaffolded VBA source into that workbook so the initial workbook and `src/` tree start in sync. Without an argument it creates `build/Book.xlsm`; when the argument has no extension, `.xlsm` is appended. Any other extension is rejected because workbook creation always uses Excel macro-enabled format `52`. New projects write `[userform].code_source = "sidecar"` into `xlflow.toml`.

`init` accepts an existing workbook path, copies that workbook into the new project's `build/<basename>` path, records that project-local `build/...` path in `xlflow.toml` under `[excel].path` (for example `build/Sales.xlsx`), and then automatically `pull`s VBA from the copied workbook into `src/` so the initial source tree reflects workbook reality without a second command. Initialized projects write `[userform].code_source = "frm"` so existing `.frm`-embedded code remains authoritative by default. Because this bootstrap pull opens Excel/VBIDE, `init` is now an Excel COM-backed command and returns an environment failure if the automatic import cannot complete. `init --with-module` additionally installs `XlflowAssert.bas`, `XlflowRuntime.bas`, `XlflowUI.bas`, and `XlflowDebug.bas` into the configured module source root after the bootstrap pull, then automatically `push`es those helper modules into the copied workbook so source and workbook finish in sync. `init --with-module` refuses to overwrite existing helper source files.

`new` and `init` create or update a project-local `.gitignore`. The managed entries ignore Excel temporary files (`~$*.xls*`, `*.tmp`) and xlflow-generated state (`.xlflow/`, `build/`). Existing `.gitignore` content is preserved; missing managed entries are appended without duplicating entries that are already present.

`module install` installs the bundled helper modules `XlflowAssert.bas`, `XlflowRuntime.bas`, `XlflowUI.bas`, and `XlflowDebug.bas` into the configured `[src].modules` root of an existing xlflow project without changing the workbook by default. `module install --push` additionally imports those helper modules into the configured workbook through the normal `push` path. The command refuses to overwrite any existing target helper source file.

`generate test <module-name>` creates a new test module file under the configured `[src].modules` directory. The generated file includes the standard module header, `Option Explicit`, lifecycle hook stubs (`BeforeAll`, `AfterAll`, `BeforeEach`, `AfterEach`), and a sample test sub. The command fails if a file with the same name already exists. The module name must not include the `.bas` extension.

`new` and `init` do not create `prompts/agent.md`. Use `--with-skill` to install the bundled `xlflow` AI agent skill during project creation. `--agent` selects one of `agents`, `codex`, `claude`, `cursor`, or `gemini`. When `--with-skill` is used without `--agent` in an interactive terminal, xlflow opens a Bubble Tea provider selector. With `--json` or non-interactive input, `--agent` is required.

Interactive `new` and `init` runs may show a welcome banner that checks the latest GitHub Release via the GitHub Releases API. `--no-update-check` disables that network request for the current invocation. Setting `XLFLOW_NO_UPDATE_CHECK=1` also disables it. JSON and non-interactive runs do not render the welcome banner and do not perform this update check.

`skill install` installs the bundled `xlflow` skill without creating or changing an xlflow project scaffold. Provider targets are:

- `agents`: `.agents/skills/xlflow`
- `codex`: `.codex/skills/xlflow`
- `claude`: `.claude/skills/xlflow`
- `cursor`: `.cursor/skills/xlflow`
- `gemini`: `.gemini/skills/xlflow`

For GitHub Copilot, use `agents` because Copilot reads repository instructions from `.agents`. `--target <dir>` installs to `<dir>/xlflow` instead of a provider default. `--agent` and `--target` cannot be combined. Existing skill directories are not overwritten unless `--force` is set. If neither `--agent` nor `--target` is provided, interactive terminals use the Bubble Tea provider selector; `--json` and non-interactive runs return a configuration error instead.

`version` reports build metadata. `version --verbose` additionally includes the resolved executable path, Go/build information when available, embedded-versus-override PowerShell script resolution details, and a supported-feature list. This command does not require Excel COM.

`list forms` opens the configured workbook through the same session-aware Excel bridge as other workbook-backed read commands and lists `VBProject.VBComponents` entries whose type is `3` (`vbext_ct_MSForm`). Each returned form includes `name`, `component_type="MSForm"`, `has_frx`, `source_path`, and optional `frx_path`. The source paths are project-relative expected source-tree locations resolved with the same folder-aware path logic as `pull`, not proof that a `.frm` export already exists. `list forms` does not load UserForms at runtime and therefore does not execute `UserForm_Initialize`. When VBProject access is blocked, the command returns `vbproject_access_denied` with guidance to enable Trust Center access.

`inspect form` opens the configured workbook through Excel COM and returns structured UserForm state instead of file-based workbook cells. The command supports three basis modes: `runtime`, `designer`, and `both`; the default is `runtime`. Runtime inspection runs against a temporary workbook copy created from the current source workbook state, so loading the form, executing `UserForm_Initialize`, and invoking an explicit initializer do not mutate the source workbook or attached live session. Designer inspection reads `VBProject.VBComponents(name).Designer` directly from the source workbook without loading the form at runtime or requiring workbook VBA to execute. `--both` returns both snapshots in one response. `--initializer <method>` is optional and is valid only with `runtime` or `both`; when present, xlflow invokes that public form method with `ThisWorkbook` inside the temporary runtime copy after runtime load and before control enumeration. Runtime inspection always warns that it executed `UserForm_Initialize`, and it adds warnings for temporary-copy execution and explicit initializer invocation when applicable. The top-level inspect payload uses `inspect.target="form"` and `inspect.source="excel_com"`. Single-basis results use `inspect.form`; `--both` uses `inspect.forms.runtime` and `inspect.forms.designer`. Control snapshots include name/type and common UI state such as caption, text/value, geometry, enabled/visible state, selected index, and list items when the control exposes them. Missing forms fail with `form_not_found`; invalid flag combinations fail with `inspect_form_args_invalid`.

`form snapshot` is a strict designer persistence command. Unlike `inspect form --designer`, it opens a temporary workbook copy and runs an injected VBA helper so the persisted spec can capture concrete control types suitable for future rebuild workflows. `--out` is required; there is no `--designer` flag because snapshot is fixed to strict designer mode. `.json` writes JSON; `.yaml` and `.yml` write YAML; any other extension fails with `form_snapshot_args_invalid` before Excel opens. Because the command uses the helper path, workbook VBA must be executable enough to run that helper. The persisted file always uses `schemaVersion`, `kind`, `basis`, `coordinateSystem`, `form`, `controls`, and `warnings`, and converts inspect snake_case keys such as `prog_id`, `tab_index`, and `selected_index` into camelCase spec fields. Persisted `controls` are a flat array with stable `id`, optional `parentId`, and optional `zIndex`; nested control trees are reconstructed from those references instead of persisting recursive `controls` arrays as the canonical source of truth. Explicit duplicate control `id` values are validation errors rather than auto-corrected because `id` and `parentId` are the canonical structural references. `form.observed` keeps raw captured form state, and `form.build` keeps the rebuild intent when present. Persisted `warnings` are limited to form-local snapshot warnings; operational command warnings such as `save_required` or helper cleanup issues remain only in the command result. The recommended canonical source-controlled artifact for UserForm design is `src/forms/specs/*.yaml`. `form snapshot` does not emit code-behind; in `sidecar` mode, `pull` is the command that captures `src/forms/code/*.bas`. Exported `.frm` / `.frx` files remain pull/build artifacts rather than the primary Designer source of truth. Successful command JSON includes the normal workbook/session metadata plus `forms` summary metadata and top-level `output.path` / `output.format`. Like `inspect form`, `form snapshot` may auto-reuse a matching recorded session workbook when `--session` is omitted.

`form build` creates a Designer-backed workbook UserForm from a persisted `xlflow.userform` spec. The spec path must end with `.json`, `.yaml`, or `.yml`, and xlflow validates `schemaVersion=1`, `kind="xlflow.userform"`, `basis="designer"`, `form.name`, control `id`, parent references, and supported control types before Excel opens. Parse failures return `spec_parse_failed`; schema-level top-level failures return `spec_schema_invalid`; structural spec problems such as duplicate `id` values, missing `parentId` targets, or unsupported control types return `spec_validation_failed`. These failures also include top-level `spec` metadata such as `path`, `format`, optional `line`, optional `column`, optional `field`, and optional `suggestion`. In `sidecar` mode, xlflow also runs the same source preflight used by `push`/`run` before Excel opens and first synchronizes tracked `.frm` embedded code from `src/forms/code/<FormName>.bas` so the sidecar remains authoritative. If a sidecar contains exported UserForm header lines such as `Attribute VB_Name`, preflight fails with `source_preflight_failed` before Excel opens because sidecar files must contain only code-behind text. The command reads the spec in Go, serializes a normalized JSON payload to the configured bridge provider, and uses the VBIDE Designer API rather than editing `.frx` directly. By default, an existing UserForm with the same `form.name` fails with `form_already_exists`; `--overwrite` removes the existing UserForm component and recreates it from the spec, but it must refuse to delete a same-name non-UserForm component. Because Excel requires a save boundary between deleting and recreating a UserForm, overwrite first exports the existing UserForm to a temporary backup, deletes it, saves, and then attempts the rebuild. If the rebuild fails after that checkpoint, xlflow restores the original UserForm from the temporary export and saves the restoration before returning failure. New builds remove the partially created component by reference on failure, including cleanup when Excel rejects the requested component name, so default names such as `UserForm1` are not left behind after an aborted build. The public replacement workflow is `form build --overwrite` rather than in-place Designer mutation. The canonical source-controlled artifact for UserForm Designer structure is `src/forms/specs/*.yaml`; code-behind authority depends on `[userform].code_source`. In `sidecar` mode, xlflow reapplies `src/forms/code/<FormName>.bas` to the rebuilt form when present, synchronizes the tracked `.frm` artifact from that sidecar before build, and overwrite falls back to the pre-delete workbook code-behind if no sidecar exists yet. After a successful `form build` or hidden `form apply`, xlflow immediately exports the resulting UserForm back into tracked source artifacts so `src/forms/<FormName>.frm`, `.frx`, and optional `src/forms/code/<FormName>.bas` stay aligned with the workbook even during session-backed `--no-save` workflows. That exported `.frm` text normalizes the top-level `Caption` property from the live Designer state so the tracked artifact does not regress to default names such as `UserForm1` after a rebuild. In `frm` mode, xlflow preserves the deleted workbook form's code-behind without consulting `src/forms/code`. Exported `.frm` / `.frx` files remain build/pull artifacts, but successful build/apply now materializes them as part of the command contract. `form build` saves by default. `--no-save` is valid only with `--session`, and `--overwrite --no-save` is rejected because Excel requires an intermediate save after deleting the old UserForm and before recreating it. Successful results may include contract warnings for weak Designer-backed fields: form-level width/height are best-effort, and design-time `ComboBox` / `ListBox` `list` / `selectedIndex` should be treated as observed-only for round-trip expectations even though xlflow still attempts to apply them.

`form apply` remains hidden and is not maintained for sidecar-aware UserForm code-behind workflows. Use `form build --overwrite` instead.

`form export-image` exports a runtime-rendered workbook UserForm to PNG for visual verification. It follows the same safety model as `inspect form --runtime`: xlflow creates a temporary workbook copy from the current workbook state, loads the form there, optionally invokes `--initializer <method>` with `ThisWorkbook`, shows the form modeless with a unique caption token, captures the matching window, unloads the form, and deletes the temporary workbook copy. `--out` is required and must resolve to `.png`; any other extension fails with `unsupported_image_format` before Excel opens. Existing output files fail with `output_file_exists` unless `--overwrite` is set. Success JSON includes normal workbook/session metadata plus top-level `target`, `forms`, `output`, and `warnings`. `target` includes `kind=file|live_session`, workbook `path`, `form`, and `capture_state="temporary_copy"`. `forms` includes the form `name`, `basis="runtime"`, and optional `initializer`. `output` includes `path`, `format="png"`, optional `created_parent_dirs`, and image dimensions when available. The command always warns that it executed `UserForm_Initialize`, warns when an explicit initializer was invoked, and marks the capture path experimental because it depends on Windows desktop Excel GUI behavior.

`pull` exports standard modules, class modules, userforms, and workbook document modules into the configured source directories. Userforms may emit both `.frm` and `.frx` artifacts. When `[userform].code_source = "sidecar"`, `pull` also writes UserForm code-behind sidecars to `src/forms/code/<FormName>.bas` when the form module contains VBA lines; when `code_source = "frm"`, it leaves code-behind inside `.frm`. Document modules are exported as source text suitable for linting and re-import. Source-controlled `.bas`, `.cls`, and `.frm` files are UTF-8 without BOM. Excel/VBIDE import and export files are treated as CP932 at the bridge boundary, and `pull` converts exported text to UTF-8 before writing the source tree. Pulled UserForm `.frm` artifacts also normalize the top-level `Caption` line from the live Designer state so tracked form text stays consistent with the workbook's design-time caption even when Excel exports stale default names. When `[vba].folders=true`, xlflow reads Rubberduck-style `@Folder("A.B")` annotations near the top of each component and exports nested files under the configured type root, for example `src/modules/A/B/Main.bas`. Before re-export, `pull` removes existing exported `.bas`, `.cls`, `.frm`, and `.frx` files under the configured source roots so moved components do not leave stale paths behind. `pull --session` exports from the workbook opened by `session start`. When workbook UserForms are detected, `pull` adds top-level `warnings` / `hints` explaining that `.frm` text may not fully reflect layout, `.frx`, or Designer-backed state.

`pull` supports the `.NET` bridge in both explicit `--bridge dotnet` mode and Windows `auto` mode. In `auto`, xlflow prefers the `.NET` bridge and falls back to PowerShell only when `.NET` is unavailable, incompatible, or unsupported for the command. The command routes through the `.NET` Excel bridge executable (`xlflow-excel-bridge.exe`) and produces the same envelope fields (`target`, `session`, `workbook`, `source`, `logs`). The `.NET` bridge implementation handles standard modules, class modules, UserForms (`.frm` / `.frx`), document modules, Rubberduck folder annotations, and sidecar UserForm code-behind sidecars with the same behavior as the PowerShell path.

`push` reads source-controlled `.bas`, `.cls`, and `.frm` files as UTF-8 without BOM, writes CP932 temporary import copies under `.xlflow/tmp/`, and imports those temporary files through VBIDE. `.frx` files are binary userform companions and are copied without text conversion. Source enumeration is recursive under each configured `[src]` root. In `sidecar` mode, `src/forms/code/*.bas` is reserved for UserForm code-behind sidecars: those files participate in source fingerprinting but are not imported as standalone standard modules. Instead, before Excel opens xlflow synchronizes tracked `.frm` embedded code from the sidecar so `src/forms/code/<FormName>.bas` remains authoritative, then after a `.frm` import succeeds it reapplies that sidecar to the imported UserForm `CodeModule`. If a sidecar contains exported UserForm header lines such as `Attribute VB_Name`, preflight fails with `source_preflight_failed` before Excel opens instead of synchronizing corrupted `.frm` content. In `sidecar` mode, push also validates `src/forms/specs/*.{yaml,yml,json}` against tracked `.frm` artifacts before Excel opens: the spec filename, `form.name`, `.frm` basename, and `.frm` `Attribute VB_Name` must agree, otherwise preflight fails with `source_preflight_failed` and issue code `FRM201` so xlflow cannot import the wrong Designer-backed form. In `frm` mode, `.frm` embedded code remains authoritative and `src/forms/code` is ignored. When `[vba].folder_annotation="update"`, xlflow rewrites `@Folder("A.B")` comments in temporary import copies from the file's relative directory below its configured root; the tracked source file itself is not rewritten during `push`. When `folder_annotation="preserve"`, existing annotations are kept as-is; `ignore` disables folder annotation read/write. Before starting Excel, `push` runs fatal source preflight checks for patterns that are known to surface as VBE modal dialogs instead of COM errors, including typographic quote characters, likely C-style quote escapes, statically-known object/member mismatches such as `Worksheet.DisplayGridlines`, removed legacy trace-helper APIs such as `XlflowLog` / `XlflowSetTraceFile`, contaminated UserForm sidecars that still contain `Attribute VB_*` export headers, and spec/artifact mismatches that would cause a different UserForm name to be imported than the one declared in `src/forms/specs`. Recursive source trees are also validated for duplicate VBA component basenames; conflicts fail with `duplicate_module_name` before Excel opens. These failures return `lint_failed`, `analyze_failed`, `source_preflight_failed`, `duplicate_module_name`, or related environment failures, include top-level `issues` and/or `analysis` when applicable, and use validation exit code `1`. By default (or with `--backup=always`), `push` creates a timestamped workbook backup under `.xlflow/backups/<backup-id>/`, writes `metadata.json` beside the copied workbook file, replaces non-document VBA components, updates document modules, runs VBE Compile, saves the workbook, and writes source fingerprints to `.xlflow/state/push.json`. If VBE Compile fails after import, `push` returns `vba_compile_failed` with `error.phase = "compile_vba"` and validation exit code `1`; the workbook is not saved and source fingerprints are not updated. Session-backed compile failures also report `session.dirty = true` / `save_required = true` because the live workbook has already received the imported source.

`push` supports the `.NET` bridge in both explicit `--bridge dotnet` mode and Windows `auto` mode. In `auto`, xlflow prefers the `.NET` bridge and falls back to PowerShell only when `.NET` is unavailable, incompatible, or unsupported for the command. The command routes through the `.NET` Excel bridge executable (`xlflow-excel-bridge.exe`) and produces the same envelope fields (`target`, `session`, `workbook`, `backup`, `source`, `logs`). The `.NET` bridge implementation handles standard modules, class modules, UserForms (`.frm` / `.frx`), document modules, backup mode, changed-only fingerprinting, sidecar UserForm code-behind synchronization, and Rubberduck folder annotations with the same behavior as the PowerShell path.

`push --backup=never` skips the workbook backup. `push --fast` is a development-mode shorthand for `--backup=never --changed-only`. `push --changed-only` compares source fingerprints against `.xlflow/state/push.json`; when unchanged, it skips Excel/VBIDE import and returns `source.changed=false`. When changed or state is missing, v1 safely falls back to the normal full component replacement and refreshes the state file after success. `push --session` forces attachment to the workbook kept open by `xlflow session start`. When `--session` is omitted and `.xlflow/session.json` points at the configured workbook, xlflow auto-reuses that matching live session workbook instead of opening a fresh hidden instance. `push --no-save` is allowed only with `--session` and leaves workbook changes unsaved until `xlflow save` or `xlflow session stop`. Human output must distinguish skipped imports, saved workbook updates, and live-session-only updates that have not yet been written back to disk. When source UserForms are detected, `push` adds top-level `warnings` / `hints` about partial `.frm` fidelity; `push --session --no-save` also adds `userform_unsaved_session_state`.

`backup list` lists the available workbook-file backups for the configured workbook. Each result includes backup `id`, `created_at`, `reason`, original workbook path, and backup file path.

`rollback --latest` restores the newest backup for the configured workbook. `rollback --backup <backup-id>` restores a specific backup. Before replacing the workbook file, xlflow creates a safety workbook backup with reason `pre-rollback`. Rollback restores only the workbook file and never rewrites `src/` automatically. If an xlflow session for the same workbook is active, rollback fails with `workbook_in_use`. Successful results warn that source files may now be out of sync and hint to run `inspect` and `pull`.

`session start` opens the configured workbook in Excel and writes `.xlflow/session.json` with process metadata. Session Excel is kept visible even when `excel.visible=false`, because later CLI invocations must reattach to that specific Excel instance through its window handle. xlflow disables events while opening the workbook, but it must keep workbook macros executable afterward so `run --session` and `test --session` can invoke user VBA against the same open workbook. Commands that only inspect or rewrite workbook structure without running user VBA force-disable automation macros before opening. `session status` reports whether the recorded process is running, whether the configured workbook is open, and whether the live workbook is dirty and needs `xlflow save`. `session stop` saves and closes the workbook, quits Excel, and removes the metadata. `save` saves the matching live session workbook when `.xlflow/session.json` points at the configured workbook; `save --session` forces that same path and fails if no matching session workbook is running. Phase 1 UserForm detection for `session` / `save` is best-effort: detected forms add warnings/hints, but failure to inspect UserForms does not fail the session command. Like `list forms`, `inspect form`, `form snapshot`, `form build`, `form export-image`, `pull`, `macros`, `export-image`, and `test`, matching recorded session workbooks may be auto-reused when `--session` is omitted.

For multi-step Excel COM verification flows, the preferred contract is to keep one live workbook session open and reuse it for `push`, `run`, `test`, inspect, and save operations instead of reopening Excel for each command. The release-gate shape is `session start -> push --fast --session --no-save -> run/test --session -> save --session -> session stop`. Session-free single commands remain valid when the check is intentionally one-shot or when the user is verifying a saved-file-only path, but repeated workbook-backed commands should default to the session-backed path because it is materially faster and avoids repeated Excel startup costs.

`edit` is a minimal workbook mutation surface for AI-agent development loops. In the MVP it is session-only by policy: `edit cell`, `edit range`, `edit rows`, and `edit columns` require `--session` and fail with `session_required` instead of silently mutating a saved workbook file or opening a hidden isolated copy. `edit cell` supports exactly one of `--value`, `--formula`, or `--fill`; `edit range` supports exactly one of `--fill` or `--clear contents|formats|all`; `edit rows` supports `--height`; `edit columns` supports `--width`. `--events keep|on|off` applies only to `edit cell --value` and `edit cell --formula`. When xlflow temporarily changes `Application.EnableEvents`, it restores the prior value whenever possible and reports the before/after state in JSON.

`runner install`, `runner remove`, and `runner status` manage the persistent workbook module `XlflowRunner`. In v1 this module is a stable marker for fast-run workflows; argument-free `run --fast` uses direct execution when eligible and otherwise keeps the normal temporary harness path.

The legacy trace command surface is removed. Workbook-side runtime debugging uses `XlflowDebug.Log`, and machine-readable execution/debug results use `xlflow run --json` or `xlflow test --json`. Source preflight still reports legacy `XlflowLog` and `XlflowSetTraceFile` calls as validation failures so removed helper APIs do not remain in managed source trees.

`run` uses the positional macro argument when provided. Otherwise it uses `project.entry` from `xlflow.toml`. `--input` overrides `excel.path` for one invocation. `--arg` may be repeated and must use explicit prefixes: `string:hello`, `string:`, `int:7`, and `bool:true`. Empty values are valid only for `string:` arguments. Malformed `int:` and `bool:` values are rejected by the CLI before Excel starts and exit with code `2`. `--msgbox`, `--inputbox`, and `--filedialog` may be repeated to provide scripted responses for `XlflowUI.MsgBox`, `XlflowUI.InputBox`, and `XlflowUI` file dialog wrappers. `--msgbox` accepts `abort|cancel|ignore|no|ok|retry|yes` case-insensitively. `--filedialog` uses `kind:dialog-id=value`, where `kind` is one of `get-open`, `file-open`, `save-as`, or `folder`. Repeating the same `kind:dialog-id=value` pair appends additional file dialog values in order, which is how multi-select open dialogs are represented. The reserved value `@cancel` represents an explicit cancel result and must be used alone for that dialog id. Dialog ids use the stable `dialog-id=value` format and are normalized to lowercase ASCII letters/digits separated by `_` before xlflow injects workbook markers. Ids that normalize to an empty value, or multiple ids that collide after normalization such as `confirm save` and `confirm-save`, are rejected before Excel starts. `--ui-stream` is an opt-in observability flag for headless-style runs that stream resolved `XlflowUI` events to stderr while preserving stdout for the final result envelope. The default run never saves. `--save` persists the opened workbook in place after a successful run. `--save-as` writes a copy after a successful run and must keep the same workbook extension as the opened workbook. `--save` and `--save-as` cannot be combined.

`run` supports the `.NET` bridge in both explicit `--bridge dotnet` mode and Windows `auto` mode. In `auto`, `xlflow run --json` prefers the `.NET` bridge and falls back to PowerShell only when `.NET` is unavailable, incompatible, or unsupported for the command. The command routes through the `.NET` Excel bridge executable (`xlflow-excel-bridge.exe`) and produces the same envelope fields (`target`, `session`, `workbook`, `macro`, `logs`). The `.NET` bridge implementation handles macro invocation, argument passing, save/save-as operations, and session-aware workbook attachment with the same behavior as the PowerShell path. When `--session` is used and the workbook is not explicitly saved, the `.NET` bridge reports `dirty=true`, `save_required=true`, and `source_of_truth=live_workbook` matching the PowerShell contract. `--msgbox`, `--inputbox`, and `--filedialog` response injection, `--ui-stream` module injection, and `__XLFLOW_DEBUG_PIPE__` injection also work through the `.NET` bridge path. Runtime injection is transactional: if defined-name injection cannot be completed, xlflow rolls back partial injection and does not invoke the target macro.

`run --direct` executes an argument-free macro through `Excel.Run($MacroName)` without injecting the temporary harness module. It cannot be combined with `--arg`, `--ui-stream`, or `--diagnostic`; those combinations fail before Excel starts. Direct runs return weaker VBA diagnostics than the harness path, but default `run` still suppresses Excel-owned VBA runtime error dialogs and routes them back to structured CLI output unless `--gui-compile-errors` is set. `run --fast` uses direct execution when the macro has no CLI arguments and diagnostic mode is not requested; otherwise it falls back to the normal harness. `run --session` forces attachment to the workbook opened by `session start`. When `--session` is omitted and `.xlflow/session.json` points at the configured workbook, xlflow auto-reuses that matching live session workbook rather than an arbitrary active Excel instance.

`run` performs the same fatal source preflight checks as `push` before Excel starts whenever it targets the configured project workbook, including `--session` runs. If that preflight finds known VBE-dialog-causing source problems, the run fails with validation exit code `1` and top-level `issues` and/or `analysis` instead of launching Excel. `run --headless` is for AI agents, tests, and CI. Before Excel starts, xlflow also scans the configured VBA source tree for GUI boundaries. If any boundary is found, the run fails with `gui_boundary_detected`, exit code `1`, and top-level `gui_boundaries` containing the detected file, line, kind, symbol, severity, message, and suggestion. `run --interactive` is for human-assisted Excel workflows. It runs with Excel visible and alerts enabled so a person can complete dialogs, message boxes, or forms. `--headless` and `--interactive` cannot be combined. `--timeout` defaults to `5m`; if a run exceeds the timeout, xlflow returns `macro_timeout` with exit code `1` and guidance that a dialog, form, file picker, or loop may still be waiting. Running without either mode keeps the legacy behavior except for the timeout.

`run --diagnostic` is an opt-in compile-first mode for agent debugging. It runs the same preflight checks as normal `run`, then uses the temporary harness path even when `--fast` is set. Before verifying and invoking the macro, xlflow executes VBE Compile for the workbook VBA project. If VBE shows a modal compile dialog, xlflow reads the dialog text as localized opaque text, closes the dialog, reads `ActiveCodePane.GetSelection` when available, and returns `vba_compile_failed` with validation exit code `1`. JSON output includes `error.phase = "compile_vba"` and top-level `run_diagnostic.kind = "compile"` with `message`, `location`, `nearby_code`, and `dialog` fields when available. During normal macro invocation, xlflow also suppresses Excel-owned VBA runtime error dialogs by default and returns `run_diagnostic.kind = "runtime"` with dialog metadata when available. `--gui-compile-errors` is the explicit opt-out for both compile and runtime dialog suppression. Diagnostic mode does not automate arbitrary user `MsgBox`, file picker, or UserForm UI; those remain governed by `--headless` and `--interactive`.

Before user VBA starts, `run` injects workbook-scoped reserved names that describe the resolved xlflow execution mode, currently including `__XLFLOW_MODE__` and a runtime helper version marker. New `xlflow new` projects also scaffold `src/modules/XlflowRuntime.bas`, which reads workbook-scoped mode first and falls back to `Environ$("XLFLOW_MODE")` only as a secondary override. Stable helper calls are module-qualified VBA functions such as `XlflowRuntime.ModeName()`, `XlflowRuntime.IsHeadless()`, `XlflowRuntime.IsAgent()`, and `XlflowRuntime.IsTest()`. `run --headless` resolves to `headless`; `run --interactive` resolves to `interactive`; plain `run` defaults to `interactive` unless the xlflow process environment sets `XLFLOW_MODE=interactive|headless|ci|agent|test`. During `run` and `test`, xlflow may also inject additional reserved names for helper modules, including `__XLFLOW_DEBUG_PIPE__` for `XlflowDebug.Log` transport and the existing `XlflowUI` runtime stream markers.

New scaffolded projects also add `src/modules/XlflowUI.bas` and `src/modules/XlflowDebug.bas`. Existing projects can add the same bundled helper modules later with `init --with-module` during bootstrap or `module install` after bootstrap. `XlflowUI.MsgBox`, `XlflowUI.InputBox`, `XlflowUI.GetOpenFilename`, `XlflowUI.FileDialogOpen`, `XlflowUI.GetSaveAsFilename`, and `XlflowUI.FolderPicker` require a stable dialog id that contains at least one ASCII letter or digit. In interactive mode they delegate to the native VBA or Excel dialog APIs. In headless-like modes (`headless`, `ci`, `agent`, `test`) they resolve scripted responses from xlflow-injected workbook markers keyed by the normalized dialog id. `GetOpenFilename` and `FileDialogOpen` return a Variant string array when `MultiSelect=True`, a single string when exactly one value is resolved, and `False` on cancel. `GetSaveAsFilename` and `FolderPicker` return a single string or `False` on cancel. Missing or invalid scripted responses are deterministic VBA errors owned by `XlflowUI` rather than interactive fallback. `XlflowDebug.Log` is an explicit workbook-side debug wrapper for terminal-visible logging. It always writes to the VBA Immediate Window and, during xlflow `run` or `test`, also mirrors the rendered message to xlflow's stderr/debug envelope.

`run` adds a `macro` object with `name`, `args`, and `duration_ms`. Failed macro runs return `macro_failed` with `error.source`, `error.number`, `error.message`, `error.line` when VBA exposes a non-zero `Erl` value, and `error.phase` when the failed phase is known. Stable run phases are `open_workbook`, `prepare_vbide`, `compile_vba`, `verify_macro`, `inject_harness`, `invoke_macro`, and `save_result`. When Excel exposes enough information to distinguish a missing or invalid target macro from user-code failure, `run` returns `macro_not_found` instead of `macro_failed`; when Excel blocks invocation because macros are disabled by workbook security state, `run` returns `macro_disabled`. Suppressed runtime error dialogs keep `error.phase = "invoke_macro"` and may populate `error.message` from the localized VBA dialog text plus `run_diagnostic.dialog` / `run_diagnostic.location` when Excel exposes that context. Plain-text success output must include the elapsed duration and whether the workbook was saved, copied, left unchanged on disk, or may now differ from disk because a live session workbook was used without an explicit save. When `ui.events` are present, human output also includes a `UI` section summarizing captured dialog activity, and when `debug.events` are present it includes a `Debug` section summarizing `XlflowDebug.Log` lines. Plain-text failure output must use the formatted message `Module line <n> Err <n>: <description>` when line and error number are available, and otherwise omit the `line <n>` segment. Because `run` injects a temporary VBA harness to measure duration while avoiding modal VBA runtime error dialogs, VBIDE access failures return an environment error such as `vbide_access_denied` and exit code `3`.

`--arg` public types are `string`, `int`, `double`, and `bool`. `double` values
must be finite invariant-culture numbers; empty values, locale decimal commas,
`NaN`, and infinities are rejected before Excel starts. PowerShell and .NET
bridges generate equivalent VBA numeric literals for `double` arguments.

`run --no-save` is an explicit alias for the default "leave disk unchanged"
behavior and cannot be combined with `--save` or `--save-as`. On timeout,
xlflow still returns a valid JSON envelope and exit code `1`, but the result
must be treated as `vba_may_still_be_running`; xlflow does not perform
synchronous COM cleanup on that path.

`run` no longer provides a trace helper path. Runtime debug detail comes from `XlflowDebug.Log`, which writes Immediate Window output and, during xlflow execution, also emits structured `debug.events` in the final JSON envelope plus realtime stderr debug lines. If a failed run has insufficient context, callers should add `XlflowDebug.Log` near procedure entry, important branches, external file access, destructive operations, and error handlers, then rerun with `xlflow run --json`.

`test` resolves the runtime execution mode to `test` before user VBA starts and injects the same workbook-scoped runtime markers used by `run`. New scaffolded projects can branch inside tests with `XlflowRuntime.IsTest()` or inspect the string value with `XlflowRuntime.ModeName()`. `test` also accepts the same repeated `--msgbox`, `--inputbox`, and `--filedialog` flags as `run`, plus `--ui-stream` for realtime stderr streaming of resolved `XlflowUI` events while tests are running. As with `run`, the workbook-scoped marker is the primary contract and `Environ$("XLFLOW_MODE")` remains only a secondary fallback for manual helper adoption in older projects.

`test` supports the `.NET` bridge in both explicit `--bridge dotnet` mode and Windows `auto` mode. In `auto`, xlflow prefers the `.NET` bridge and falls back to PowerShell only when `.NET` is unavailable, incompatible, or unsupported for the command. The command routes through the `.NET` Excel bridge executable (`xlflow-excel-bridge.exe`) and produces the same envelope fields (`target`, `session`, `workbook`, `tests`, `logs`). The `.NET` bridge implementation handles test discovery (`Test*`/`*_Test`), `@Tag` annotation collection, lifecycle hooks (`BeforeAll`/`AfterAll`/`BeforeEach`/`AfterEach`), test filtering (by name/module/tag), test runner VBA code generation with per-module hook wrappers, inconclusive detection (`vbObjectError + 516`), runtime injection, UI stream injection, and session-aware workbook attachment with the same behavior as the PowerShell path. Runtime injection is transactional: if defined-name injection cannot be completed, xlflow rolls back partial injection and does not execute tests.

`export-image` opens the target workbook through the same session-aware Excel bridge as other workbook-backed commands and exports one worksheet range to PNG using Excel's visual copy path. `--sheet` and `--range` are required. When `[workbook]` is omitted, the command targets `excel.path`; when `--session` is omitted and `.xlflow/session.json` points at that same workbook, xlflow auto-reuses the live session workbook. Without `--out` or `--output-dir`, the default output path is `.xlflow/artifacts/images/<workbook-name>/<sheet>_<range>_<timestamp>.png`. `--out` is the full output path, `--output-dir` chooses only the directory, and `--name` chooses only the filename. Parent directories are created automatically. Existing files fail with `output_file_exists` unless `--overwrite` is set. Only PNG is supported in v1; unsupported format requests fail with `unsupported_image_format`. Success JSON includes top-level `target`, `output`, and optional `warnings`, plus normal workbook session metadata. `target` includes `kind=file|live_session`, workbook path, sheet, and range. `output` includes `path`, `format`, `default`, optional `created_parent_dirs`, and image dimensions when available. Cleanup failures after a successful export are reported as warnings rather than turning the export into a failure.

`fmt` is a source-only, conservative, non-destructive VBA formatter. It targets `.bas` and `.cls` files under configured project source directories and `tests/`; `.frm` files are skipped. At most one mode flag (`--write`, `--check`, or `--diff`) is allowed per invocation. When no mode flag is set, `fmt` runs in inspect mode and reports whether files would be changed without modifying them. `--stdin` reads VBA source from stdin (`.bas` assumed) and writes formatted output to stdout; it cannot be combined with other mode flags or `--line-numbers`. When `--stdin` is combined with `--json`, the JSON envelope is written to stdout instead of formatted text; the envelope contains `output.changed` / `output.unchanged` summary fields but does not include the formatted source body. `--check` returns exit code `1` when unformatted files are detected. `--diff` writes unified diffs to the logs and returns exit code `0` even when changes exist. `--write` persists formatted output back to source files. `--line-numbers` follows the same contract: without `--write`, `fmt --line-numbers ...` only reports the prospective numbering changes; with `--write`, it applies them to disk. `--line-numbers` controls explicit VBA line-number handling with `preserve`, `add`, `remove`, and `renumber`; the default policy is `preserve`, so plain `fmt` does not inject line numbers automatically. `add` numbers supported executable statements with a stable `10`/`10` sequence, but it must not number `Select Case`, `Case` / `Case Else`, or `End Select` control lines, and it must only number the first physical line of an explicit line-continuation statement. `remove` strips supported numeric line prefixes, and `renumber` normalizes supported numbered statements to the same stable sequence. When numeric `GoTo`/`Gosub`/`Resume` label targets make removal or renumbering ambiguous, the formatter must leave the source unchanged for that transformation and report a warning instead of guessing. Success JSON includes top-level `target` with `kind="source"`, `path` reflecting the resolved search scope, and `description="source files"`, plus `output` with `mode`, `changed`, `unchanged`, `skipped`, `total`, optional `changed_paths`, `skipped_paths`, `skipped_reasons`, and nested `line_numbers`. `output.line_numbers` always contains `mode` and `applied`; in `write` mode it also contains optional `files_changed`, `lines_added`, `lines_removed`, and `lines_renumbered`, while inspect/check/diff modes use the prospective fields `files_to_change`, `lines_to_add`, `lines_to_remove`, and `lines_to_renumber`. `output.line_numbers.warnings` remains optional in every mode. Skipped files produce `fmt_skipped_unsupported_extension` warnings. Formatting is deterministic and idempotent: running `fmt` twice produces the same result. The formatter uses 4-space indentation, strips trailing whitespace, normalizes blank lines, preserves class module metadata (`VERSION`, `BEGIN`/`END`, `MultiUse`, and `Attribute VB_*` lines) verbatim, and preserves existing line numbers where possible under the default policy. `.frm` files specified by explicit path are skipped rather than formatted.

`analyze` scans configured source directories without Excel COM for runtime-risk patterns. It returns top-level `analysis`; findings contain `code`, `severity`, `file`, `module`, `procedure`, `line`, `message`, `reason`, `suggestion`, and `nearby_code`. Findings are validation failures with exit code `1`.

`check` runs `lint`, `analyze`, then `doctor`. It continues after lint/analyze findings so source issues and environment status are returned together. JSON output includes top-level `check`, `issues`, `analysis`, and doctor diagnostics. Lint/analyze findings return exit code `1`; doctor/environment failure returns exit code `3`.

`macros` opens the configured workbook and discovers public runnable VBA entrypoints without executing user code. JSON output includes top-level `macros`, where each entry contains `module`, `name`, `qualified_name`, `kind` when available, and `args` when available. `macros --session` reads from the workbook opened by `session start`. Agents should use this command before guessing a `run` target.

`macros` supports the `.NET` bridge in both explicit `--bridge dotnet` mode and Windows `auto` mode. In `auto`, xlflow prefers the `.NET` bridge and falls back to PowerShell only when `.NET` is unavailable, incompatible, or unsupported for the command. The command routes through the `.NET` Excel bridge executable (`xlflow-excel-bridge.exe`) and produces the same envelope fields (`target`, `session`, `workbook`, `macros`, `logs`). The `.NET` bridge implementation discovers public runnable VBA entrypoints with the same behavior as the PowerShell path. When `--session` is used, the `.NET` bridge reports session-attached workbook dirty state using the same fallback logic as the PowerShell bridge: if the dirty state cannot be determined, `dirty` and `save_required` default to `true` and `source_of_truth` defaults to `live_workbook`.

`ui button add`, `ui button list`, and `ui button remove` open the configured workbook through the same session-aware Excel bridge as other workbook-backed commands. `--session` forces attachment to the workbook kept open by `xlflow session start`. When `--session` is omitted and `.xlflow/session.json` points at the configured workbook, the commands auto-reuse that matching live session workbook instead of opening a fresh hidden instance. `add` and `remove` save the live session workbook explicitly via `Workbook.Save()` after successful mutation so the session workbook stays open; `list` is read-only and does not save. When `Workbook.Save()` fails after a successful add or remove, the command returns `save_failed` as an environment failure.

`ui button add` adds or updates an xlflow-managed Excel form-control button. The target worksheet is selected by `--sheet`; if it does not exist, the command fails with `sheet_not_found` unless `--create-sheet` is set. `--cell` is the top-left placement anchor, `--text` becomes the button caption, and `--macro` is assigned to the button `OnAction`. `--width` and `--height` are in Excel points and default to `160` and `40`. The stable internal button name is `xlflow.button.<id>`, where `<id>` is the normalized `--id` value or, when omitted, a normalized value derived from `--macro`. Re-running `add` with the same id updates the existing button instead of creating duplicates. `--verify-macro` checks the workbook VBIDE project for the macro before saving; missing macros fail with `macro_not_found`, and unavailable VBIDE access is an environment failure.

`ui button list` reports only xlflow-managed form-control buttons whose internal names start with `xlflow.button.`. When `--sheet` is provided, only that worksheet is inspected and a missing worksheet fails with `sheet_not_found`. `list` does not save the workbook.

`ui button remove` deletes an xlflow-managed form-control button by `--id`, optionally restricted to `--sheet`. Missing worksheets fail with `sheet_not_found`; missing buttons fail with `button_not_found`. `remove` saves the workbook only after a successful deletion.

`inspect` keeps the current saved-workbook behavior by default, but `inspect workbook|sheets|range|used-range|cell --session` reads the live workbook currently attached to the managed xlflow session through Excel COM. Default file-backed inspect continues to return workbook path, name, active sheet when available from file metadata, per-sheet summaries, rectangular range snapshots, single-cell snapshots, and the lightweight saved-file used range. `--session` uses the same target-specific payload shapes but sets `inspect.source="excel_com"` and `inspect.target_info.kind="live_session"` so callers can distinguish live-session reads from saved-file reads. Range-based inspect commands still default to `--max-rows 100` and `--max-cols 30`; when output is clipped they return `truncated=true`, `returned_range`, and a warning. `inspect range` and `inspect used-range` accept `--include-style` in both file-backed and live-session modes to preserve the existing `values` matrix and deterministically add style metadata for the returned range: `style_included=true`, `cells`, `rows`, `columns`, and `merged_ranges`. When `--include-style` is passed, those style blocks are always present; empty selections return empty arrays for the style blocks. When `--include-style` is absent, `style_included` and the style blocks are omitted. File-backed inspect still warns when a matching live session has unsaved changes and now adds hints for `inspect ... --session` and `save --session`. Live-session inspect reports progress on stderr just like other Excel COM-backed commands. Without `--json`, `--format text` is the default human output. `--format markdown` emits Markdown tables, and `--format json` emits the inspect payload only; machine consumers should still prefer the stable top-level `--json` envelope.

`inspect-gui` scans configured source directories and reports GUI interaction boundaries without opening Excel. JSON output includes top-level `gui_boundaries`. Human output shows each boundary location, kind, symbol, and suggested refactor.

`attach --active` inspects the current active Excel workbook. It verifies that the active workbook path matches configured `excel.path` and reports top-level `workbook.path`, `workbook.configured_path`, `workbook.active`, and `workbook.matches_config`. In this version, `attach` does not change the connection target for `pull`, `push`, or `run`; it only validates the human-opened workbook.

`test` opens the configured workbook, discovers argument-free `Sub` procedures from the workbook VBIDE state, and runs procedures whose names start with `Test` or end with `_Test`. `--filter` uses exact procedure-name matching. `--module` filters by exact module name. `--tag` filters by tag attached via `' @Tag("name")` comment lines directly above the test sub. Because `test` executes user VBA, xlflow must keep workbook macros executable for both fresh opens and `test --session`. `test --session` runs against the workbook opened by `session start` via the recorded session metadata. Duplicate discovered test names, no discovered tests, missing filter targets, and VBA test failures are validation failures. Excel, COM, VBIDE, PowerShell, and script failures are environment failures.

For session-aware workbooks, `test --session` is the preferred validation path whenever the workbook is already open or when it will be followed by additional workbook-backed commands. Avoid reopening the workbook between `run`, `test`, `save`, and inspect commands unless the specific behavior under test is the reopen boundary itself.

`diff` compares two workbook files and optionally two exported VBA source trees. Workbook inputs must use `.xlsx`, `.xlsm`, `.xltx`, or `.xltm`. Workbook state comparison covers sheet additions/removals plus used-range cell values and formulas. VBA comparison is enabled only when both `--vba-before` and `--vba-after` are provided, recursively compares `.bas`, `.cls`, and `.frm` files, ignores other files such as `.frx`, and normalizes CRLF/LF line endings before comparison. Differences are successful command results with exit code `0`; malformed arguments fail with exit code `2`, and unreadable workbooks or source trees fail with exit code `3`.

## Configuration

The MVP only auto-discovers `xlflow.toml` from the current working directory. `vba.toml` is intentionally not supported.

```toml
[project]
name = "sample"
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"
visible = false
display_alerts = false
bridge = "auto"

[src]
modules = "src/modules"
classes = "src/classes"
forms = "src/forms"
workbook = "src/workbook"

[vba]
folders = true
folder_annotation = "update"
default_component_folders = true

[userform]
code_source = "sidecar"

[lint]
require_option_explicit = true
forbid_select = true
forbid_activate = true
forbid_on_error_resume_next = true
detect_implicit_variant = true
forbid_public_module_fields = true
forbid_interactive_input = true
```

## JSON Envelope

All JSON output uses a stable top-level envelope.

```json
{
  "status": "ok",
  "command": "lint",
  "error": null,
  "logs": []
}
```

`status` is either `ok` or `failed`. `error` is `null` on success and a structured object on failure. Error objects contain `code`, `message`, `source`, `number`, `line`, `phase`, `h_result`, and `details`. `h_result` is a hex string (e.g. `"0x80040154"`) populated for COM-origin failures. `details` is an object with additional context such as `source` and `stack_trace`. All error fields except `message` are optional and omitted when empty or zero.

Command-specific fields are added at the top level:

- `diagnostics` for `doctor` (includes `excel.com_activation` which indicates successful `Excel.Application` creation on the STA thread)
- `workbook` and `backup` for Excel file commands
- `source` for commands that write project source files
- `macro` for `run`
- `macros` for `macros`
- `tests` for `test`
- `diff` for `diff`
- `inspect` for `inspect`
- `issues` for `lint`
- `analysis` for `analyze` and `check`
- `check` for `check`
- `run_diagnostic` for enriched `run` failures
- `debug` for `run` and `test` debug-log events emitted by `XlflowDebug.Log`
- `target`, `session`, optional `warnings`, and optional `hints` for workbook-state-aware commands
- `output` for `export-image`
- `output` for `form export-image`
- `forms` for `form build`
- `edit` for `edit`
- `process` for `process list` (array of process objects) and `process cleanup` (cleanup result object)
- `project`, `session`, `state`, `warnings`, and `hints` for `status`
- `session` for session status metadata
- `runner` for persistent runner module status
- `version` for `version`
- `gui_boundaries` for `inspect-gui`, `run --headless` preflight failures, and `doctor` source summaries
- `ui` for `run --ui-stream` / `test --ui-stream` dialog events and `ui button` commands
- `output` for `fmt` result summaries, `export-image`, `form export-image`, and `form snapshot` paths

`fmt` result objects contain `mode` (`"inspect"`, `"write"`, `"check"`, or `"diff"`), `changed`, `unchanged`, `skipped`, `total`, optional `changed_paths`, optional `skipped_paths`, and optional `skipped_reasons`. Each skipped reason contains `path` and `reason` (e.g. `"unsupported extension: .frm"`).

`test` result objects contain `name`, `module`, `status`, `duration_ms`, `tags`, and an optional `error`. `tags` is always present; in Phase 2 it is an empty array (`[]`) for forward compatibility, and Phase 3 populates it when tag parsing is implemented.

`status` values are `passed`, `failed`, and `inconclusive`. `inconclusive` is produced when a test calls `XlflowAssert.AssertInconclusive`.

`error.code` values for test-level failures include:

- `test_failed` — the test body raised an error or assertion failure.
- `test_inconclusive` — the test called `AssertInconclusive`.
- `before_all_failed` — the module's `BeforeAll` hook failed, causing all tests in that module to fail.
- `after_all_failed` — the module's `AfterAll` hook failed, causing all tests in that module to fail.
- `before_each_failed` — the test's `BeforeEach` hook failed before the test body ran.
- `after_each_failed` — the test's `AfterEach` hook failed during cleanup.

`run` and `test` may return `ui.events`, where each event contains `event`, `kind`, `dialog_id`, `prompt`, `title`, `response_source`, optional `resolved_result`, optional `resolved_value`, `redacted`, `runtime_mode`, and optional `error`. When `--ui-stream` is enabled, these same events are also summarized to stderr in real time. `run` and `test` may also return `debug.events`, where each event contains `event`, `message`, `runtime_mode`, `source`, and optional `error`; `debug.count` reports the total number of captured `XlflowDebug.Log` events and `debug.truncated=true` indicates that xlflow kept only the most recent events in the final envelope. `ui button add` and `ui button remove` return `ui.button` with `id`, `name`, `sheet`, `text`, `macro`, `cell`, `left`, `top`, `width`, `height`, and `updated`. `ui button list` returns `ui.buttons` with the same fields for each managed button.

`diff` result objects contain `summary`, `sheets`, `cells`, and `vba`. Cell diffs contain `sheet`, `address`, `kind`, `before`, and `after`, where `kind` is `value` or `formula`. VBA diffs contain `file`, `kind`, and optional changed line details.

`inspect` returns a top-level `inspect` object with `target`, optional `target_info`, `format`, `source`, and one of `workbook`, `sheets`, `range`, `cell`, `form`, or `forms`. Workbook and sheet summaries contain `name`, `index`, `visible`, `used_range`, `row_count`, and `column_count`. Range snapshots contain `sheet`, `range` and/or `used_range`, `returned_range`, `row_count`, `column_count`, `values`, `truncated`, `max_rows`, `max_cols`, and optional `warnings`. When `--include-style` is set on range-based commands, range snapshots deterministically contain `style_included=true` plus `cells`, `rows`, `columns`, and `merged_ranges`; those arrays are empty when the returned range is empty or no merges are present. When `--include-style` is absent, those fields are omitted. Cell snapshots contain `sheet`, `address`, and `value`. Form snapshots contain `name`, `basis`, optional `caption`, optional `width` / `height`, `coordinate_system`, and `controls`; controls include common UI state such as `caption`, `text`, `value`, geometry, visibility, list items, and recursive child controls for supported container controls. For designer inspection, only true top-level controls should appear at the root `controls` array; children belong only under their container's recursive `controls`. `inspect --format json` keeps this standalone inspect payload shape and adds workbook-state metadata alongside it using top-level `target_state`, `session`, `warnings`, and `hints` fields when present. When exported `.frm` files are present under configured `src.forms`, file-based `inspect workbook|sheets|range|used-range|cell` adds a warning that it inspected the saved workbook file and cannot verify live UserForm Designer/runtime state.

`export-image` success results add top-level `target` and `output`, plus optional top-level `warnings`. `target` contains `kind`, `path`, `sheet`, and `range`. `output` contains `path`, `format`, `default`, optional `created_parent_dirs`, and optional `width_px` / `height_px`. Warning objects contain `code` and `message`.

`form export-image` success results add top-level `target`, `forms`, and `output`, plus optional top-level `warnings`. `target` contains `kind`, `path`, `form`, and `capture_state`. `forms` contains `name`, `basis`, and optional `initializer`. `output` contains `path`, `format`, optional `created_parent_dirs`, and optional `width_px` / `height_px`.

`form build` success results add top-level `target`, `session`, optional `warnings` / `hints`, and `forms`. `forms` contains `name`, `basis`, `action`, `control_count`, `spec_path`, optional `caption`, optional `coordinate_system`, and `overwrite`.

`edit` success results add top-level `edit` alongside the normal workbook-state metadata. `edit` includes `kind`, `sheet`, one selector field (`cell`, `range`, `rows`, or `columns`), `mutation`, and optional `events`. `mutation` reports the edited property summary plus affected row/column/cell counts when relevant. `events` is returned for event-aware cell edits and may include `mode`, `enable_events_before`, `enable_events_after`, and `restored`.

`backup list` success results add top-level `backups`. `backups` is an array of backup records. Each record contains `id` (string), `created_at` (ISO 8601 timestamp), `reason` (string such as `before-push` or `pre-rollback`), `workbook` (project-relative or absolute workbook path), and `path` (project-relative or absolute backup workbook path).

`rollback` success results add top-level `rollback`, plus the normal `target`, `warnings`, and `hints`. `rollback.restored_from` contains `id`, `path`, `reason`, and `created_at` for the restored backup. `rollback.safety_backup` contains `id` and `path` for the safety backup created before restore. `rollback.target.path` identifies the restored workbook path.

Workbook-state-aware commands may return top-level `target`, `session`, `warnings`, and `hints`. `target.kind` uses the fixed vocabulary `source`, `file`, or `live_session`. `target` may also include `path`, `description`, `note`, and command-specific fields such as `sheet`, `range`, `form`, and `capture_state`. `session` may include `active`, `workbook_path`, `workbook_name`, `dirty`, `save_required`, `live_newer_than_disk`, `source_of_truth`, `userforms_present`, `userform_count`, and `mode`; `session status` also keeps `running`, `workbook_open`, and `metadata`. Warning and hint objects contain `code` and `message`. `push`, `pull`, `run`, `save`, `session`, `macros`, `inspect`, `export-image`, `form build`, `form export-image`, and `edit` should use these fields to make workbook state explicit without removing the existing compatibility fields under `workbook`.

### `status`

`xlflow status` is a read-only project-level command that aggregates project, source, workbook, and session state. It does not modify workbook files, source files, or `.xlflow/state`.

`status` success results add top-level `project`, `session`, `state`, `warnings`, and `hints`. The `command` envelope field is `"status"`.

`project` contains:

- `root` — resolved project root directory.
- `workbook_path` — resolved configured workbook path.
- `src_paths` — array of resolved source directory paths (modules, classes, forms, workbook).
- `project_name` — configured project name from `xlflow.toml`.

`session` reuses the `session status` payload. When no session is active, `session` contains `active: false` and the command still succeeds. When a live session matches the configured workbook but the workbook dirty state cannot be inspected, `session.dirty` may be `null` rather than `false`; `session.save_required` remains the authoritative signal for unsaved changes.

`state` contains:

- `src_newer_than_workbook` — true when any workbook-affecting source file (`.bas`, `.cls`, `.frm`, `.frx`, plus sidecar code) has an mtime newer than the saved workbook mtime.
- `live_session_newer_than_disk` — reported by the session probe when a live session workbook differs from the saved file on disk.
- `workbook_saved` — true when no live session workbook is open with unsaved changes.
- `source_of_truth` — `"saved_workbook"`, `"live_workbook"`, or `"unknown"`.
- `workbook_last_modified_at` — ISO 8601 mtime of the saved workbook file (omitted when the file is missing).
- `latest_source_modified_at` — ISO 8601 mtime of the most recent workbook-affecting source file (omitted when no source files exist).
- `push_state_last_modified_at` — ISO 8601 mtime of `.xlflow/state/push.json` (omitted when the file does not exist).

`warnings` and `hints` are arrays of objects with `code` and `message`. Warning and hint codes include:

- `session_dirty` / `save_session` — live session has unsaved changes; `xlflow save --session` is recommended.
- `source_newer_than_workbook` / `push_source` — source files are newer than the saved workbook; `xlflow push` is recommended.
- `live_session_newer_than_disk` / `save_before_push` — live session workbook is newer than the saved file on disk.

Source freshness (`src_newer_than_workbook`) is a heuristic based on file modification times. Clock skew, manual file copies, and filesystem timestamp granularity can cause false positives or negatives. Consumers should treat this as an indication, not a strict synchronisation guarantee.

### `process list`

`xlflow process list` enumerates all local Excel processes regardless of whether they were started by xlflow. The command is workbook- and configuration-independent.

Success results add top-level `process` as an array of process objects. Each object contains:

- `pid` — integer process ID.
- `has_workbook` — boolean (`true` when at least one open workbook is confirmed, `false` when no workbook is confirmed, `null` when workbook state could not be determined).

When no Excel processes are running, `process` is an empty array and the command still succeeds.

### `process cleanup`

`xlflow process cleanup` terminates Excel processes in one of three modes:

- `<pid>` — graceful shutdown of a single process (falling back to force-stop if the process persists).
- `--auto` — graceful shutdown of only those Excel processes that have no open workbooks.
- `--all` — force-stop of ALL Excel processes regardless of workbook state.

`cleanup --all` always prompts for interactive confirmation unless `--yes` is passed. When `--json` is set with `cleanup --all`, `--yes` is required; otherwise a configuration error is returned. Confirming cancellation returns `process_cancelled` (exit code 0).

`cleanup` success results add top-level `process` as an object with:

- `action` — `"cleanup"`.
- `mode` — `"pid"`, `"auto"`, or `"all"`.
- `total` — number of targeted processes.
- `results` — array of per-process objects, each containing `pid`, `terminated` (boolean), and `method` (`"graceful"`, `"force"`, `"none"`, or `"unknown"` when the process exited but the shutdown method could not be determined).

`process cleanup` may return these error codes:

- `process_args_invalid` (exit code 2): invalid argument combinations.
- `process_not_found` (exit code 2): the requested PID does not correspond to an Excel process.
- `process_cancelled` (exit code 0): the user declined the `--all` confirmation prompt. This error code is set by the CLI layer and is never produced by the PowerShell bridge; its mapping does not belong in `exitCodeForScriptResult`.
- `process_enumeration_failed` (exit code 3): process enumeration failed at the system level.
- `process_termination_failed` (exit code 3): one or more targeted processes could not be terminated.
- `process_cleanup_failed` (exit code 3): an unexpected error occurred during process cleanup.

## Exit Codes

- `0`: success
- `1`: user-code or validation failure, including lint findings, analysis findings, GUI boundary preflight failures, macro failure, macro timeout, VBE compile failure, missing macro target, legacy trace-helper findings, missing UI sheets or buttons, VBA test failure, no tests found, missing filter targets, active workbook mismatches, duplicate test names, invalid exported ranges, existing output files, unsupported export-image formats, unsupported form export-image formats, missing UserForms, `form_already_exists`, `unsupported_form_control`, `designer_write_failed`, capture window lookup failures, image capture failures, `edit` session requirements, invalid workbook edit selectors, invalid edit colors, `fmt_check_failed` (unformatted files in `--check` mode), and workbook event-handler failures returned by the bridge
- `2`: CLI argument or configuration error, including invalid `push`, `run`, `session`, `save`, `runner`, `export-image`, `form build`, `form export-image`, `edit`, and `process` option combinations, invalid `fmt` mode combinations, plus `process_args_invalid` and `process_not_found` errors from the bridge
- `3`: environment failure, including Excel, COM, VBIDE, PowerShell, script execution failures, and process enumeration/termination errors (`process_enumeration_failed`, `process_termination_failed`, `process_cleanup_failed`)

`diff` intentionally returns `0` when differences are found. Consumers should inspect `diff.summary.total_diffs` to distinguish changed and unchanged inputs.

## VBA Test Rules

New and initialized projects include `src/modules/XlflowAssert.bas` with `AssertEquals expected, actual, [message]`. The helper is scalar-only: it compares normal scalar values, treats `Null` as equal only to `Null`, and raises a clear assertion error for object or array inputs. Compare object properties such as `Range.Value2` instead of passing object references.

Example:

```vb
Public Sub TestCreateReport()
    AssertEquals 10, Sheets("Result").Range("A1").Value
End Sub
```

## Lint Rules

- `VB001`: missing `Option Explicit`
- `VB002`: `Select` usage
- `VB003`: `Activate` usage
- `VB004`: `On Error Resume Next` usage
- `VB005`: possible implicit `Variant`
- `VB006`: module-level `Public` variable usage
- `VB007`: automation-hostile GUI boundaries such as file pickers, modal dialogs, UserForms, message pumps, or external process launches. Direct `MsgBox` and `InputBox` usage remains in scope for this rule; `XlflowUI.MsgBox` and `XlflowUI.InputBox` are the approved dialog wrappers for runtime-aware automation. JSON findings may include `kind`, `symbol`, and `suggestion`.
- `VB008`: typographic quote characters that can trigger VBE compile dialogs
- `VB009`: likely C-style quote escapes in VBA string literals
- `VB010`: unterminated `Sub`, `Function`, or `Property` procedure
- `VB011`: unexpected `End Sub`, `End Function`, or `End Property`
- `VB012`: mismatched procedure end statement
- `VB013`: missing whitespace before a line-continuation underscore

Projects that intentionally use interactive GUI entrypoints may set `[lint].forbid_interactive_input = false` to suppress `VB007`. This changes lint behavior only; `run --headless` still rejects GUI boundaries during preflight.

Compile-dialog prevention findings `VB008` through `VB013` are always enabled and block source preflight before `push` or `run` opens Excel.

## Analysis Rules

- `VBA101`: object variable assignment likely missing `Set`
- `VBA102`: object-returning function assignment likely missing `Set`
- `VBA103`: object-returning function body likely missing `Set <FunctionName> = ...`
