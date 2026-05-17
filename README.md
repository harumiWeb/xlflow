<p align="center">
    <img width="600" alt="xlflow logo" src="docs/images/logo.png" />
</p>

<p align="center">
  <em>Excel VBA development, rebuilt for CLI-first humans and AI agents.</em>
</p>

<p align="center">
  <a href="README.md">English</a>
  |
  <a href="README.ja.md">日本語</a>
</p>

<div align="center">

![GitHub Release](https://img.shields.io/github/v/release/harumiWeb/xlflow?include_prereleases) ![Scoop](https://img.shields.io/scoop/v/xlflow?bucket=https%3A%2F%2Fgithub.com%2FharumiWeb%2Fscoop-bucket) ![GitHub License](https://img.shields.io/github/license/harumiWeb/xlflow) ![GitHub go.mod Go version](https://img.shields.io/github/go-mod/go-version/harumiWeb/xlflow) [![Ask DeepWiki](https://deepwiki.com/badge.svg)](https://deepwiki.com/harumiWeb/xlflow)

</div>

# :surfing_man: xlflow

**xlflow** is an Excel VBA development framework for the AI agent era.

It turns `.xlsm` workbooks into a source-controlled, CLI-driven development workflow where VBA can be exported, edited, linted, imported, tested, traced, and executed from the command line.

> [!TIP]
> Think of xlflow as a development harness around Excel VBA: it does not replace Excel, but it makes Excel VBA projects much easier to operate from terminals, scripts, CI-like local checks, and AI coding agents.

## Demo

These [samples](example) were created by an AI agent using xlflow with only minimal natural language instructions.

<table>
  <tr>
    <td align="center" width="50%">
      <img src="docs/images/world-news.png" alt="world news" width="100%">
      <sub>Macro that summarizes world news in Excel using NewsAPI</sub>
    </td>
    <td align="center" width="50%">
      <img src="docs/images/stock-price.png" alt="stock price" width="100%">
      <sub>Macro that retrieves stock prices and displays them in Excel</sub>
    </td>
  </tr>
  <tr>
    <td align="center" width="50%">
      <img src="docs/images/gen-qrcode.png" alt="generate qrcode" width="100%">
      <sub>Macro that generates QR codes using cell colors and displays them in Excel</sub>
    </td>
    <td align="center" width="50%">
      <img src="docs/images/tetris.gif" alt="tetris" width="100%">
      <sub>Macro that allows playing Tetris within Excel</sub>
    </td>
  </tr>
  <tr>
    <td align="center" width="50%">
      <img src="docs/images/space-invader.gif" alt="space invader" width="100%">
      <sub>Macro that allows playing Space Invader on a UserForm</sub>
    </td>
    <td align="center" width="50%">
      <img src="docs/images/calendar-picker.png" alt="calendar picker" width="100%">
      <sub>Rich calendar picker</sub>
    </td>
  </tr>
</table>

---

## Why xlflow?

Traditional VBA development is still heavily tied to the Excel UI and the Visual Basic Editor.
That works for small manual edits, but it becomes painful when you want repeatable development, source control, tests, diffs, or AI-agent-assisted changes.

| Pain in normal VBA development                       | What xlflow adds                                               |
| ---------------------------------------------------- | -------------------------------------------------------------- |
| VBA code is trapped inside `.xlsm` files             | Export/import VBA as `.bas`, `.cls`, and `.frm` source files   |
| Macro entrypoints are unclear                        | Discover runnable `Public Sub` procedures with `xlflow macros` |
| Existing UserForm names are unclear                  | Discover workbook UserForms with `xlflow list forms`           |
| Existing UserForm layout is hard to review in diffs  | Persist Designer state with `xlflow form snapshot`             |
| Runtime failures are hard to locate                  | Return structured errors, diagnostics, and trace logs          |
| File dialogs and `MsgBox` block automation           | Detect GUI boundaries before headless runs                     |
| Workbook changes are hard to review                  | Compare values, formulas, sheets, and exported VBA source      |
| AI agents cannot safely operate Excel through the UI | Provide stable CLI commands and JSON output                    |

```text
pull → edit → push → lint → test/run → trace → diff
```

---

## What xlflow can do

| Area           | Capabilities                                                                                                                                      |
| -------------- | ------------------------------------------------------------------------------------------------------------------------------------------------- |
| Source control | Export and import standard modules, class modules, UserForms, and document modules                                                                |
| Execution      | Run macros from the CLI with typed arguments                                                                                                      |
| Testing        | Discover and run VBA test procedures                                                                                                              |
| Linting        | Catch `Option Explicit` omissions, `Select`/`Activate`, broad error handling, implicit variants, public module fields, and interactive operations |
| GUI safety     | Detect file pickers, input boxes, modal message boxes, and other automation-hostile boundaries                                                    |
| Debugging      | Collect trace events and return runtime diagnostics                                                                                               |
| Diffing        | Compare workbook cell values, formulas, sheet structure, and exported VBA source                                                                  |
| AI agents      | Return stable JSON and install bundled Skills for Codex, Claude, Cursor, Gemini, GitHub Copilot-style agent workflows, and other agents           |

> [!IMPORTANT]
> xlflow is **Windows-first**. Workbook operations use **Microsoft Excel + COM + PowerShell**.

> [!NOTE]
> Excel COM-backed commands report the xlflow bridge host in top-level `bridge`.
> If workbook VBA launches its own external PowerShell process, that host can still differ from xlflow's bridge host. Inspect or log the workbook-side executable when debugging `powershell.exe` vs `pwsh.exe` behavior.

---

## Requirements

| Requirement                                  | Needed for                                                                                                                                                                          |
| -------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Windows                                      | Excel COM automation                                                                                                                                                                |
| Microsoft Excel                              | `new`, `init`, `list forms`, `inspect form`, `form snapshot`, `form build`, `form export-image`, `pull`, `push`, `run`, `export-image`, `edit`, `test`, `macros`, `trace`, `doctor` |
| PowerShell                                   | Excel automation bridge                                                                                                                                                             |
| Trust access to the VBA project object model | Reading and writing VBA projects                                                                                                                                                    |

> [!NOTE]
> Commands that do not require Excel COM, such as `lint`, parts of `diff`, and Go unit tests, can be verified in non-Excel environments.

> [!WARNING]
> In Excel, enable **Trust access to the VBA project object model** before using commands that read or write VBA code. Without it, `pull`, `push`, `run`, and related commands may fail even when Excel itself is installed.
>
> <details>
> <summary>Details</summary>
> In Excel options, please enable "Trust Center" → "Macro Settings" → "Trust access to the VBA project object model".
> </details>

---

## Installation

### Go install

```bash
go install github.com/harumiWeb/xlflow/cmd/xlflow@latest
```

`go install` may contact the Go module mirror and checksum database configured in your Go environment. For direct source checkout development and CI, treat the Go version declared in `go.mod` as the supported toolchain source of truth; the repository CI and release workflows resolve Go from that file.

### Scoop

```powershell
scoop bucket add harumiweb https://github.com/harumiWeb/scoop-bucket
scoop install xlflow
```

### GitHub Releases

Download prebuilt Windows binaries from:

[https://github.com/harumiWeb/xlflow/releases](https://github.com/harumiWeb/xlflow/releases)

> [!IMPORTANT]
> Current prebuilt distribution targets **Windows** only.
> Commands that interact with workbooks still require **Microsoft Excel**, Excel COM automation, and **Trust access to the VBA project object model**.
> The release binary already embeds the runtime PowerShell bridge scripts, so `xlflow.exe` can run workbook commands without sidecar `*.ps1` files.

Verify the downloaded ZIP against the published `checksums.txt` file:

```powershell
Get-FileHash .\xlflow_windows_x86_64.zip -Algorithm SHA256
certutil -hashfile .\xlflow_windows_x86_64.zip SHA256
```

The reported SHA256 must match the entry for `xlflow_windows_x86_64.zip` in `checksums.txt`.

> This confirms file integrity against the published checksum file. It does not prove publisher identity and is not a substitute for Windows Authenticode signing.

Verify the GitHub Actions provenance attestation with GitHub CLI:

```powershell
gh attestation verify .\xlflow_windows_x86_64.zip --repo harumiWeb/xlflow
```

> This confirms the artifact attestation published for the release artifact. It does not mean the ZIP is Authenticode-signed by a Windows publisher certificate.

Verify the installation:

```bash
xlflow version
xlflow --help
```

For development checkout usage:

```bash
go run ./cmd/xlflow --help
```

With Taskfile:

```bash
task run -- --help
```

---

## Quick start

### 1. Create or initialize a project

Create a new xlflow project and macro-enabled workbook:

```bash
xlflow new Book.xlsm
```

Or start from an existing workbook:

```bash
xlflow init Book.xlsm
```

Install the AI agent Skill during project creation:

```bash
xlflow new Book.xlsm --with-skill --agent codex
```

Interactive `xlflow new` and `xlflow init` render a welcome banner and may check the latest GitHub Release through the GitHub Releases API. Disable that request for one invocation with `--no-update-check`, or set `XLFLOW_NO_UPDATE_CHECK=1` to disable it for interactive scaffolding in your environment.

### 2. Check the Excel automation environment

```bash
xlflow doctor --json
```

> [!TIP]
> If `pull`, `push`, `run`, or `test` fails because of Excel, COM, PowerShell, or VBIDE settings, run `doctor` first.

### 3. Export VBA into source files

```bash
xlflow pull --json
```

Edit the exported `.bas`, `.cls`, and `.frm` files under `src/` with your normal editor. When folder mode is enabled, nested directories under each configured source root are mapped to Rubberduck-compatible `@Folder(...)` annotations during `push`.

### 4. Import edited source back into the workbook

```bash
xlflow push --json
```

### 5. Discover and run macros

```bash
xlflow macros --json
xlflow run Main.Run --json
```

For unattended automation, prefer headless mode:

```bash
xlflow run Main.Run --headless --json
```

If the macro intentionally shows file pickers, message boxes, or UserForms, use interactive mode:

```bash
xlflow run Main.Run --interactive --timeout 5m --json
```

### 6. Lint and test

```bash
xlflow lint --json
xlflow test --json
```

---

## Common workflows

### AI-agent-assisted VBA editing

```text
1. Read xlflow.toml
2. Start xlflow session start for normal editing work
3. If the latest source of truth is unclear, run xlflow pull --session --json
4. Edit files under src/
5. Run xlflow push --fast --session --no-save --json
6. Run xlflow lint --json
7. Run xlflow test --session --json
8. Run xlflow macros --session --json
9. Run xlflow run <qualified_name> --headless --session --json
10. Use xlflow run --trace --session --json when runtime failures are unclear
11. Run xlflow save --session --json before xlflow session stop
12. Use xlflow diff --json when workbook changes must be reviewed
```

> [!IMPORTANT]
> AI agents and CI-like scripts should prefer `--json`. The JSON envelope is designed to be stable and easier to parse than human-readable output.
> `xlflow run` now compiles VBA and returns structured compile diagnostics by default. Use `--gui-compile-errors` only when a human explicitly wants raw Excel/VBE dialogs.

### Human-assisted Excel sessions

Use `attach` when a human has Excel open and you want to validate the active workbook before working with it:

```bash
xlflow attach --active --json
```

> [!NOTE]
> `attach` is a safety check. It confirms that the active Excel workbook matches the configured `excel.path`; it does not change the target used by `pull`, `push`, or `run`.

### GUI-heavy macros

Inspect GUI boundaries before deciding whether a macro can run headlessly:

```bash
xlflow inspect-gui --json
```

| Result                                                        | Suggested mode                                             |
| ------------------------------------------------------------- | ---------------------------------------------------------- |
| No GUI boundaries                                             | `xlflow run ... --headless --json`                         |
| File picker, `InputBox`, modal `MsgBox`, or UserForm detected | `xlflow run ... --interactive --timeout 5m --json`         |
| GUI code wraps core logic                                     | Refactor core logic into parameterized headless procedures |

> [!WARNING]
> Headless automation and modal Excel UI do not mix. Use `inspect-gui` before unattended runs and keep GUI entrypoints thin.

---

## Command map

| Command             | Purpose                                                     | Typical usage                                                                |
| ------------------- | ----------------------------------------------------------- | ---------------------------------------------------------------------------- |
| `new`               | Create a new xlflow project and `.xlsm` workbook            | `xlflow new Book.xlsm`                                                       |
| `init`              | Initialize xlflow from an existing workbook                 | `xlflow init Book.xlsm`                                                      |
| `doctor`            | Diagnose Excel, COM, PowerShell, and VBIDE access           | `xlflow doctor --json`                                                       |
| `attach`            | Validate the workbook currently active in Excel             | `xlflow attach --active --json`                                              |
| `pull`              | Export VBA components into `src/`                           | `xlflow pull --json`                                                         |
| `push`              | Import VBA source back into the workbook                    | `xlflow push --json`                                                         |
| `session`           | Keep the configured workbook open for fast loops            | `xlflow session start`                                                       |
| `save`              | Save the workbook held by a session                         | `xlflow save --session --json`                                               |
| `runner`            | Manage the persistent xlflow runner marker module           | `xlflow runner install --json`                                               |
| `macros`            | Discover runnable macro entrypoints                         | `xlflow macros --json`                                                       |
| `list forms`        | Discover workbook UserForms and expected source paths       | `xlflow list forms --json`                                                   |
| `form snapshot`     | Persist strict Designer UserForm state as JSON or YAML spec | `xlflow form snapshot UserForm1 --out src/forms/specs/UserForm1.yaml --json` |
| `form build`        | Create a Designer-backed UserForm from a saved spec         | `xlflow form build src/forms/specs/UserForm1.yaml --json`                    |
| `form export-image` | Export a runtime UserForm to a PNG image                    | `xlflow form export-image UserForm1 --out artifacts/UserForm1.png --json`    |
| `run`               | Execute a macro from the CLI                                | `xlflow run Main.Run --json`                                                 |
| `export-image`      | Export a worksheet range to a PNG image                     | `xlflow export-image --sheet QR --range A1:AE31 --json`                      |
| `edit`              | Mutate a live session workbook for setup and tuning         | `xlflow edit cell --sheet Input --cell B2 --value ABC123 --session --json`   |
| `trace`             | Enable, collect, and clean VBA trace logs                   | `xlflow trace enable --json`                                                 |
| `test`              | Run VBA tests                                               | `xlflow test --json`                                                         |
| `diff`              | Compare workbook content and optional VBA source            | `xlflow diff before.xlsm after.xlsm --json`                                  |
| `inspect`           | Inspect saved workbook snapshots without Excel COM          | `xlflow inspect range --sheet Result --address A1:F20 --json`                |
| `lint`              | Lint VBA source                                             | `xlflow lint --json`                                                         |
| `analyze`           | Analyze runtime-risk patterns without opening Excel         | `xlflow analyze --json`                                                      |
| `check`             | Run `lint`, `analyze`, and `doctor` as a preflight          | `xlflow check --keepalive --json`                                            |
| `inspect-gui`       | Detect GUI interaction boundaries                           | `xlflow inspect-gui --json`                                                  |
| `skill install`     | Install the bundled xlflow Skill for AI agents              | `xlflow skill install --agent codex`                                         |
| `version`           | Show the installed xlflow build metadata                    | `xlflow version`                                                             |

---

## Commands in detail

<details open>
<summary><strong>Project setup: <code>new</code>, <code>init</code>, <code>doctor</code>, <code>attach</code></strong></summary>

### `xlflow new`

Creates a new xlflow project and `.xlsm` workbook.

```bash
xlflow new
xlflow new Sales
xlflow new Sales.xlsm
```

When no argument is provided, xlflow creates `build/Book.xlsm`.
When the name has no extension, `.xlsm` is appended.
`new` creates a macro-enabled workbook, so extensions other than `.xlsm` are rejected.
Pass `--no-update-check` when you want to skip the interactive GitHub Release lookup during scaffolding.

`new` creates the project structure, including `xlflow.toml`, `src/`, `tests/`, `build/`, and `.xlflow/`.
It also creates or updates `.gitignore` to ignore Excel temporary files and xlflow-generated artifacts.

### `xlflow init`

Creates an xlflow project from an existing Excel workbook.

```bash
xlflow init Book.xlsm
```

The given workbook is copied under `build/`, and its project-local path is recorded in `[excel].path` in `xlflow.toml`.
Pass `--no-update-check` when you want to skip the interactive GitHub Release lookup during scaffolding.

### `xlflow doctor`

Diagnoses the Excel automation environment.

```bash
xlflow doctor --json
```

It checks whether Excel is installed, whether the workbook can be opened, and whether VBIDE access is available.
When source files are available, `doctor` also reports GUI boundary candidates that may block headless runs.

### `xlflow attach`

Validates the workbook currently active in Excel.

```bash
xlflow attach --active --json
```

This is useful for human-assisted sessions where Excel is already open.

</details>

<details open>
<summary><strong>VBA source loop: <code>pull</code>, <code>push</code>, <code>macros</code>, <code>run</code></strong></summary>

### `xlflow pull`

Exports VBA components from the configured workbook.

```bash
xlflow pull --json
```

It exports standard modules, class modules, UserForms, and document modules such as Workbook and Worksheet modules into `src/`.
When `[userform].code_source = "sidecar"`, `pull` also writes code-behind sidecars to `src/forms/code/<FormName>.bas` when the form module contains VBA lines. In `frm` mode, `pull` leaves code-behind inside `.frm`.
Use `xlflow pull --session --json` when you want to require the recorded session workbook explicitly. If `.xlflow/session.json` already points at the configured workbook, plain `xlflow pull --json` auto-reuses that matching live workbook.
When workbook UserForms are detected, `pull` adds warnings that `.frm` text alone may not capture `.frx` or Designer-backed state.

### `xlflow push`

Imports VBA source under `src/` back into the Excel workbook.

```bash
xlflow push --json
```

It reads `.bas`, `.cls`, and `.frm` files and imports them through VBIDE.
UserForm `.frx` files are treated as binary companion files. When `[userform].code_source = "sidecar"`, `src/forms/code/*.bas` sidecars are not imported as standalone modules; xlflow first synchronizes tracked `.frm` embedded code from the sidecar, then `push` reapplies that sidecar onto the matching UserForm `CodeModule` after the `.frm` import succeeds. When `code_source = "frm"`, `.frm` embedded code remains authoritative.
By default, `push` creates a backup under `.xlflow/backups` and saves the workbook.
When source UserForms are detected, `push` adds warnings and deeper-form inspection hints. `push --session --no-save` adds an extra warning that live workbook UserForm state may now differ from disk.

For faster development loops:

```bash
xlflow push --fast --json
xlflow push --changed-only --json
xlflow push --backup=never --json
```

`--fast` is shorthand for `--backup=never --changed-only`.
When `--changed-only` sees the same source fingerprint in `.xlflow/state/push.json`, xlflow skips Excel/VBIDE import.

### `xlflow macros`

Discovers runnable `Public Sub` entrypoints.

```bash
xlflow macros --json
```

> [!TIP]
> AI agents and automation scripts should run this command before guessing a macro name. Use the returned `qualified_name` with `xlflow run` to avoid entrypoint mistakes.
> During session-based development, `xlflow macros --json` already auto-reuses a matching recorded session workbook; add `--session` when you want that requirement to be explicit.

### `xlflow list forms`

Lists workbook `UserForm` components without loading them at runtime.

```bash
xlflow list forms --json
xlflow list forms --session --json
```

Each result includes the form name, whether the expected `.frx` companion exists in the source tree, and the expected project-relative `.frm` / `.frx` paths resolved with the same folder-aware rules as `pull`.

Like other workbook-backed read commands, `list forms` auto-reuses a matching recorded session workbook when `.xlflow/session.json` points at the configured workbook. Add `--session` when you want that requirement to be explicit.

### `xlflow inspect form`

Inspects a workbook `UserForm` through Excel COM and returns structured form/control state.

```bash
xlflow inspect form UserForm1 --runtime --json
xlflow inspect form UserForm1 --runtime --initializer InitializeForm --json
xlflow inspect form UserForm1 --designer --json
xlflow inspect form UserForm1 --both --initializer InitializeForm --json
```

`--runtime` is the default. It runs against a temporary workbook copy created from the current workbook state, loads the form there, inspects the loaded controls, and unloads it before returning so the source workbook is not mutated. `--designer` reads VBIDE Designer state from the source workbook without loading the form at runtime. `--both` returns both snapshots.

`--initializer <method>` is optional and is valid only with `--runtime` or `--both`. It calls the named public form method with `ThisWorkbook` before runtime control enumeration. This is useful for forms whose visible state is populated by a custom initializer instead of `UserForm_Initialize`.

Runtime inspection always warns that `UserForm_Initialize` ran. When `--initializer` is used, xlflow also warns that the explicit initializer ran.

### `xlflow form snapshot`

Persists a strict design-time snapshot of a workbook `UserForm` as a reviewable JSON or YAML spec file.

```bash
xlflow form snapshot UserForm1 --out src/forms/specs/UserForm1.json --json
xlflow form snapshot UserForm1 --out src/forms/specs/UserForm1.yaml --session --json
```

`xlflow inspect form --designer` remains a direct VBIDE Designer read from the source workbook and is intended to work without running workbook VBA. `form snapshot` is stricter: it opens a temporary workbook copy and runs an injected VBA helper so the persisted spec can include concrete control types suitable for later rebuild workflows.

`--out` is required. The output extension and serialized format must match exactly: `.json` writes JSON, and `.yaml` / `.yml` write YAML. Any other extension fails before Excel opens. Because snapshot uses the helper path, it can fail when the workbook's VBA project cannot execute the injected helper.

Persisted `warnings` are reserved for form-local snapshot warnings that belong to the saved spec itself. Operational warnings such as `save_required` remain in the command envelope and human output instead of being written into the artifact.

For ongoing UserForm work, treat `src/forms/specs/*.yaml` as the canonical source-controlled artifact for Designer structure. For code-behind, the authority depends on `[userform].code_source`: new projects default to `sidecar`, where `src/forms/code/*.bas` is canonical, while init/imported projects default to `frm`, where embedded `.frm` code remains canonical until you migrate intentionally. `form snapshot` is the capture path for Designer spec, `pull` is the capture path for code-behind sidecars in sidecar mode, `form build --overwrite` is the rebuild path back into the workbook, and exported `.frm` / `.frx` files should be treated as build or pull artifacts rather than the primary source of truth for Designer behavior.

Like other workbook-backed read commands, `form snapshot` auto-reuses a matching recorded session workbook when `.xlflow/session.json` points at the configured workbook. Add `--session` when you want that requirement to be explicit.

### `xlflow form build`

Creates a Designer-backed workbook `UserForm` from a saved `xlflow.userform` spec.

```bash
xlflow form build src/forms/specs/UserForm1.yaml --json
xlflow form build src/forms/specs/UserForm1.yaml --session --overwrite --json
```

The spec path must end with `.json`, `.yaml`, or `.yml`. xlflow validates the schema in Go before Excel opens, then uses the VBIDE Designer API to create the form and its controls rather than editing `.frx` directly.

When the spec cannot be parsed or validated, `form build` returns structured `spec_parse_failed`, `spec_validation_failed`, or `spec_schema_invalid` errors plus top-level `spec` metadata such as `path`, `format`, optional `line`, optional `column`, optional `field`, and a remediation suggestion. YAML mistakes such as unquoted `-` or `:` values should be corrected by quoting the scalar or switching the artifact to JSON.

By default, a form with the same `form.name` fails with `form_already_exists`. `--overwrite` removes that existing component and recreates it from the spec. This is the recommended replacement workflow when the form design should be rebuilt from source-of-truth spec data under `src/forms/specs/*.yaml`. In `sidecar` mode, xlflow synchronizes the tracked `.frm` artifact from `src/forms/code/<FormName>.bas`, reapplies that sidecar to the rebuilt form when present, and falls back to the deleted workbook form's code-behind if no sidecar exists yet. In `frm` mode, rebuild preserves the deleted workbook form's code-behind without consulting `src/forms/code`. The command saves by default; `--session --no-save` leaves the live workbook dirty and returns save-required state. In that mode, the live workbook is newer than disk until `xlflow save --session` persists it.

Successful `form build` results may still return contract warnings for weak Designer-backed fields. Form-level `width` / `height` are best-effort only, and design-time `ComboBox` / `ListBox` `list` / `selectedIndex` should be treated as observed-only for round-trip expectations even though xlflow still attempts to apply them.

Recommended UserForm loop:

```text
1. xlflow list forms --session --json
2. xlflow inspect form <FormName> --designer --session --json
3. xlflow pull --session --json
4. xlflow form snapshot <FormName> --out src/forms/specs/<FormName>.yaml --session --json
5. in sidecar mode, review or edit src/forms/code/<FormName>.bas if code-behind changed
6. edit src/forms/specs/<FormName>.yaml
7. xlflow form build src/forms/specs/<FormName>.yaml --session --overwrite --json
8. xlflow inspect form <FormName> --designer --session --json
9. xlflow form export-image <FormName> --out artifacts/<FormName>.png --session --json
```

### `xlflow form export-image`

Exports a runtime-rendered workbook `UserForm` to a PNG image.

```bash
xlflow form export-image UserForm1 --out artifacts/UserForm1.png --json
xlflow form export-image UserForm1 --out artifacts/UserForm1.png --initializer InitializeForm --session --overwrite --json
```

This command follows the same safe runtime model as `inspect form --runtime`: xlflow creates a temporary workbook copy from the current workbook state, loads the form there, optionally invokes the named initializer with `ThisWorkbook`, shows the form modeless, captures its window, and then unloads it. The source workbook and recorded live session are not mutated by the capture itself.

`--out` is required and must resolve to `.png`. Existing files fail unless `--overwrite` is set. Successful JSON includes top-level `target`, `forms`, `output`, and `warnings`. The command always warns that it executes `UserForm_Initialize`, and it is marked experimental because it depends on Windows desktop Excel GUI capture behavior.

### `xlflow run`

Runs a macro from the CLI.

```bash
xlflow run Main.Run --json
```

Macros with arguments are supported:

```bash
xlflow run Report.Generate \
  --arg string:fixtures\sample.xlsx \
  --arg int:3 \
  --arg bool:true \
  --json
```

`--arg` accepts typed arguments with `string:`, `int:`, and `bool:` prefixes.
Empty values are allowed only for `string:`.

By default, `run` does not save the workbook. To persist results, explicitly pass `--save` or `--save-as`.

```bash
xlflow run Report.Generate --save --json
xlflow run Report.Generate --save-as build\Result.xlsm --json
```

When execution fails, xlflow returns `macro_failed` or `macro_not_found` with VBA error number, description, module name, phase, and line number when available.
Runtime failures also include `run_diagnostic` when xlflow can match the failure to nearby source or a known VBA pattern such as a missing `Set` assignment.
By default, `run` compiles the VBA project first and converts VBE compile dialogs into structured `vba_compile_failed` JSON with module, line, column, message, and nearby code when available. Use `--gui-compile-errors` only when you intentionally want Excel/VBE compile dialogs to appear instead.

| Mode                   | Behavior                                                                                                                                    |
| ---------------------- | ------------------------------------------------------------------------------------------------------------------------------------------- |
| `--headless`           | Rejects GUI boundaries before Excel starts and returns `gui_boundary_detected` with top-level `gui_boundaries`                              |
| `--interactive`        | Runs with Excel visible and alerts enabled for human operation                                                                              |
| `--direct`             | Runs an argument-free, trace-disabled macro without temporary harness injection; plain `--direct` auto-disables default compile diagnostics |
| `--fast`               | Uses direct execution when eligible and diagnostic mode is disabled, otherwise falls back to normal run                                     |
| `--diagnostic`         | Explicitly keeps structured compile diagnostics enabled (default true)                                                                      |
| `--gui-compile-errors` | Opts out of structured compile diagnostics and allows Excel/VBE compile dialogs to appear                                                   |
| `--session`            | Uses the workbook opened by `xlflow session start`                                                                                          |
| `--timeout 5m`         | Stops execution if it does not complete in time and returns `macro_timeout`                                                                 |

### `xlflow session`

Keeps Excel and the configured workbook open between commands:

```bash
xlflow session start
xlflow pull --session --json   # when the workbook may be newer than src/
xlflow push --fast --session --no-save --json
xlflow run Main.Run --headless --session --json
xlflow save --session --json
xlflow session stop
```

`--session` remains the explicit assertion mode. When `.xlflow/session.json` already points at the configured workbook, plain `list forms`, `inspect form`, `form snapshot`, `form build`, `pull`, `push`, `macros`, `run`, `export-image`, `form export-image`, `test`, `trace`, and `save` auto-reuse that matching live workbook and report that reuse in JSON and human output.

When `push --session --no-save` succeeds, or `run --session` completes without `--save` / `--save-as`, the live workbook may differ from the `.xlsm` on disk until `xlflow save --session`.
If UserForms are involved, treat that save step as part of review hygiene before comparing `.frm` / `.frx` output.
xlflow now warns more aggressively about this unsaved session state, but `xlflow save --session` remains the canonical persistence step before `session stop`.

</details>

<details open>
<summary><strong>Debugging and testing: <code>trace</code>, <code>test</code>, <code>diff</code></strong></summary>

### `xlflow trace`

Collects log events from VBA during macro execution.

Enable the trace module when you want it persisted in the workbook and source tree:

```bash
xlflow trace enable --json
xlflow trace status --json
```

During session-based development, add `--session` to `trace enable`, `trace status`, `trace disable`, and traced `run` commands.

Write logs from VBA:

```vb
Call XlflowLog("start GenerateReport")
Call XlflowLog("lastRow=" & lastRow)
Call XlflowLog("finished GenerateReport")
```

Run the macro with trace enabled:

```bash
xlflow run Main.Run --trace --json
```

Trace events are returned in the top-level JSON `trace` field.
This helps identify how far execution progressed before a runtime error.

`xlflow run --trace` can temporarily inject and revert the helper when it is missing.
Human output and JSON `trace.lifecycle` distinguish that temporary path from a helper that is already persisted in workbook/source state.
Trace logs are written under `.xlflow/traces`.
Use `xlflow trace disable --json` to remove the persistent helper and `xlflow trace clean --json` to remove trace log files.
`xlflow trace inject` remains as a compatibility alias for `trace enable`.

### `xlflow test`

Runs VBA tests.

```bash
xlflow test --json
```

xlflow discovers argument-free `Sub` procedures whose names start with `Test` or end with `_Test`.
Use `xlflow test --session --json` when an xlflow session is already open.

To run a single test, use `--filter`:

```bash
xlflow test --filter TestCreateReport --json
```

New and initialized projects include `src/modules/XlflowAssert.bas`.
Use `AssertEquals expected, actual, [message]` to compare scalar values.

```vb
Public Sub TestCreateReport()
    AssertEquals 10, Sheets("Result").Range("A1").Value2
End Sub
```

> [!NOTE]
> `AssertEquals` does not support object or array comparison. Compare scalar properties such as `Range.Value2` instead of passing `Range` objects directly.

### `xlflow diff`

Compares two workbooks.

```bash
xlflow diff before.xlsm after.xlsm --json
```

It detects sheet additions/removals, cell value differences, and formula differences.

To compare exported VBA source as well, pass `--vba-before` and `--vba-after`:

```bash
xlflow diff before.xlsm after.xlsm \
  --vba-before before-src \
  --vba-after after-src \
  --json
```

> [!IMPORTANT]
> Differences are reported as successful command results. `diff` returns exit code `0` even when differences are found. Inspect `diff.summary.total_diffs` in JSON to determine whether anything changed.

</details>

<details open>
<summary><strong>Quality gates: <code>lint</code>, <code>analyze</code>, <code>check</code>, <code>inspect-gui</code></strong></summary>

### `xlflow lint`

Lints VBA source.

```bash
xlflow lint --json
```

It detects patterns that are unsafe or inconvenient for AI agents and unattended automation:

- Missing `Option Explicit`
- `Select` usage
- `Activate` usage
- `On Error Resume Next` usage
- Possible implicit `Variant`
- Module-level `Public` variables
- Interactive operations such as `Application.GetOpenFilename`, `Application.FileDialog`, `InputBox`, and modal `MsgBox`

### `xlflow analyze`

Analyzes VBA source for runtime-risk patterns without opening Excel.

```bash
xlflow analyze --json
```

The first analyzer rules report likely missing `Set` assignments for object variables and object-returning functions.
Findings are returned in top-level `analysis` with file, module, procedure, line, nearby code, reason, and suggestion.

### `xlflow check`

Runs the standard preflight sequence:

```bash
xlflow check --keepalive --json
```

`check` runs `lint`, `analyze`, and `doctor`, then returns an aggregate top-level `check` object.
It continues after lint/analyze findings so the report includes all cheap source feedback before the Excel COM doctor result.

### `xlflow inspect`

Reads the saved workbook file directly without opening Excel.

```bash
xlflow inspect workbook --json
xlflow inspect sheets --format markdown
xlflow inspect range --sheet "Result" --address "A1:F20" --json
xlflow inspect range --sheet "QR" --address "A1:AE31" --include-style --json
xlflow inspect used-range "Result" --max-rows 50 --max-cols 10 --format markdown
xlflow inspect cell "Result!B3" --json
```

Use it to inspect workbook structure and cell output after `push` / `run` workflows when the workbook state has been saved to disk.
If exported `.frm` files exist under `src/forms`, `inspect` warns that it did not verify live UserForm Designer/runtime state.
`inspect` is a file snapshot reader, so unsaved changes in an already-open Excel window are intentionally out of scope for this command family.
Add `--include-style` on `inspect range` or `inspect used-range` when the workbook meaning depends on fill colors, borders, merged cells, row heights, or column widths.

### `xlflow export-image`

Exports a worksheet range as a PNG through Excel COM.

```bash
xlflow export-image --sheet "QR" --range "A1:AE31" --json
xlflow export-image --sheet "QR" --range "A1:AE31" --out artifacts\qr.png --overwrite --json
```

This is the visual verification companion to `inspect`. Use it when workbook correctness depends on charts, fills, layout, printable forms, QR-code cells, or other rendering details that a saved-file snapshot is not enough to prove.

Without `--out`, xlflow writes under `.xlflow/artifacts/images/<workbook-name>/` using a generated filename. `--output-dir` selects only the directory, and `--name` selects only the filename. Only PNG is supported in v1.

Like other workbook-backed commands, `export-image` auto-reuses a matching recorded session workbook when `.xlflow/session.json` points at the configured workbook. Add `--session` when you want that requirement to be explicit. Successful JSON includes top-level `target`, `output`, and optional `warnings`.

### `xlflow edit`

Mutates the live workbook held by an xlflow session for development-time setup and visual tuning.

```bash
xlflow edit cell --sheet "Input" --cell "B2" --value "ABC123" --events on --session --json
xlflow edit range --sheet "QR" --range "A1:AE31" --fill "#FFFFFF" --session --json
xlflow edit rows --sheet "QR" --rows "1:31" --height 12 --session --json
xlflow edit columns --sheet "QR" --columns "A:AE" --width 2.2 --session --json
```

Use it to prepare workbook state before `run`, trigger `Worksheet_Change` handlers with `--events on`, clear or repaint regions between iterations, and tune row heights or column widths before exporting an image. MVP `edit` is intentionally session-only: `--session` is required, successful edits mark the live workbook dirty, and you must run `xlflow save --session` to persist the changes to disk.

### `xlflow inspect-gui`

Reports GUI interaction boundaries without opening Excel.

```bash
xlflow inspect-gui --json
```

The report includes file, line, kind, symbol, and a suggested refactor.
Use it before deciding whether a macro should run with `--headless` or `--interactive`.

</details>

<details open>
<summary><strong>AI agent support: <code>skill install</code></strong></summary>

### `xlflow skill install`

Installs the bundled xlflow Skill for AI agents.

```bash
xlflow skill install --agent codex
xlflow skill install --agent claude
xlflow skill install --agent cursor
xlflow skill install --agent gemini
xlflow skill install --target .agents/skills
```

Supported provider targets are:

| Agent target | Install path            |
| ------------ | ----------------------- |
| `agents`     | `.agents/skills/xlflow` |
| `codex`      | `.codex/skills/xlflow`  |
| `claude`     | `.claude/skills/xlflow` |
| `cursor`     | `.cursor/skills/xlflow` |
| `gemini`     | `.gemini/skills/xlflow` |

For GitHub Copilot-style workflows, use the shared `.agents` target:

```bash
xlflow skill install --agent agents
```

</details>

---

## Configuration

xlflow reads `xlflow.toml` from the project root.

```toml
[project]
name = "sample"
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"
visible = false
display_alerts = false

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

`project.entry` is used when `xlflow run` is invoked without a macro name.

Set `forbid_interactive_input = false` when the project intentionally uses dialogs or UserForms and you want to suppress `VB007` warnings. This only affects lint output; `xlflow run --headless` still blocks GUI boundaries.

Syntax safety lint rules for typographic quotes, C-style quote escapes, unclosed or mismatched procedures, and malformed line-continuation underscores are always enabled because they prevent VBE compile dialogs before `push` or `run` opens Excel.

---

## JSON output

Every command can return AI-agent-friendly JSON by passing `--json`.

The basic envelope is:

```json
{
  "status": "ok",
  "command": "lint",
  "error": null,
  "logs": []
}
```

On failure, `status` is `failed`, and `error.code` and `error.message` are returned.

```json
{
  "status": "failed",
  "command": "run",
  "error": {
    "code": "macro_failed",
    "message": "Main Err 5: inputPath is required",
    "source": "Main",
    "number": 5,
    "phase": "invoke_macro"
  },
  "logs": []
}
```

> [!TIP]
> AI agents and automation scripts should treat `status`, `command`, `error.code`, and command-specific top-level fields as the primary contract.

---

## Exit codes

| Code | Meaning                                                             |
| ---: | ------------------------------------------------------------------- |
|  `0` | Success                                                             |
|  `1` | Validation failure, such as lint, macro, or test failure            |
|  `2` | CLI argument or configuration error                                 |
|  `3` | Environment error, such as Excel, COM, VBIDE, or PowerShell failure |

> [!NOTE]
> `diff` returns exit code `0` even when differences are found. Inspect `diff.summary.total_diffs` to determine whether inputs differ.

---

## Recommended VBA rules

VBA executed by xlflow should be written for unattended automation.

- [x] Always use `Option Explicit`
- [x] Use explicit `Workbook`, `Worksheet`, and `Range` references
- [x] Prefer `Long` over `Integer`
- [x] Keep GUI entrypoints thin
- [x] Extract parameterized headless procedures for the core logic
- [x] Pass input values through `xlflow run --arg`, configuration files, deterministic paths, or environment variables
- [x] Emit error messages that make failures diagnosable
- [x] Verify destructive workbook changes with tests or diff
- [ ] Do not rely on `Select`, `Activate`, or `ActiveSheet`
- [ ] Do not depend on UI dialogs or modal `MsgBox` in headless procedures
- [ ] Avoid broad `On Error Resume Next`

---

## Local verification

Run repository linters with:

```bash
task lint
```

`task lint` runs `golangci-lint run` and `PSScriptAnalyzer` against tracked `.ps1` sources.
Make sure `Invoke-ScriptAnalyzer` is available in your local PowerShell environment.

Run the fast repository verification with:

```bash
task verify
```

Currently, `task verify` runs `go test ./...` as non-COM test coverage.

Excel COM E2E verification should be run on Windows with Excel and VBIDE access enabled.

```bash
xlflow doctor --json
```

After `doctor` reports a healthy environment, run `new`, `doctor`, `pull`, `lint`, `push`, `run`, `test`, and `diff` against a real workbook.

---

## Current status

xlflow is an MVP-stage tool.

Its primary goal is to bring Excel VBA into AI-agent and CLI-based development workflows.
Typical use cases include:

| Use case                           | Why xlflow helps                                                |
| ---------------------------------- | --------------------------------------------------------------- |
| Source control for existing VBA    | VBA modules become normal files                                 |
| AI-agent-assisted VBA modification | Agents can edit source, run checks, and inspect JSON output     |
| CLI execution of Excel macros      | Macros can be invoked from scripts and terminals                |
| Automated VBA testing              | Tests can be discovered and executed consistently               |
| Debugging with runtime logs        | Trace events show how far execution progressed                  |
| Workbook change review             | `diff` makes workbook changes easier to inspect                 |
| Internal Excel automation          | Existing VBA assets can move toward safer development workflows |

> [!CAUTION]
> xlflow is useful, but it cannot make every legacy workbook safely headless. GUI-heavy macros, workbook-level side effects, external dependencies, and fragile Excel state still need deliberate refactoring.

---

## License

MIT
