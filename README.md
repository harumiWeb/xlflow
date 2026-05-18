<p align="center">
    <img width="600" alt="xlflow logo" src="docs/images/logo.png" />
</p>

<p align="center">
  <em>Excel VBA development, rebuilt for CLI-first humans and AI agents.</em>
</p>

<p align="center">
  <a href="https://harumiweb.github.io/xlflow/">Official Documentation</a>
</p>

<p align="center">
  <a href="README.md">English</a>
  |
  <a href="README.ja.md">日本語</a>
</p>

<div align="center">

![GitHub Release](https://img.shields.io/github/v/release/harumiWeb/xlflow?include_prereleases) ![WinGet Package Version](https://img.shields.io/winget/v/HarumiWeb.Xlflow) ![Scoop](https://img.shields.io/scoop/v/xlflow?bucket=https%3A%2F%2Fgithub.com%2FharumiWeb%2Fscoop-bucket) ![GitHub License](https://img.shields.io/github/license/harumiWeb/xlflow) ![GitHub go.mod Go version](https://img.shields.io/github/go-mod/go-version/harumiWeb/xlflow) [![Ask DeepWiki](https://deepwiki.com/badge.svg)](https://deepwiki.com/harumiWeb/xlflow)

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
| Existing UserForm layout is hard to review in diffs  | Persist Designer state with `xlflow form snapshot`             |
| Runtime failures are hard to locate                  | Return structured errors, diagnostics, and trace logs          |
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

### winget

```powershell
winget install HarumiWeb.Xlflow
```

Use `upgrade` to update an existing installation:

```powershell
winget upgrade HarumiWeb.Xlflow
```

> [!NOTE]
> winget availability may lag behind a GitHub Release while the manifest is submitted and accepted upstream.
> Use Scoop or the GitHub Releases ZIP when you need the newest release immediately.

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

### Go install

```bash
go install github.com/harumiWeb/xlflow/cmd/xlflow@latest
```

`go install` may contact the Go module mirror and checksum database configured in your Go environment. For direct source checkout development and CI, treat the Go version declared in `go.mod` as the supported toolchain source of truth; the repository CI and release workflows resolve Go from that file.

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

`new` automatically pushes the scaffolded VBA modules into the new workbook, so a later `pull` starts from the same initial source.

Or start from an existing workbook:

```bash
xlflow init Book.xlsm
```

`init` automatically pulls VBA out of the copied workbook into `src/`, so you can edit source files immediately without a separate bootstrap `pull`.

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

### Letting an AI agent edit VBA

#### Installing the Skill

When having an AI agent edit VBA using xlflow, it is recommended to install the Skill provided by xlflow into the agent's environment.

```bash
xlflow skill install
```

You can also install it at the same time you launch a project.

```bash
xlflow new Book.xlsm --with-skill
```

If you wish to manage skills using a manager such as vercel-labs/skills, please install the skill as follows.

```bash
npx skills add harumiWeb/xlflow/internal/agentskill/templates --skill xlflow
```

#### Creating a project

While you can leave the creation of the project itself to an AI agent, it is recommended that the initial project setup be performed by a human.

```bash
xlflow new Book.xlsm --with-skill
```

#### Letting an AI agent edit

Using the installed skill, please provide instructions for what you want to achieve in natural language.

```bash
/xlflow Create a macro that enters "Hello, world!" into cell A1 using VBA
```

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
| `backup list`       | List rollback-capable workbook backups                      | `xlflow backup list --json`                                                  |
| `pull`              | Export VBA components into `src/`                           | `xlflow pull --json`                                                         |
| `push`              | Import VBA source back into the workbook                    | `xlflow push --json`                                                         |
| `rollback`          | Restore the workbook from a saved backup                    | `xlflow rollback --latest --json`                                            |
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

Detailed command behavior, options, JSON payloads, and troubleshooting notes now live in the documentation site.

- [Command reference](https://harumiweb.github.io/xlflow/commands/)
- [JSON output](https://harumiweb.github.io/xlflow/reference/json-output)
- [Configuration](https://harumiweb.github.io/xlflow/reference/config-file)
- [Troubleshooting](https://harumiweb.github.io/xlflow/reference/troubleshooting)

Use the README as a quick overview and the documentation site as the source for command-level details.

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

## License

MIT License. See [LICENSE](LICENSE).
