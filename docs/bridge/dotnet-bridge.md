# .NET Excel Bridge

The .NET Excel bridge is the Windows automation provider for workbook-backed xlflow commands.

It is introduced by ADR-0008 and uses the provider/fallback contract from ADR-0009.

## Responsibilities

The .NET bridge owns Windows automation concerns that are difficult to keep reliable in PowerShell:

- Excel COM automation.
- VBIDE automation.
- Workbook open/save/close and session attachment.
- Macro execution.
- UserForm import/export.
- Win32 and UI Automation dialog watching.
- Compile/runtime error capture.
- Process/window ownership checks.
- Clipboard and image export fallback behavior.

The Go CLI remains responsible for command parsing, config loading, source-tree decisions, static checks, public envelope mapping, and provider selection.

## Implemented Commands

### `doctor`

`xlflow doctor --bridge dotnet --json` runs environment diagnostics through the .NET bridge without launching PowerShell. The response includes a `diagnostics` object at the top level:

```json
{
  "status": "ok",
  "command": "doctor",
  "diagnostics": {
    "selected_bridge": "dotnet",
    "protocol_version": 1,
    "runtime": {
      "os": "Windows 11",
      "process_architecture": "X64",
      "dotnet_runtime": ".NET 8.0"
    },
    "excel": {
      "com_activation": true,
      "version": "16.0",
      "build": "12345",
      "vbide_access": true,
      "automation_security": 1,
      "trust_vba_access": null
    }
  }
}
```

When Excel COM activation fails, the bridge returns a structured error with `code`, `message`, `phase`, `source`, `number`, `h_result`, and `details` fields. The `h_result` field is a hex HRESULT string (e.g. `"0x80040154"`).

Fields:

- `selected_bridge` — always `"dotnet"` when the .NET bridge handles the command.
- `protocol_version` — bridge protocol version (currently `1`).
- `runtime.os` — `Environment.OSVersion` string.
- `runtime.process_architecture` — process architecture (e.g. `X64`).
- `runtime.dotnet_runtime` — `.NET` runtime description.
- `excel.com_activation` — `true` when `Excel.Application` was successfully created on the STA thread established by `[STAThread]` in `Program.cs`. This encompasses both COM object creation and the STA execution context required by Excel.
- `excel.version` — Excel application version.
- `excel.build` — Excel application build number.
- `excel.vbide_access` — `true` when the VBA project object model is accessible.
- `excel.automation_security` — observed `AutomationSecurity` value.
- `excel.trust_vba_access` — observed Trust access state; `null` when not determinable.
- `excel.error` — present only when a non-fatal diagnostic warning occurred.

### `inspect`

The .NET bridge now supports the session-backed read-only inspection targets used by the Excel runner:

- `xlflow inspect workbook --session --bridge dotnet --json`
- `xlflow inspect sheets --session --bridge dotnet --json`
- `xlflow inspect range --session --bridge dotnet --json`

These commands preserve the existing top-level envelope fields produced by the Go CLI:

- `target`
- `session`
- `workbook`
- `inspect`
- `logs`

The `.NET` implementation intentionally stays limited to the live/session-backed runner path for now. Saved-file inspect flows still use the existing non-bridge workbook reader.

### `macros`

`xlflow macros --bridge dotnet --json` discovers VBA macro entrypoints from the configured workbook through the .NET bridge:

```powershell
xlflow macros --bridge dotnet --json
xlflow macros --bridge dotnet --session --json
xlflow macros --bridge dotnet --entry Module1.Main --json
xlflow macros --bridge dotnet --runnable-only --json
```

The command opens the workbook (or attaches to an active session), iterates VBComponents, and parses each code module for `Sub`/`Function` declarations:

- Skips `Private` and `Friend` procedures
- Skips `Xlflow*` prefixed modules
- Skips event procedures (`Workbook_*`, `Worksheet_*`, `Auto_Open`, `Auto_Close`)
- Skips procedures with parameters
- Marks procedures in `userform`, `document_module`, or `unknown` component types as not runnable

The response includes `target`, `session`, `workbook`, `macros`, `default_entry`, `suggestions`, `warnings`, and `hints` envelope fields matching the PowerShell macros contract.

Each macro entry contains: `module`, `name`, `qualified_name`, `kind`, `args`, `line`, `component_type`, `visibility`, `has_parameters`, `runnable`, `reason_not_runnable`, and `run_command`.

### `run`

`xlflow run Module1.Main --bridge dotnet --json` executes a VBA macro through the .NET bridge:

```powershell
xlflow run Module1.Main --bridge dotnet --json
xlflow run Module1.Main --bridge dotnet --session --json
xlflow run Module1.Main --bridge dotnet --save --json
xlflow run Module1.Main --bridge dotnet --save-as Result.xlsm --json
xlflow run Module1.Main --bridge dotnet --no-save --json
xlflow run Module1.Main --bridge dotnet --trace --json
xlflow run Module1.Main --bridge dotnet --msgbox '{"prompt":"ok"}' --inputbox '{"name":"Jane"}' --ui-stream --json
```

The command opens the workbook (or attaches to an active session), invokes the macro via `Application.Run`, and captures the result:

- Supports typed arguments (`string`, `int`, `double`, `bool`) via base64-encoded JSON
- Supports fully qualified macro names (`Module1.Main`)
- Captures `Err.Number`, `Err.Description`, and execution duration
- Supports `--save`, `--save-as`, and default no-save behavior
- Supports `--timeout` via bridge request timeout
- Returns `macro_failed`, `macro_not_found`, or `macro_disabled` structured errors
- Runs modal `Application.Run` and VBE Compile calls in disposable child bridge
  processes so the parent can watch dialogs and produce diagnostics
- Detects Excel/VBE dialogs through Win32 owner-chain polling with optional UI
  Automation metadata and button invocation
- Captures dialog text, buttons, HWND/PID/thread identity, visibility, action,
  and worker state in `run_diagnostic`
- Suppresses runtime error dialogs with the explicit End action by default;
  Debug is not preferred because it can leave VBE in break mode
- Does not require a dialog to be visible, because Excel can defer painting a
  runtime error dialog until it receives focus
- Supports `--trace` for VBA trace collection: temporarily injects `XlflowTrace` module if not present, reads trace events after execution, and reverts temporary injection
- Supports `--msgbox`, `--inputbox`, `--filedialog` response injection via defined names
- Supports `--ui-stream` pipe integration and `__XLFLOW_DEBUG_PIPE__` injection for `XlflowDebug.Log` transport

The response includes `target`, `session`, `workbook`, `macro`, `runtime`, `trace`, `run_diagnostic`, `suggestions`, and `warnings` envelope fields.

### `test`

`xlflow test --bridge dotnet --json` runs VBA tests through the .NET bridge:

```powershell
xlflow test --bridge dotnet --json
xlflow test --bridge dotnet --session --json
xlflow test --bridge dotnet --filter TestSomething --json
xlflow test --bridge dotnet --module Module1 --json
xlflow test --bridge dotnet --tag smoke --json
```

The command:

1. Opens the workbook (or attaches to an active session)
2. Injects runtime markers and UI/debug stream helpers
3. Discovers test procedures (`Test*` or `*_Test` pattern, public/implicit `Sub`)
4. Collects `@Tag("...")` annotations from preceding comment lines
5. Discovers `BeforeAll`/`AfterAll`/`BeforeEach`/`AfterEach` hooks
6. Generates and injects a runner module with per-test dispatch
7. Executes BeforeAll, each test, AfterAll per module
8. Restores runtime markers and saves the workbook

Test results use `vbObjectError + 516` as the inconclusive sentinel. The response includes `workbook` and `tests` envelope fields.

### `trace`

`xlflow trace enable|status|disable|clean --bridge dotnet --json` manages the `XlflowTrace` VBA helper module:

```powershell
xlflow trace enable --bridge dotnet --json
xlflow trace status --bridge dotnet --json
xlflow trace disable --bridge dotnet --json
xlflow trace disable --bridge dotnet --force --json
xlflow trace clean --bridge dotnet --json
```

- `enable` installs the `XlflowTrace` module and saves the workbook; optionally writes `XlflowTrace.bas` to `ModulesDir`
- `disable` validates source-match safety before changing the workbook, then removes the module and saves; removes source file only if it matches the bundled helper or `--force` is used
- `status` reports `workbook_injected`, `source_exists`, and `source_matches_bundled`
- `clean` removes the trace log directory without opening Excel

The trace direct-open path disables automation macros for safety.

On timeout, the bridge returns `macro_timeout`, does not save the workbook, and
includes actionable suggestions plus the dialog and worker state available at
the timeout boundary. Excel may continue executing user VBA after the child COM
caller is terminated, so COM-based harness cleanup is not attempted
synchronously while Excel is busy. Callers must treat timeout results as
`vba_may_still_be_running` until the Excel session is reset or the workbook is
reopened. When the .NET bridge cannot finish returning its own timeout payload
before the outer Go deadline, xlflow still returns a valid JSON timeout envelope
from Go with the same limitation documented explicitly.

### `process`

The .NET bridge now supports:

- `xlflow process list --bridge dotnet --json`
- `xlflow process cleanup --bridge dotnet --json`

Behavior matches the PowerShell contract:

- `process list` returns `process: [{ pid, has_workbook }]`
- `process cleanup --auto` targets only workbook-free Excel processes
- `process cleanup --all` force-stops the selected Excel processes
- partial cleanup failures return structured `process_termination_failed` errors while preserving the `process` result payload

### `pull`

`xlflow pull --bridge dotnet --json` exports VBA components from the configured workbook through the .NET bridge:

```powershell
xlflow pull --bridge dotnet --json
```

The command opens the workbook (or attaches to an active session), iterates VBComponents, and writes each component to the configured source directory:

| Component Type           | Target Dir     | Extension |
| ------------------------ | -------------- | --------- |
| Standard module (type 1) | `src/modules`  | `.bas`    |
| Class module (type 2)    | `src/classes`  | `.cls`    |
| UserForm (type 3)        | `src/forms`    | `.frm`    |
| Document module (other)  | `src/workbook` | `.bas`    |

The response includes `target`, `session`, `workbook`, and `source` envelope fields matching the PowerShell pull contract.

### `push`

`xlflow push --bridge dotnet --json` imports source VBA components into the configured workbook through the .NET bridge:

```powershell
xlflow push --bridge dotnet --json
xlflow push --bridge dotnet --fast --session --no-save --json
```

The command:

1. Attaches to the session workbook (or opens the configured workbook)
2. Creates a backup when `BackupMode` is not `"never"` (via `SaveCopyAs`)
3. Discovers source files from `ModulesDir`, `ClassesDir`, `FormsDir`, `WorkbookDir`
4. Removes existing components by name match, then imports via `VBComponents.Import`
5. Returns `target`, `session`, `workbook`, `backup`, and `source` envelope fields

## Process Model

The bridge is a separate executable named `xlflow-excel-bridge.exe`.

The Go CLI starts it for command-scoped work and communicates through stdin/stdout JSON. stdout is reserved for protocol JSON only. Diagnostic logs go into response `logs` or stderr.

## COM Threading Rule

Excel COM operations must run in an STA context.

Implementation rules:

- Entry points that touch Excel COM must be `[STAThread]` or dispatch to a dedicated STA thread.
- Do not resume Excel COM work on arbitrary ThreadPool threads after `await`.
- Keep initial COM command handlers synchronous unless a dedicated STA dispatcher exists.
- Release COM references deliberately and keep workbook/session ownership explicit.

This rule is part of the public implementation contract because violating it can produce intermittent Excel/VBIDE failures that unit tests may not catch.

## Capabilities

The bridge advertises supported command keys through:

```powershell
xlflow-excel-bridge.exe --capabilities-json
```

The Go resolver must check capabilities before selecting .NET in `auto` mode.

For issue #76, `auto` may select `.NET` for `doctor`, `inspect`, and `process`. `pull`, `push`, `macros`, and `run` are intentionally excluded from auto mode to preserve the existing PowerShell behavior; use `--bridge dotnet` explicitly to route these commands through the .NET bridge. If `.NET` reports capability/runtime/protocol failures for a supported command, the Go side may fall back to PowerShell only in `auto` mode. Explicit `--bridge dotnet` remains strict.

## Selection

Users will be able to select the bridge through:

```powershell
xlflow doctor --bridge dotnet --json
xlflow run Main.Run --bridge dotnet --json
```

When `--bridge dotnet` is selected explicitly, xlflow must not fallback to PowerShell. Missing bridge executable, version mismatch, protocol mismatch, or unsupported command must be reported as structured errors.

## Initial Migration Order

The planned migration order is:

1. `doctor` — done
2. `pull` and `push` — done
3. `macros` and `run` — done
4. dialog watcher — done
5. `test`, `trace`, and runtime injection — done
6. `form` and `export-image`
7. release packaging
8. default bridge switch

`run` is intentionally not first because it combines macro invocation, runtime injection, compile checks, dialog capture, timeout behavior, and session handling.

## Release Notes

The .NET bridge avoids PowerShell execution policy, but it does not bypass all corporate controls. AppLocker, WDAC, Defender, EDR, or code-signing rules may still block an unsigned or unapproved executable. Release packaging and signing decisions must be documented before making .NET the default Windows bridge.
