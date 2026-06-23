---
name: xlflow-tmp-workspace-e2e
description: Create disposable `xlflow` projects under this repository's `tmp_workspaces` directory and verify the Windows Excel COM flow end-to-end. Use when Codex needs to validate `xlflow` behavior with real workbooks, run the standard `new/init/doctor/pull/push/run/lint` workflow, confirm workbook state after macro execution, or reproduce/fix CLI integration bugs without touching the main project tree.
---

# xlflow tmp_workspaces E2E

Execute the repository-standard `xlflow` verification flow in isolated workspaces under `tmp_workspaces`.

Follow the workflow exactly unless the user narrows scope.

When the task is release preparation, treat this skill as a release gate rather than an optional spot check.

Prefer a session-first Excel workflow for any real-workbook verification that needs more than one Excel COM-backed command. Reopening the workbook separately for `push`, `run`, `test`, `pull`, `save`, or repeated inspect steps frequently makes validation much slower and can stall or look hung. Unless the check is truly one-shot, start a session once, reuse it, and stop it only after the verification slice is complete.

## Preconditions

- Read `tasks/lessons.md` before creating the workspace.
- Work from the repository root: `C:\dev\go\xlflow`.
- Use `xlflow` from PATH, not `go run`, unless the user explicitly asks otherwise.
- Prefer a fresh workspace per verification attempt. Do not reuse an old `tmp_workspaces` project unless the user asks to inspect a prior artifact.

## Workspace Rules

Create one or two fresh directories under `tmp_workspaces`:

- primary E2E workspace:
  - `tmp_workspaces/<topic>-e2e`
- optional init workspace:
  - `tmp_workspaces/<topic>-init`

Before using a workspace:

- remove the old directory if it exists and the user did not ask to preserve it
- recreate the directory empty
- record the exact absolute path for reporting

## Standard Flow

### 0. Session-first rule for Excel COM verification

For blank scaffold checks such as `new`, `doctor`, first `pull`, and first `lint`, normal one-shot commands are fine.

For any workflow that chains workbook-backed execution or mutation commands, prefer this shape:

```powershell
xlflow session start --json
xlflow push --fast --session --no-save --json
xlflow run <MacroName> --session --json
xlflow test --session --json
xlflow save --session --json
xlflow session stop --json
```

Adjust the middle commands to the scenario, but keep the same principle: reuse one live workbook session instead of reopening Excel for each step. Only skip the session when the user explicitly asks for saved-file-only validation or when the path is intentionally verifying non-session behavior.

### 1. Confirm tool resolution

Run `Get-Command xlflow` and confirm the resolved binary path.

If the installed binary must reflect current source changes, rebuild with:

```powershell
go install ./cmd/xlflow
```

### 2. Blank-workbook verification

In the primary workspace, run this sequence:

```powershell
xlflow new --json
xlflow doctor --json
xlflow pull --json
xlflow lint --json
```

Verify all commands return success.

After `new`, confirm the scaffold exists:

- `xlflow.toml`
- `build/Book.xlsm` or the requested workbook
- `src/modules`
- `src/classes`
- `src/forms`
- `src/workbook`
- `.xlflow`

After `pull`, inspect exported workbook modules under `src/workbook`.

Pay special attention to the document-module normalization rule from `tasks/lessons.md`:

- `ThisWorkbook.bas` and worksheet modules must not retain `VERSION`, `BEGIN`, `MultiUse`, `END`, or `Attribute VB_` lines in source form
- blank workbook modules should still lint cleanly

### 3. Macro round-trip verification

Use this path when `push`, `run`, workbook save behavior, or VBA round-trip behavior matters.

Strong default: use a session for this flow. Do not run `push`, `run`, and follow-up workbook checks as separate fresh Excel opens unless the point of the test is specifically to verify non-session behavior.

Create a minimal runnable module at `src/modules/Main.bas`:

```vb
Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    ThisWorkbook.Worksheets(1).Range("A1").Value = "xlflow ok"
End Sub
```

Then run:

```powershell
xlflow lint --json
xlflow session start --json
xlflow push --fast --session --no-save --json
xlflow run Main.Run --session --json
xlflow save --session --json
xlflow session stop --json
xlflow pull --json
xlflow lint --json
```

Verify:

- `push` reports imported source files and workbook-module updates when expected
- `run` reports the requested macro name
- session-backed `save` / `stop` leave the workbook clean without stray save-required state
- `pull` exports `Main.bas` plus workbook modules
- the final `lint` still passes

### 4. Workbook-state verification

When macro behavior should persist to the workbook, verify it through Excel COM instead of trusting CLI success output alone.

If the behavior under test already uses a session, keep that session for the workbook mutation commands and do the direct COM state check only after `save --session` or `session stop`, unless the test explicitly needs to inspect the live unsaved workbook.

Use a PowerShell check like:

```powershell
$path = 'C:\dev\go\xlflow\tmp_workspaces\<topic>-e2e\build\Book.xlsm'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($path)
try {
  $wb.Worksheets.Item(1).Range('A1').Value2
} finally {
  $wb.Close($false)
  $excel.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
  [gc]::Collect()
  [gc]::WaitForPendingFinalizers()
}
```

Adapt the cell or workbook assertions to the behavior under test.

### 5. init verification

Use a second fresh workspace when `init` must be covered.

Run:

```powershell
xlflow init <absolute-path-to-existing-workbook> --json
xlflow pull --json
```

Verify the copied workbook exists under the second workspace `build/` directory and that exported components appear under `src/`.

### 6. Release-gate coverage

Use this section when the user asks for release readiness, release preflight, or "before release" verification.

Release-gate verification must cover:

- blank workbook scaffold:
  - `new`
  - `doctor`
  - `pull`
  - `lint`
- standard module round-trip:
  - `push`
  - `run`
  - workbook-state verification through Excel COM
  - `pull`
  - `lint`
- class module round-trip
- UserForm round-trip, including `.frm` and `.frx`
- `init` from an existing workbook

When the release includes `pack` changes, also cover the pack artifact smoke described in section 7:

- produce a `.xlsm` with `xlflow pack` (file-level, no Excel)
- open the produced artifact in Excel and run a packed macro
- assert an observable effect, such as a sentinel cell value, to confirm the packed VBA compiled and ran

When the release includes session-related changes, also cover:

- `session start`
- `push --fast --session --no-save`
- `run --session` and/or `test --session`
- `save --session`
- `session stop`
- explicit confirmation of any save-required warnings or session metadata that changed

Even when the change is not primarily about session state, prefer the same session-backed execution pattern for multi-step Excel COM release checks unless a specific saved-file reopen path is itself under test.

Do not report release readiness if one of these paths was skipped without calling it out explicitly.

### 7. pack artifact smoke

Use this section when the release includes changes to the experimental `pack` command (`docs/specs/pack-command.md`, ADR-0012).

`pack` is a pure-Go, file-level path: it regenerates `xl/vbaProject.bin` from source without opening Excel, so it always reports `pack.vbe_validation = "not_performed"`. This smoke exists to confirm, at the release gate, that a representative packed artifact actually compiles and runs in real Excel. It is release-gate evidence only; it does not run in PR CI and it does not change `pack`'s permanent no-validation contract.

Keep the automated PR path Linux/pure-Go only. Do not add Excel or COM steps to PR CI; the smoke is performed here, in the release gate, on Windows with Excel.

In the primary workspace, make the standard `src/modules/Main.bas` macro from section 3 write a known sentinel value to `A1`. Keep the sentinel in a module that already exists in the template (for example `Main`): the experimental `pack` MVP updates existing modules only and rejects a brand-new module with `pack_ambiguous_layout`. Produce the artifact at the file level:

```powershell
xlflow pack --out dist/Book.xlsm --experimental --json
```

Confirm the JSON envelope reports `status = "ok"`, `pack.backend = "pure-go"`, and `pack.vbe_validation = "not_performed"`.

Open the produced artifact, run the packed macro, and assert the sentinel cell:

```powershell
$path = 'C:\dev\go\xlflow\tmp_workspaces\<topic>-e2e\dist\Book.xlsm'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true   # keep Excel visible: a hidden VBE compile-error dialog otherwise looks like a hang
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1  # msoAutomationSecurityLow: allow the packed macros to run
$wb = $excel.Workbooks.Open($path)
try {
  $excel.Run('Main.Run')                      # forces a VBE compile, then runs the packed macro
  $sentinel = $wb.Worksheets.Item(1).Range('A1').Value2
  if ($sentinel -ne 'xlflow ok') {            # the value Main.Run writes in section 3
    throw "pack smoke failed: A1 was '$sentinel', expected 'xlflow ok'"
  }
} finally {
  $wb.Close($false)
  $excel.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
  [gc]::Collect()
  [gc]::WaitForPendingFinalizers()
}
```

The smoke passes only when:

- `pack` exits `0` and produces the `.xlsm`
- the packed workbook opens in Excel without a compile error
- `Run` executes the packed macro and the sentinel cell holds the expected value

A compile error on open, a missing or wrong sentinel value, or a non-zero `pack` exit is a release blocker. If `Run` appears to hang, bring the (visible) Excel window to the foreground — a VBE compile-error dialog is the usual hidden cause. Record the `pack` JSON output and the observed sentinel value in the release-gate report.

## Failure Handling

Treat verification failures as product bugs unless evidence shows test setup is wrong.

When a command fails:

1. capture the exact failing command and JSON output
2. inspect generated files or workbook state
3. identify root cause before editing code
4. add or update tests when fixing the product
5. rerun the affected E2E path from a fresh workspace

Do not stop at "the command failed"; continue through fix and re-verification unless the user redirects.

## Reporting Contract

Always report:

- the exact workspace path or paths used
- whether blank-workbook verification ran
- whether macro round-trip verification ran
- whether `init` verification ran
- whether release-gate coverage was complete, partial, or intentionally narrowed
- concrete evidence:
  - command results
  - generated artifact checks
  - Excel COM workbook-state checks when relevant
- any untested path that remains out of scope
