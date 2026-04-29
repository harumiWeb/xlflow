---
name: xlflow-tmp-workspace-e2e
description: Create disposable `xlflow` projects under this repository's `tmp_workspaces` directory and verify the Windows Excel COM flow end-to-end. Use when Codex needs to validate `xlflow` behavior with real workbooks, run the standard `new/init/doctor/pull/push/run/lint` workflow, confirm workbook state after macro execution, or reproduce/fix CLI integration bugs without touching the main project tree.
---

# xlflow tmp_workspaces E2E

Execute the repository-standard `xlflow` verification flow in isolated workspaces under `tmp_workspaces`.

Follow the workflow exactly unless the user narrows scope.

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
xlflow push --json
xlflow run Main.Run --json
xlflow pull --json
xlflow lint --json
```

Verify:

- `push` reports imported source files and workbook-module updates when expected
- `run` reports the requested macro name
- `pull` exports `Main.bas` plus workbook modules
- the final `lint` still passes

### 4. Workbook-state verification

When macro behavior should persist to the workbook, verify it through Excel COM instead of trusting CLI success output alone.

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
- concrete evidence:
  - command results
  - generated artifact checks
  - Excel COM workbook-state checks when relevant
- any untested path that remains out of scope
