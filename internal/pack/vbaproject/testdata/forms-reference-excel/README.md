# First-UserForm reference evidence (Issue #886)

Repository-authored fixtures captured on Windows with Excel 16.0 and trusted
VBIDE access. `environment.json` records the exact build, owned process and
confirmed cleanup. No third-party fixture is used.

`form-free.bin`, `first-form.bin`, and `two-forms.bin` contain saved/reopened
projects before a UserForm, after the first form, and after the second form.
Their JSON files retain reference GUID/version/broken-state observations and
Designer state. Existing Forms CONTROL groups are preserved; generated blank
projects use a REGISTERED Forms 2.0 reference instead of a user-specific `.exd`.

Capture command:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-userform-generation-e2e.ps1 -Phase create -WorkspacePath tmp_workspaces/issue-886-reference-observe-20261003-r1 -FixtureDirectory internal/pack/vbaproject/testdata/forms-reference-excel
```

Workspace:
`C:\Users\HARUMI\orca\workspaces\xlflow\acornworm\tmp_workspaces\issue-886-reference-observe-20261003-r1`.
Owned Excel PID 428604 exited with cleanup confirmed. The initial workspace
without a VBA module produced no VBA project and was rejected; it is not
successful evidence.

For future captures, add `-ReferenceFixturesOnly` to publish only the three
reference-stage binaries/observations and environment, omitting unrelated
control-property fixtures. The retained directory contains this focused subset
and the separately captured `blank-generated.bin/json` from the blank gate.

The separate blank CLI gate passed:

```powershell
rtk task install
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-forms-reference-e2e.ps1 -WorkspacePath tmp_workspaces/issue-886-blank-20261003-r1
```

Workspace:
`C:\Users\HARUMI\orca\workspaces\xlflow\acornworm\tmp_workspaces\issue-886-blank-20261003-r1`.
It generated two forms (empty plus eleven supported controls) through actual
`pack --blank`, checked exactly one unbroken Forms reference, compiled an
explicit `MSForms.TextBox` declaration, ran code-behind and sentinel
`issue-883-ok`, saved/reopened, compared Designer observations, and confirmed
owned Excel PID 341352 exited. The reused sentinel/harness is from Issue #883;
this run adds first-reference and actual blank CLI evidence.
`normalized.xlsm.bin.json` retains both reference snapshots and cleanup.
The first artifact setup missed a standard module attribute and was rejected
before Excel; its workspace is retained as failed setup evidence.

Additional release regressions passed with the installed main/bridge binaries:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-pack-e2e.ps1 -KeepWorkspace -WorkspaceSuffix '-issue-886-20261003'
```

- `C:\Users\HARUMI\orca\workspaces\xlflow\acornworm\tmp_workspaces\pack-stable-e2e-issue-886-20261003`: template standard/class edits, nested Designer/reference preservation, `pack stable ok`.
- `C:\Users\HARUMI\orca\workspaces\xlflow\acornworm\tmp_workspaces\pack-stable-blank-e2e-issue-886-20261003`: two blank artifacts, both `pack blank ok`, baseline references unbroken and project-reference rename control.
- `C:\Users\HARUMI\orca\workspaces\xlflow\acornworm\tmp_workspaces\issue-886-roundtrip-20261003`: `new --no-update-check`, `doctor`, `pull`, `lint` passed; standard GateMain, GateClass and GateForm round-trip passed with `.frm/.frx` exports and sentinel `issue-886-roundtrip-ok`.
- `C:\Users\HARUMI\orca\workspaces\xlflow\acornworm\tmp_workspaces\issue-886-init-20261003`: `init <absolute roundtrip workbook> --no-update-check --userform-code-source sidecar`, `pull --backend excel`, and saved-cell inspect passed; class and `.frm/.frx` artifacts present.

Round-trip command sequence inside the primary workspace:

```powershell
rtk xlflow --json session start
rtk xlflow --json form build src/forms/specs/GateForm.yaml --session --no-save
rtk xlflow --json push --fast --session --no-save
rtk xlflow --json run GateMain.RunGate --session
rtk xlflow --json inspect cell --sheet Sheet1 --address A1 --session
rtk xlflow --json save --session
rtk xlflow --json pull --backend excel --session
rtk xlflow --json session stop
```

The initial push before form build correctly failed FRM201 before mutation.
The successful sequence above generated compatibility artifacts first.
No real signed-project certificate fixture or non-Windows Excel host was tested.

These gates run locally only. Ordinary Go tests and Linux CI use committed
binary evidence without starting Excel. Template Designer creation, nested
generation and custom ActiveX remain outside Issue #886.

Local validation passed:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/dev/go.ps1 test ./... -timeout 15m
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/dev/go.ps1 test ./internal/pack/... ./internal/sourceinventory ./internal/cli -count=1
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/dev/go.ps1 test -race ./internal/pack/ovba ./internal/pack/vbaproject ./internal/sourceinventory -count=1
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/dev/go.ps1 test ./internal/pack/ovba -run '^$' -fuzz '^FuzzProjectReferences$' -fuzztime 5s
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/dev/go.ps1 vet ./internal/pack/... ./internal/sourceinventory ./internal/cli
rtk task lint
rtk pnpm docs:check
rtk git diff --check
```

The full suite passed during implementation; final inventory validation and
fixture-test additions were followed by the complete affected package suites,
focused Japanese/control generation tests and vet. Source gofmt/goimports and
changed Markdown/JSON oxfmt checks passed. No remote CI result is claimed.

Two independent, read-only scoped agent reviews completed without confirmed
findings. The reference reviewer independently passed ovba/vbaproject/pack
tests and a five-second fuzz smoke (~2.83 million executions). The blank-input
reviewer inspected source authority, fallback, validation and publication;
it did not rerun tests or Excel. Case-variant `specs`/`code` directory spelling
was recorded as an unconfirmed follow-up rather than a specification defect.
Live Excel evidence was executed by the implementation agent.
