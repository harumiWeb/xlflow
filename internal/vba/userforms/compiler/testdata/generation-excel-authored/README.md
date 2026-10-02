# Excel-authored generation evidence (Issue #883)

This repository-created fixture was captured with Excel 16.0 build 17932 on
64-bit Windows, with trusted VBIDE access enabled. `environment.json` records
the owned Excel process and confirmed cleanup. No third-party workbook is used.

`baseline.xlsm` contains the eleven supported built-in classes and an empty
form. `baseline.bin` is its unchanged `xl/vbaProject.bin`. The four
`enabled-*-False.bin` files were extracted from separately saved/reopened
Excel workbooks after disabling that control. Their companion observations
bind the saved values to actual Designer state. ToggleButton stores
`VariousPropertyBits=0x2c800819`; SpinButton, ScrollBar and Image store `0x19`.
The focused Go test protects these bits, class parsing and CompObj identity.

The original observation script's optional dynamic COM property reads can
produce null for captions and other properties. Those nulls are not positive
evidence of missing persisted values. The current harness uses explicit COM
dispatch. Image does not expose a public TabIndex; normalized binary readback
checks its site TabIndex instead.

The persisted `DisplayedSize` height and exported VBFrame `ClientHeight`
match each other, while runtime `InsideHeight` is about 4.55 points smaller
in this environment. The generation contract therefore uses explicit
`form.build.clientWidth/clientHeight`. It does not infer outer dimensions or
apply a fixed offset.

Capture command (must use a fresh workspace and fixture directory):

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-userform-generation-e2e.ps1 -Phase create -WorkspacePath tmp_workspaces/issue-883-observe-20261002-final -FixtureDirectory internal/vba/userforms/compiler/testdata/generation-excel-authored
```

Actual capture workspace:
`C:\Users\HARUMI\orca\workspaces\xlflow\dolphin\tmp_workspaces\issue-883-observe-20261002-final`.

Generation and runtime verification use opt-in tests and the local harness;
ordinary tests and CI do not launch Excel. Use absolute environment paths
because Go tests execute in the package directory, and create the new output
directory before running the artifact helper:

```powershell
$env:XLFLOW_GENERATION_WORKBOOK = '<absolute fixture path>/baseline.xlsm'
$env:XLFLOW_GENERATION_OUTPUT = '<absolute new workspace>/generated.xlsm'
$env:XLFLOW_GENERATION_EXPECTED = '<absolute new workspace>/expected.json'
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/dev/go.ps1 test ./internal/vba/userforms/compiler -run TestGenerateExcelNewFormArtifact -count=1 -v
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-userform-generation-e2e.ps1 -Phase verify -WorkbookPath '<generated.xlsm>' -ExpectedPath '<expected.json>' -NormalizedWorkbookPath '<new normalized.xlsm>'
$env:XLFLOW_GENERATION_NORMALIZED_PROJECT = '<normalized.xlsm.bin>'
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/dev/go.ps1 test ./internal/vba/userforms/compiler -run TestReadExcelNormalizedNewForms -count=1 -v
```

The artifact helper retains both existing Designers and the existing Forms
reference, then adds independently generated forms. It verifies Japanese and
emoji captions/text, Boolean values, numeric positions, geometry, tab order,
disabled controls, code-behind invocation and the `issue-883-ok` cell sentinel.
Saving/reopening and Go binary readback check normalization. This does not
exercise pack CLI integration, blank-project reference creation or containers;
those remain subsequent issues #886/#887 and later generation stages.

Verified generated workspace:
`C:\Users\HARUMI\orca\workspaces\xlflow\dolphin\tmp_workspaces\issue-883-generate-20261002`.
The final successful run used `grid.xlsm`, `grid-expected.json`, and
`grid-normalized.xlsm` (client size 324 x 282 points; runtime InsideHeight
277.45 points). The verify command was:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-userform-generation-e2e.ps1 -Phase verify -WorkbookPath tmp_workspaces/issue-883-generate-20261002/grid.xlsm -ExpectedPath tmp_workspaces/issue-883-generate-20261002/grid-expected.json -NormalizedWorkbookPath tmp_workspaces/issue-883-generate-20261002/grid-normalized.xlsm
```

Result: expected Designer fields and client export dimensions matched;
`issue-883-ok` sentinel succeeded; saved/reopened values matched; owned Excel
PID 214748 exited. The subsequent `TestReadExcelNormalizedNewForms` passed,
including exact persisted client sizes, geometry and all eleven site tab
indices. The unchanged normalized binary and expected projection are retained
in sibling directory `generation-excel-generated` and checked by ordinary
pure-Go tests. VBIDE exports quantize some non-grid client values (for example
320 points to 319.8); the runtime test uses whole export-grid dimensions while
pure-Go generation tests cover arbitrary HIMETRIC rounding.

Earlier isolated attempts are retained for inspection:

- `C:\Users\HARUMI\orca\workspaces\xlflow\dolphin\tmp_workspaces\issue-883-observe-20261002`
- `C:\Users\HARUMI\orca\workspaces\xlflow\dolphin\tmp_workspaces\issue-883-observe-20261002-retry1`
- `C:\Users\HARUMI\orca\workspaces\xlflow\dolphin\tmp_workspaces\issue-883-observe-20261002-retry2`

They exposed harness file-sharing,
assembly-loading and VBE identifier-case problems and are not successful gate
evidence. Cleanup was confirmed after each attempt. No blank-project pack CLI,
MSForms-reference mutation or full release gate is claimed by this evidence.

Local verification commands:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/dev/go.ps1 test ./... -count=1
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/dev/go.ps1 test ./internal/excel ./internal/vba/userforms/... ./internal/pack/... -count=1
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/dev/go.ps1 test ./internal/cli -run 'UserForm|FormSpec|Pack' -count=1
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/dev/go.ps1 test ./internal/lspserver -run UserForm -count=1
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/dev/go.ps1 vet ./internal/vba/userforms/... ./internal/pack/... ./internal/excel
rtk pnpm docs:check
rtk git diff --check
```

The full run passed the other packages, including static analysis, but caught
three package failures during the field-name transition (`excel`, `spec`,
`spec/intel`). Those were corrected and their complete package tests passed
on rerun, along with all UserForm/pack tests. CLI/LSP focused tests, vet,
documentation checks, scoped formatting checks and PowerShell syntax/analyzer
error checks passed. Remote CI has not been run for these uncommitted changes.
