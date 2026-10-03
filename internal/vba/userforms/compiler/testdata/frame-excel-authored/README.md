# Frame persistence evidence (Issue #884)

These are repository-authored fixtures, captured with Windows Excel 16.0,
build 17932 and trusted VBIDE access. Ordinary tests only read the binary files;
they never start Excel. The Excel harness refuses existing Excel processes,
identifies its own process through `Application.Hwnd`, and publishes evidence
only after that process exits. Existing evidence is never overwritten.

`baseline.xlsm` and `baseline.bin` contain two common controls at the root, an
empty Frame, and a ParentFrame with SiblingText and NestedFrame/NestedLabel.
`baseline.json` records Designer properties and hierarchy before save and after
reopen. The baseline inspection did **not** run its macro; runtime confirmation
belongs to the generated fixtures below.

`added.bin`, `removed.bin`, and `reordered.bin` are consecutive Excel saves:
add DiffLabel to NestedFrame, remove that label, then call SiblingText.ZOrder(0).
They bind descendant membership cookies, ID high-water marks, root TypeInfoVer,
and persisted sibling order. The differential capture confirms Excel cleanup
but makes no independent runtime claim. Reproduce it with the developer-only
`scripts/test-userform-frame-counters-e2e.ps1` using a fresh workspace.

The sibling directory `frame-excel-generated` contains `cli-blank` and
`cli-template` binary, canonical expected JSON, COM observations and environment
records. Both artifacts passed workbook-qualified RunFrameSentinel, then save,
reopen and the same runtime check. The template variant additionally replaces
EmptyFrame with a Label, removes NestedLabel, moves RootText into NestedFrame,
reorders root and ParentFrame siblings, and adds AddedFrame/AddedLabel. Saved
binary readback, rather than COM collection enumeration, verifies sibling order.

Create a new baseline with:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-userform-frame-e2e.ps1 -Phase create -WorkspacePath tmp_workspaces/issue-884-new-baseline
```

This requires the authored fixture directory to be absent and does not replace
the committed files. For standalone pure-Go compiler artifacts, set absolute
`XLFLOW_FRAME_WORKBOOK`, `XLFLOW_FRAME_OUTPUT`, and `XLFLOW_FRAME_EXPECTED` paths
and run `TestGenerateExcelFrameArtifact`. Set `XLFLOW_FRAME_MODE=template` for the
structural variant; omit it for new generation. All output paths must be new.
Execute the resulting workbook with the frame harness `-Phase verify`, its
canonical `-ExpectedPath`, a fresh `-WorkspacePath`, and `-ObserveOnly` to retain
new measurements without publishing fixtures.

The CLI checks actually executed in this worktree were:

```powershell
# cwd: tmp_workspaces/issue-884-cli-pack
rtk xlflow pack --blank --out dist/BlankClassTableFixed.xlsm --json
rtk xlflow pack --template ../issue-884-frame-e2e-20261003-053216-84e382/baseline.xlsm --out dist/TemplateTypeInfoFixed.xlsm --json

# cwd: repository root
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-userform-frame-e2e.ps1 -Phase verify -Variant cli-blank -WorkbookPath tmp_workspaces/issue-884-cli-pack/dist/BlankClassTableFixed.xlsm -ExpectedPath tmp_workspaces/issue-884-frame-e2e-20261003-053216-84e382/generated-expected.json -WorkspacePath tmp_workspaces/issue-884-cli-blank-class-table-fixed-3
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-userform-frame-e2e.ps1 -Phase verify -Variant cli-template -WorkbookPath tmp_workspaces/issue-884-cli-pack/dist/TemplateTypeInfoFixed.xlsm -ExpectedPath tmp_workspaces/issue-884-cli-pack/edited-type-info-fixed.json -WorkspacePath tmp_workspaces/issue-884-cli-template-type-info-fixed
```

Measured workspace root:
`C:\Users\HARUMI\orca\workspaces\xlflow\crinoid\tmp_workspaces`.
Successful retained workspaces: `issue-884-baseline-confirmed-2`,
`issue-884-counter-evidence`, `issue-884-cli-blank-class-table-fixed-3`,
`issue-884-type-info-fixed-2`, and `issue-884-cli-template-type-info-fixed`.
Failed attempts remain separate and supplied no generated fixture evidence.

This gate covers Frame/common-control structure and runtime persistence. It
does not cover MultiPage/Page/TabStrip mutation, root outer-dimension changes,
picture/list persistence, or the unrelated full release gate.

Local checks passed after the fixes:

```powershell
rtk task install
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/dev/go.ps1 test ./internal/vba/userforms/... ./internal/pack/... ./internal/filepull/... -count=1
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/dev/go.ps1 test ./internal/cli -run 'TestPackBlank.*Form|TestPack.*Form.*Authority' -count=1
rtk proxy golangci-lint run ./internal/vba/userforms/... ./internal/pack/... ./internal/filepull/...
rtk pnpm docs:check
```

The new PowerShell harnesses also passed PSScriptAnalyzer. A focused independent
static review found no confirmed bug; it did not independently execute Excel.
The final `pull --backend file --json` in the CLI workspace extracted four VBA
components and preserved the edited hierarchy and persisted sibling order.
