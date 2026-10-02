# Issue 882: Excel-authored persistence evidence

These fixtures were created from an empty workbook by the repository's local
`scripts/test-userform-mutation-e2e.ps1` harness on 2026-10-02. Excel 16.0,
build 17932, Windows (64-bit) NT 10.00, trusted VBIDE access. No external code,
downloaded workbook, or xlflow CLI binary was used. Normal tests consume the
committed bytes and never start Excel.

## Artifacts and reproduction

`baseline.xlsm` is the original workbook; `00-baseline.bin` is its unmodified
`xl/vbaProject.bin`. The other `.bin` files are the same ZIP entry extracted
from each named Excel save. Companion `.json` records contain live state
before saving and Designer state after closing/reopening. `.binary.json`
records contain pure-Go `oforms.ReadForm(..., "MutationForm", 932)` inspection;
absent values are explicitly marked, not interpreted as zero defaults.

The form is `MutationForm`; root names are `LabelMain`, `TextMain`,
`ButtonMain`, `CheckMain`, `OptionMain`, `ComboMain`, `ListMain`, `FrameMain`.
`NestedText` is a TextBox inside `FrameMain`. Named standard module
`Main.RunMutationSentinel` creates and unloads `MutationForm` and writes
`issue-882-ok` to worksheet 1 cell A1. A runtime failure writes its error
number/description to A1 instead of opening a VBA error dialog.

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-userform-mutation-e2e.ps1 -Phase create -FixtureDirectory internal/vba/userforms/compiler/testdata/new-excel-evidence
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-userform-mutation-e2e.ps1 -Phase verify -WorkbookPath internal/vba/userforms/compiler/testdata/excel-authored/baseline.xlsm -ExpectedPath internal/vba/userforms/compiler/testdata/excel-authored/00-baseline.json
```

Create requires new workspace and fixture directories. All temporary workbooks
are retained. Verify opens read-only, compares explicit canonical FormSpec
fields and `properties`, ignores observational/unsupported metadata, and runs
the workbook-qualified sentinel. Geometry tolerance is 0.05 points for Excel
twip versus binary HIMETRIC rounding; string and flag comparisons are exact.
Form width/height are not compared because outer/client dimensions differ.

Verify explicitly executes workbook VBA with the developer's Excel authority.
Use only this trusted generated baseline and its known compiler-produced
derivatives. Read-only opening and disabled events do not sandbox the sentinel,
and matching Designer properties does not authenticate workbook code. Do not
verify an untrusted workbook in a credential-bearing developer environment.

Main's opt-in `TestGenerateExcelMutationArtifact` accepts
`XLFLOW_MUTATION_WORKBOOK`, `XLFLOW_MUTATION_OUTPUT`, and
`XLFLOW_MUTATION_EXPECTED`; use this baseline as input and pass its output
workbook and expected FormSpec JSON to the verify phase.

## List and selectedIndex persistence

Both controls were unbound (`RowSource=""`), single-column, single-selection.
`AddItem` inserted `alpha`, `beta`, `gamma` on the Designer control. After
save/close/reopen, both had `ListCount=0` and `ListIndex=-1`. Adding the same
items and then assigning `ListIndex=1` yielded live `Value="beta"`. After
save/close/reopen, items and selectedIndex again disappeared. ComboBox
`Value="beta"` persisted; ListBox Value returned to null. The selection stage
re-adds items if the preceding reopen discarded them. Each stage is saved and
reopened before the next mutation. The JSON captures these boundaries.

This is evidence for rejecting list/selectedIndex edits in the binary compiler,
not evidence that RowSource-backed lists or runtime initialization are unsupported
by Excel. Those paths were not investigated.

## Enabled and site bits

Each control was individually disabled, saved/reopened, enabled, saved/reopened.
`04-all-disabled.bin` additionally stores every control disabled simultaneously.
Every individual Enabled transition survived reopening.

| Kind                              | File-format VariousPropertyBits when omitted | Excel baseline field       | Disabled field             |
| --------------------------------- | -------------------------------------------- | -------------------------- | -------------------------- |
| Label                             | `0x0080001b`                                 | absent                     | `0x00800019`               |
| CommandButton                     | `0x0000001b`                                 | absent                     | `0x00000019`               |
| CheckBox / OptionButton / ListBox | `0x2c80081b`                                 | absent                     | `0x2c800819`               |
| TextBox / ComboBox                | `0x2c80081b`                                 | `0x2c80481b` present       | `0x2c804819`               |
| Frame                             | does not use this field                      | BooleanProperties=`0x8004` | BooleanProperties=`0x8000` |

For embedded controls Enabled is bit `0x2`. For Frame Enabled is BooleanProperties
bit `0x4`. Preserve every other bit. TextBox/ComboBox's Excel-authored explicit
baseline is not the file-format omission default. The omission defaults above
are specified in [MS-OFORMS VariousPropertyBits](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/7a72ac4a-39d9-4e2b-829e-19e3e9a1f60d)
and independently match the disabled fields for the controls where Excel omits
the baseline field. This observation did not produce a TextBox/ComboBox fixture
with that field omitted.

Site `BitFlags` is explicit `0x32` on Label and `0x40023` on Frame. It is absent
on the other seven controls (file-format default `0x33`). Site visibility uses
bit `0x2`; site defaults and VariousPropertyBits defaults are separate contracts.

## Local verification record

Final retained workspace:
`C:\Users\HARUMI\orca\workspaces\xlflow\shipworm\tmp_workspaces\issue-882-observe-20261002-112443-bb14c5`.
Create completed; verify of the committed baseline against `00-baseline.json`
passed and sentinel was `issue-882-ok`. Create PID 392472 and verify PID 75652
both exited through normal Close/Quit; no concurrent Excel session was present.

The compiler-produced artifact and an Excel-normalized SaveAs/reopen copy were
verified under
`C:\Users\HARUMI\orca\workspaces\xlflow\shipworm\tmp_workspaces\issue-882-compile-20261002`.
The final gate compared root/control captions, UTF-16 Japanese/emoji text,
checked values, parent-relative geometry, tab-order edits, Enabled/Visible, colors, border
properties, MaxLength, GroupName, Tag, and ControlTipText. The sentinel passed,
the normalized project passed pure-Go read-back, and owned Excel PID 318660
exited cleanly. Excel rounds exposed geometry to twips; the gate allows
0.05 points while retaining exact string/flag comparisons.

```powershell
$env:XLFLOW_MUTATION_WORKBOOK = 'C:/Users/HARUMI/orca/workspaces/xlflow/shipworm/internal/vba/userforms/compiler/testdata/excel-authored/baseline.xlsm'
$env:XLFLOW_MUTATION_OUTPUT = 'C:/Users/HARUMI/orca/workspaces/xlflow/shipworm/tmp_workspaces/issue-882-compile-20261002/edited.xlsm'
$env:XLFLOW_MUTATION_EXPECTED = 'C:/Users/HARUMI/orca/workspaces/xlflow/shipworm/tmp_workspaces/issue-882-compile-20261002/expected.json'
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/dev/go.ps1 test ./internal/vba/userforms/compiler -run TestGenerateExcelMutationArtifact -count=1
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-userform-mutation-e2e.ps1 -Phase verify -WorkbookPath tmp_workspaces/issue-882-compile-20261002/edited.xlsm -ExpectedPath tmp_workspaces/issue-882-compile-20261002/expected.json -NormalizedWorkbookPath tmp_workspaces/issue-882-compile-20261002/normalized-final.xlsm
$env:XLFLOW_MUTATION_NORMALIZED_PROJECT = 'C:/Users/HARUMI/orca/workspaces/xlflow/shipworm/tmp_workspaces/issue-882-compile-20261002/normalized-final.xlsm.bin'
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/dev/go.ps1 test ./internal/vba/userforms/compiler -run TestReadExcelNormalizedArtifact -count=1
```

Use fresh output paths when reproducing the normalization gate. Compressed
MS-OFORMS strings are low-byte UTF-16, not project-code-page MBCS. The
Japanese caption gate detects that distinction; VBFrame retains its separate
MBCS encoding contract.

This is a focused persistence/mutation gate. CLI scaffold/new/init/push/pull and
the complete release matrix were intentionally outside this investigation.

## PR #902 review-fix verification

The coordinate-system, caption-baseline and property-key collection fixes were
verified again with the same trusted baseline and fresh outputs under
`C:\Users\HARUMI\orca\workspaces\xlflow\shipworm\tmp_workspaces\issue-882-review-fix-20261002`.
`TestGenerateExcelMutationArtifact` produced `edited.xlsm` and `expected.json`;
the verify phase used those paths and
`-NormalizedWorkbookPath tmp_workspaces/issue-882-review-fix-20261002/normalized.xlsm`.
All nine controls matched before and after SaveAs/reopen, A1 was
`issue-882-ok`, Excel remained version 16.0 build 17932, and owned PID 374348
exited cleanly. `TestReadExcelNormalizedArtifact` passed against
`normalized.xlsm.bin`. This recheck retains the same narrowed Designer-edit
scope rather than claiming the complete CLI release matrix.
