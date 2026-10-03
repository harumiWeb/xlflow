# MultiPage and TabStrip persistence evidence (Issue #885)

Repository-authored fixtures were saved and reopened sequentially with Windows
Excel 16.0, build 17932, with trusted VBIDE access. The capture harness refuses
existing Excel processes, identifies its owned process through Application.Hwnd,
and exports binary evidence only after confirmed process exit. Ordinary tests
read saved bytes and never launch Excel.

`baseline.xlsm` / `baseline.bin` contain MultiMain with PageAlpha/PageBeta,
PageText and PageFrame/NestedLabel under PageAlpha, and an independent StripMain.
The differential saves are `disabled` (Japanese caption, hidden/disabled Page,
disabled owners), `added`, `resized`, `reordered`, `removed`, and `empty`.
Each JSON records Designer properties before save and after reopen. Page COM
geometry is unavailable in this environment and is not an authoring input.

The x stream's ID binds the unnamed internal TabStrip, and page IDs establish
page order. Page Tag lives in the Site, while caption, tooltip, accelerator and
page enabled/visible flags live in the hidden TabStrip. MultiPage Enabled is
the inverse of x Properties mask bit 3. Empty saved MultiPages can retain both
a stale Value and cached hidden tab arrays: x owns the empty Page collection,
public selection is -1, and no-op serialization retains the stored bytes.

TabsAllocated is capacity rather than current count. A changed collection
requires a complete TabData flag array; actual Excel ignores incomplete flag
arrays after growth. This difference is covered by codec regression tests and
the generated template runtime fixture.

Capture a fresh set with:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-userform-multipage-e2e.ps1 -WorkspacePath tmp_workspaces/issue-885-new-authored
```

The retained successful capture is
`C:\Users\HARUMI\orca\workspaces\xlflow\huchen\tmp_workspaces\issue-885-authored-09`.
Failed earlier captures supplied no promoted evidence. Numeric Pages.Remove
and assigning Value=-1 to an empty MultiPage crashed Excel in failed probes;
the supported bridge removes Pages by exact name and skips empty Value setters.

The sibling `multipage-excel-generated` directory records pure-Go pack and
Excel bridge runtime, save/reopen and cleanup evidence. These local checks do
not change pack's permanent `vbe_validation: not_performed` production contract.
