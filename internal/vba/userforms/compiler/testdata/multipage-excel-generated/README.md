# MultiPage generation and editing Excel gate (Issue #885)

The captured `.bin`, observations and environment records are repository-authored
and were published after the capture harness confirmed Excel process exit.
`Main.bas` contains the workbook-qualified RunMultiSentinel used by these
trusted local artifacts. The harness accepts an expected sentinel, checks its
cell value before save and after reopening the saved workbook, and records it
in JSON. It must only be used with trusted generated workbooks.

- `cli-template`: reorder Pages and standalone Tabs, add a Page/Tab, resize
  MultiPage, change the Japanese Page caption and flags, disable both owners,
  retain nested controls and their values, and preserve selection by identity.
- `cli-empty-blank`: new generation with zero Pages/Tabs and logical -1.
- `cli-empty-template`: remove all Pages, descendant controls and Tabs.
- `cli-bridge`: Excel-backed build/snapshot/run/save in an explicit session.
  This initial capture also exposed dropped Page property-bag metadata; it
  verifies structure and flags, not metadata parity.
- `cli-bridge-corrected`: repeats the session-backed build/snapshot/runtime
  workflow after the metadata alias fix. The snapshot and saved/reopened
  Designer both preserve PageAlpha Tag, ControlTipText and Accelerator.

The initial nonempty blank pack also passed RunMultiSentinel and save/reopen
in `issue-885-cli-blank-02`; its retained observations are in that workspace.

Executed commands from the CLI workspace were:

```powershell
rtk xlflow pack --blank --out dist/Blank2.xlsm --json
rtk xlflow pack --template ../issue-885-authored-09/baseline.xlsm --out dist/Template3.xlsm --json
rtk xlflow pack --blank --out dist/EmptyBlank.xlsm --json
rtk xlflow pack --template dist/Template3.xlsm --out dist/EmptyTemplate.xlsm --json
rtk xlflow session start --json
rtk xlflow form build src/forms/specs/MultiTopologyForm.yaml --overwrite --session --json
rtk xlflow form snapshot MultiTopologyForm --out bridge-snapshot.json --session --json
rtk xlflow run Main.RunMultiSentinel --session --no-save --json --timeout 30s
rtk xlflow save --session --json
rtk xlflow session stop --json
```

Overwrite build requires its existing intermediate-save path and rejects
`--no-save` before mutation. The remaining workbook-backed commands reused the
same session.

Each artifact was then opened by the developer-only
`scripts/test-userform-multipage-e2e.ps1 -WorkbookPath <artifact> -WorkspacePath
<fresh workspace> -ExpectedSentinel <recorded runtimeSentinel>` from the root.
The complete expected sentinel is retained in each artifact's JSON.

Workspace root:
`C:\Users\HARUMI\orca\workspaces\xlflow\huchen\tmp_workspaces`.
Successful workspaces: `issue-885-cli-blank-02`, `issue-885-cli-template-02`,
`issue-885-cli-template-05`, `issue-885-cli-empty-blank-01`,
`issue-885-cli-empty-template-01`, `issue-885-cli-bridge-01`, and
`issue-885-cli-bridge-02`.
Failed runtime/capture attempts are separate and supplied no promoted evidence.

These gates cover MultiPage/Page/TabStrip and nested common controls. They do
not establish support for root outer-dimension conversion, nonstandard tab
layout mutation, pictures, list persistence, or the unrelated full release gate.
