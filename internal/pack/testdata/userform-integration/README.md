# Issue #887 local Excel integration evidence

Verified on 2026-10-03 with Excel 16.0 build 17932, Windows x64,
trusted VBIDE access, and the Go and .NET binaries installed by
`rtk proxy task install`. Production pack remains pure Go and reports
`vbe_validation = "not_performed"`; these checks are developer-only.

## Canonical template and blank forms

Command from the repository root:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-pack-userforms-e2e.ps1 -WorkspacePath tmp_workspaces/issue-887-template-20261003-r6
```

Workspace:
`C:\Users\HARUMI\orca\workspaces\xlflow\acornworm\tmp_workspaces\issue-887-template-20261003-r6`.
`run-evidence.json` records success, five artifact checks, and confirmed cleanup
of every owned Excel instance without forced termination. The harness retains
all source inputs, pack envelopes, expected observations, and Excel readbacks.

Verified blank generation with two forms and omission of one; template addition
and Caption/Text/Left edits; preservation of an omitted form and omitted
properties; explicit source topology with one form and with no forms; retained
reference identities and no broken references; macro sentinels; save/reopen;
and canonical file-pull topology for all five artifacts. Unsupported Frame
generation returned `pack_userform_generation_unsupported` and retained the
existing output's SHA256.

The template creator sets both the Designer Caption and the component Caption
property. Setting only Designer Caption does not establish the runtime Caption
of an Excel-authored fixture, so the fixture must not assert that equivalence.
Pure-Go generation and explicit template edits synchronize FormControl and
VBFrame. An explicit equal Caption also repairs a legacy VBFrame mismatch.

## Existing workflows and Unicode boundary

These commands also passed:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-pack-e2e.ps1 -KeepWorkspace -WorkspaceSuffix '-issue-887-20261003-r2'
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-forms-reference-e2e.ps1 -WorkspacePath tmp_workspaces/issue-887-blank-20261003-r1
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-file-pull-e2e.ps1 -KeepWorkspace -WorkspaceSuffix '-issue-887-20261003-r1'
```

Retained workspaces under
`C:\Users\HARUMI\orca\workspaces\xlflow\acornworm\tmp_workspaces`:

- `pack-stable-e2e-issue-887-20261003-r2`
- `pack-stable-blank-e2e-issue-887-20261003-r2`
- `issue-887-blank-20261003-r1`
- `file-pull-release-e2e-issue-887-20261003-r1`
- `file-pull-userform-e2e-issue-887-20261003-r1`

They cover legacy nested Designer/code preservation, blank compilation and
runtime execution, standard/class/document round-trips, UserForm extraction,
compatibility `.frx` preservation, and the session-backed push/save workflow.

In the fresh workspace
`C:\Users\HARUMI\orca\workspaces\xlflow\acornworm\tmp_workspaces\issue-887-init-20261003-r1`,
`rtk xlflow init Book.xlsm --json` and `rtk xlflow doctor --json` passed using a
copy of the successful legacy pack artifact.

Additional focused CP932 generation and runtime checks passed:

```powershell
rtk powershell -NoProfile -ExecutionPolicy Bypass -File scripts/dev/go.ps1 run tmp_workspaces/issue-887-caption-20261003-r1/create.go
rtk powershell -NoProfile -ExecutionPolicy Bypass -File tmp_workspaces/issue-887-caption-20261003-r1/verify.ps1
```

Workspace:
`C:\Users\HARUMI\orca\workspaces\xlflow\acornworm\tmp_workspaces\issue-887-caption-20261003-r1`.
`evidence.json` records runtime Caption `日本語 "引用"`, TextBox Value `値😀`,
and successful save/reopen. Root Caption must be representable in the project
code page; control Unicode persistence remains independent of it. Rejected
root captions, quoting, opaque property-block preservation, and no-op bytes are
covered by Go regressions.

No remote CI or Linux execution was performed during this local verification.
The Windows-only Excel scripts are excluded from ordinary tests and CI.
