# Saved-Workbook File Push

## Command

```text
xlflow push [--backend excel|file] [--backup always|never] [--changed-only]
            [--fast] [--json]
```

`excel` is the default backend and keeps the existing bridge/VBIDE contract:
source is imported into the live (or freshly opened) project, the document
modules are rewritten in place, and the workbook is saved by Excel.

`--backend file` selects the pure-Go saved-workbook backend. Explicit backend
requests never fall back. There is no `auto` for push. `--backend file`
combined with `--session` or `--no-save` is rejected with `push_args_invalid`.

## Authority and host requirements

The file backend reads the configured source tree and the saved `.xlsm`
package, rebuilds `xl/vbaProject.bin` with the pure-Go pack engine, and
atomically replaces the workbook on disk. It never launches or attaches to
Excel, COM, VBIDE, or the bridge, so it runs on Linux, containers, CI, WSL,
and remote agents without a desktop session. Under WSL it stays local except
for the live-session safety probe described below.

A successful file push confirms only that the artifact was rebuilt and
structurally validated. VBA is not compiled and no VBE/runtime validation is
performed; output states `vbe_validation="not_performed"` and emits warning
`vbe_validation_skipped`. Consumers needing compile evidence must use the
Excel backend or validate the pushed workbook separately.

Successful JSON includes:

```json
{
  "push": {
    "backend": "file",
    "target": "saved_workbook",
    "vbe_validation": "not_performed"
  },
  "workbook": { "path": "build/Book.xlsm", "saved": true },
  "target": { "kind": "file", "path": "build/Book.xlsm" },
  "session": { "active": false, "mode": "none", "source_of_truth": "saved_workbook" },
  "source": {
    "changed": true,
    "changed_only": false,
    "files": 4,
    "state": ".xlflow/state/push.json"
  },
  "output": {
    "publication": "atomic_replace",
    "replaced_existing": true,
    "temporary_cleanup": { "status": "clean" }
  },
  "warnings": [
    {
      "code": "vbe_validation_skipped",
      "message": "file push did not open Excel; no VBE compile or runtime validation was performed."
    }
  ]
}
```

A changed-only skip reports `source.changed=false`, emits no `output`, and
creates no backup.

## Safety gate

File push overwrites the saved workbook without consulting any live Excel/VBE
state, so conditions that imply an open or session-owned workbook are hard
failures — unlike file pull, which can read past them with a warning. All
checks run after source preflight and before any mutation; rejection leaves
the workbook byte-identical.

- An Office lock file (`~$<name>`) beside the configured workbook →
  `push_workbook_open`.
- A `.xlflow/session.json` record matching the configured workbook →
  `push_active_session`, regardless of whether the recorded PID is still
  alive (same conservative contract as `pack`).
- On Windows, a Restart Manager report of the workbook open in Excel, or an
  indeterminate open-state probe → `push_workbook_open`.
- Under WSL, a matching live session on the Windows side, or a failed
  live-session probe → `push_active_session`.

## Source pipeline

Source coverage, transforms, and fingerprinting deliberately mirror the Excel
bridge so both backends agree on what changed:

- Discovery covers `modules/**/*.bas`, `classes/**/*.cls`,
  `forms/**/*.frm|.frx|.bas|.cls`, `workbook/**/*.bas`, and — only in
  sidecar mode — `forms/code/**/*.bas` as `form_code`. Fingerprint entries are
  `(kind, root-relative path, SHA-256)`.
- In sidecar mode, canonical `forms/specs/**/*.yaml|.yml|.json` FormSpecs are
  passed to the shared pure-Go UserForm compiler. A FormSpec can add a form,
  update a supported Designer, or replace/remove an Image picture. Referenced
  assets are resolved relative to the project root and included in the
  file-backend fingerprint. See
  [`userform-picture-assets.md`](userform-picture-assets.md) for schema,
  bounds, and containment rules.
- Only image assets referenced by canonical FormSpecs are loaded, validated,
  and included in file-push fingerprints. Unreferenced files under the reserved
  forms `assets/` directory are retained and ignored; pull does not
  automatically garbage-collect them.
- Duplicate component names across module/class/form/document entries fail
  with `duplicate_module_name` before mutation. `.frx` companions and
  `form_code` sidecars never collide.
- Document-module disk sources are normalized (class headers and `Attribute
VB_*` lines stripped; an empty body becomes `Option Explicit`), matching
  `NormalizeDocumentModuleContent`. A `.cls` file under the workbook root is
  ignored, as the Excel backend only applies `<name>.bas` documents.
- `[vba.folder_annotation]` mode `update` inserts or replaces
  `'@Folder("A.B")'` after the `VERSION`/`Attribute` header block; `ignore`
  and `preserve` leave text untouched. Folder segments derive from the file's
  directory relative to its configured root.
- `[vba.line_numbers]` instrumentation ports `ErlLineNumberTransformer.TryAdd`
  and runs as a whole-tree preflight: existing numeric labels or numeric
  `GoTo`/`GoSub`/`Resume` targets fail with `vba_line_number_safety_failed`
  before the workbook is read.
- Canonical FormSpec state is the Designer authority for sidecar forms. A
  retained `.frm` / `.frx` is not imported when a canonical spec and sidecar
  code are selected, so the file-pull
  `compatibility_artifact_unsynchronized` marker does not block that path.
  The marker still fails with `FRM201` when `frm` code or compatibility
  artifacts are selected as import authority. File push never rewrites
  compatibility files in the source tree.
- File-push UserForm topology remains template-authoritative: forms absent
  from source are preserved, and `[pack].userform_topology` is ignored. In
  sidecar mode, a missing code sidecar preserves an existing form's code; an
  explicitly empty sidecar clears it. A newly added form with no code sidecar
  starts with empty code. `pack` retains its separate topology and code-source
  contract.
- New forms, Designer edits, and supported picture edits are compiled and
  validated by the shared pure-Go pipeline before publication. Existing
  unsupported picture resources are preserved when omitted; an operation that
  would lose unsupported resource bytes fails before publication. Custom
  ActiveX generation or import is unsupported by this backend and never
  triggers an Excel fallback. Such compatibility operations require explicit
  Excel-authored `.frm` / `.frx` artifacts and an explicitly selected Excel
  workflow. The current registry probes found no `MSComctlLib.TreeCtrl.2`,
  `MSComctlLib.ListViewCtrl.2`, `RefEdit.Ctrl`, or `MSCAL.Calendar.7`; no
  real-Excel custom ActiveX import behavior has been verified.

Protected or signed projects, missing `xl/vbaProject.bin`, and structurally
ambiguous component sets fail before publication with the pack engine's
deterministic codes (`push_protected_project`, `push_signed_project`,
`push_ambiguous_layout`).

## Publication contract

The rebuilt workbook is fully staged and structurally validated (readable zip
containing a non-empty `xl/vbaProject.bin`) before the target is touched.
Publication is a single atomic create-or-replace of the configured workbook
through `coordination.PublishFile`; a returned failure leaves the previous
workbook in place and cleans up the sibling temporary artifact. With
`--backup always` (default) a backup is recorded before the replace and its
metadata carries `backend="file"`; `--backup never` skips it. A backup failure
aborts the push with `push_backup_failed` and the workbook unchanged.

After a successful replace, `.xlflow/state/push.json` is written atomically
with the bridge-compatible shape: `{fingerprint, applied_to:{saved_file}}`.
A state-write failure is reported as warning `push_state_persist_failed` and
does not roll back the already-published workbook.

## Changed-only skip

`--changed-only` skips the rebuild when both halves of the recorded state
still hold:

1. the source fingerprint (workbook path, per-file kind/path/hash entries,
   `line_numbers_enabled`, and the effective `folder_annotation` mode)
   matches the current tree — compared order-insensitively. File-backend
   fingerprints additionally include canonical FormSpecs and referenced
   asset bytes; a fingerprint from a backend that did not include those inputs
   cannot justify a skip;
2. `applied_to.saved_file` still describes the workbook: normalized path,
   last-write timestamp in .NET ticks, and byte length.

Workbook paths in `fingerprint.workbook_path` and `applied_to.saved_file.path`
are canonicalized so the same file is identical across the WSL/Windows
boundary: under WSL an absolute `/mnt/<drive>/...` path is recorded as its
Windows `D:\...` form, and comparisons normalize both sides the same way, so
a state file written by the file backend on WSL satisfies the Excel bridge's
`--changed-only` check on Windows (and vice versa) when both recorded
fingerprints represent the same effective inputs. Paths outside `/mnt/` keep
their native form — Windows cannot see them, so no interop is claimed.

A state file in the legacy bare-fingerprint shape has no delivery evidence and
never justifies a skip. A skip is a successful no-op.

## Coordination

The file backend acquires the same source-tree leases as file pull: each
managed module/class/forms/workbook root exclusively, every canonical ancestor
as shared intent, in stable LockID order. Contended roots fail with
`source_tree_busy` or wait under `--wait`/`--wait-timeout`.

The workbook lease is not held for the command's lifetime — a changed-only
skip, the safety gate, source collection, and the rebuild never touch it. Only
the mutation window takes it: immediately before the template read, the
backend acquires the same workbook identity lease other mutating commands use
and holds it through the atomic replace and the `push.json` write. The
acquisition is non-blocking — the source-tree leases are already held, and
waiting on the workbook lease there could deadlock against a command that
holds the workbook lease and waits on the source tree — so contention fails
fast with `workbook_busy` even under `--wait`.

Under the lease the backend applies the same recovery gate as coordinated
workbook commands: a workbook carrying a recovery marker fails with
`workbook_recovery_required`, and an unreadable marker fails closed with
`coordination_recovery_check_failed`. This closes the window where a lock
file and session record have cleared but a timed-out Excel operation still
requires explicit recovery.

## Cross-backend parity corpus

`testdata/push-parity/cases.json` is consumed by both the Go file-backend
tests and the .NET bridge tests. It fixes the shared contract for folder
annotation insert/replace/remove and mode handling, Erl instrumentation of
eligible procedure lines (including continuation tails, block `If`,
declarations, labels, and unsafe numeric targets), and document-module
normalization. Both sides compare LF-normalized text because the bridge joins
lines with the host newline.

## Edge cases stricter than the Excel backend

The file backend rebuilds the component topology from the source tree, so a few
inputs that the VBIDE path tolerates are rejected or applied differently:

- A document-module source (`workbook/**/*.bas`) with no matching document
  component in the workbook is a `push_ambiguous_layout` error. The Excel
  backend silently skips such orphan sources because it iterates workbook
  components, not disk files.
- A whitespace-only `code/<Form>.bas` sidecar is authoritative: the merged
  form ends up with empty code-behind. The Excel bridge treats a
  whitespace-only sidecar as absent and preserves the embedded code.
- `Attribute VB_Name` must match the expected component name exactly. The
  Excel backend's VBIDE import accepts case-only mismatches
  (`Module1.bas` declaring `VB_Name = "module1"`); the file backend rejects
  them as `push_ambiguous_layout`.
- When the configured workbook does not exist, Windows surfaces
  `push_workbook_open` (the open-state probe cannot determine the state)
  while other platforms surface `push_source_read_failed` when the template
  is read.

All of these fail or diverge on the safe side; none can silently corrupt a
workbook.

## Real-Excel release gate

After `task install`, run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-file-push-e2e.ps1 -KeepWorkspace
```

The gate file-pushes a module tree into a saved `.xlsm`, then opens the result
in real Excel, compiles, and executes a sentinel macro to prove the rebuilt
project is live-valid — the boundary the file backend itself cannot check.
The script prints the workspace path, sentinel result, and Excel/OS identity
for release evidence.

`scripts/test-formspec-file-push-e2e.ps1` covers canonical pull → file push →
pull, spec-only and asset-only changes, malformed-asset rejection without
publication, blank and template pack, and an Excel open/compile/run check after
the source assets are unavailable. Picture-specific Designer persistence is
also covered by `scripts/test-userform-pictures-e2e.ps1`. The Excel-authored
JPEG fixture is normalized to BMP; the separate canonical Issue #912 gate
verifies native JPEG picture output through file push, blank pack, and template
pack in real Excel after source asset paths are unavailable. All three paths
passed picture display, Designer property, sentinel, and SaveAs/reopen checks
on Excel 16.0 build 17932 / Windows 10.0.22631, with owned Excel cleanup
confirmed. These are
developer-only gates and do not run in PR CI. The available registry did not
contain `MSComctlLib.TreeCtrl.2`, `MSComctlLib.ListViewCtrl.2`,
`RefEdit.Ctrl`, or `MSCAL.Calendar.7`; no real-Excel custom ActiveX import
result is established by these gates.

## Stable failures

| Error code                             | Meaning                                                                                         |
| -------------------------------------- | ----------------------------------------------------------------------------------------------- |
| `push_args_invalid`                    | Unknown backend, or `--backend file` with `--session`/`--no-save`.                              |
| `workbook_format_unsupported`          | File backend was requested for a non-`.xlsm` workbook.                                          |
| `push_workbook_open`                   | Lock file, Restart Manager open report, or indeterminate probe.                                 |
| `push_active_session`                  | A matching xlflow session is recorded (or WSL probe failed).                                    |
| `duplicate_module_name`                | Case-insensitive component-name collision across kinds.                                         |
| `vba_line_number_safety_failed`        | Source is unsafe for Erl instrumentation.                                                       |
| `push_backup_failed`                   | Pre-push backup could not be created.                                                           |
| `push_protected_project`               | The saved VBA project is protected.                                                             |
| `push_signed_project`                  | The saved VBA project is signed.                                                                |
| `push_userform_generation_unsupported` | A requested UserForm generation or Designer resource is outside the supported pure-Go contract. |
| `push_ambiguous_layout`                | Component set cannot be mapped unambiguously onto the project.                                  |
| `push_output_busy`                     | The workbook cannot be replaced because it is in use.                                           |
| `push_output_replace_failed`           | Atomic replacement of the workbook failed.                                                      |
| `push_write_failed`                    | Publication failed for an uncategorized reason.                                                 |
| `push_state_persist_failed` (warning)  | Workbook published but `push.json` could not be written.                                        |
| `workbook_busy`                        | Another xlflow operation holds the workbook lease for the push window.                          |
| `workbook_recovery_required`           | The workbook carries a recovery marker requiring explicit recovery.                             |
| `coordination_recovery_check_failed`   | The workbook's recovery state could not be read safely.                                         |
| `coordination_acquire_failed`          | The publish-window workbook lease could not be acquired.                                        |
| `source_tree_busy`                     | Another process owns a managed source root.                                                     |
| `source_tree_busy_timeout`             | Waiting for a managed source root timed out.                                                    |
| `source_tree_busy_cancelled`           | Waiting for a managed source root was cancelled.                                                |
