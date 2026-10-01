# Saved-Workbook File Pull

## Command

```text
xlflow pull [--backend auto|file|excel] [--formulas] [--json]
```

`auto` is the default backend. Explicit `file` and `excel` selections never
invoke the other backend. `--backend file --session` is invalid.

On Windows, auto selects Excel for explicit session intent, a valid matching
xlflow session, a workbook reported open by Restart Manager, an indeterminate
open-state probe, UserForms, and Excel-supported project formats outside the
file backend. A closed, supported `.xlsm` selects file. Missing, unreadable,
malformed, protected, and unsafe projects fail with their existing validation
codes instead of falling through to Excel. On non-Windows hosts auto uses file
when supported and otherwise returns the deterministic file-backend error.

Stable selection reasons are `explicit_backend`, `session_requested`,
`matching_live_session`, `workbook_open_in_excel`, `open_state_probe_failed`,
`file_backend_supported`, `file_backend_unsupported_userform`, and
`file_backend_unsupported_format`. An open-state failure selects Excel and adds
warning `pull_auto_open_state_probe_failed` with the original failure detail.
Explicit selections report `backend_selection=explicit`; omitted or explicit
`--backend auto` reports `backend_selection=auto`.

## Authority and host requirements

The file backend reads the configured saved `.xlsm` package. It never launches
or attaches to Excel, COM, VBIDE, or the bridge. A matching
`.xlflow/session.json` record produces warning
`file_pull_live_session_ignored`; unsaved live workbook changes are not read.
Unreadable session metadata produces `file_pull_session_state_unavailable`
without changing saved-file authority.

Under WSL, explicit file and the normal auto pure-Go path remain local.
Explicit Excel, `--session`, and a validated matching live session delegate to
Windows xlflow.

Successful JSON includes:

```json
{
  "pull": {
    "backend": "file",
    "backend_selection": "auto",
    "selection_reason": "file_backend_supported",
    "source": "saved_workbook",
    "code_page": 932,
    "modules": {
      "standard": 1,
      "class": 1,
      "document": 2,
      "form": 0
    }
  }
}
```

## Supported project surface

The initial backend accepts `.xlsm` only and requires exactly one
`xl/vbaProject.bin`. Standard modules become `.bas`, class modules become
importable `.cls`, and document modules become body-only `.bas` source under
their configured roots. Source is decoded using the VBA project's declared
code page and written as UTF-8 without BOM with a final newline. Empty document
modules use the same body-only representation as the Excel backend:
`Option Explicit` followed by a final newline.

Configured Rubberduck `@Folder` behavior and `[vba.line_numbers]` removal apply
equally to file and Excel pulls. Unsafe names, path escapes, duplicate target
paths, malformed projects, and unsafe line-number transformations fail before
publication.

UserForms are not partially supported. Detection of any form rejects the whole
operation with `pull_userform_unsupported` before source mutation. The forms
root is otherwise unmanaged and is never deleted or rewritten.
The configured forms root must not contain, or be contained by, the module,
class, or workbook root; overlap fails before reconciliation so form artifacts
cannot be mistaken for stale managed source.

## Publication contract

Planning and validation complete before publication. Publication reconciles
managed `.bas` and `.cls` files under the module, class, and workbook roots.
Unrelated files and the forms root are retained. If writing or stale-file
removal fails, xlflow restores every affected prior file and reports
`pull_source_publish_failed`.

The guarantees are deliberately distinct:

- each individual replacement is atomic;
- a returned publication failure attempts operation-level rollback of every
  affected managed file, including prior contents and supported file modes,
  and removes newly created files/directories where possible;
- abrupt process termination, OS crash, or power loss is not a crash-safe
  whole-tree transaction and automatic recovery is not currently provided.

Do not describe this contract as an unqualified atomic source-tree update.

## Source-tree coordination

The file backend derives one `source_tree` resource identity for each managed
module, class, and document-module root. Identities canonicalize the nearest
existing ancestor and use a source-tree-specific hash domain. Each managed root
is acquired exclusively and every canonical ancestor is acquired as a shared
intent; duplicate identities are promoted to exclusive mode and the complete
set is acquired in stable LockID order. This makes ancestor/descendant roots
contend while disjoint siblings remain concurrent. The leases cover workbook
parsing, planning, publication, rollback, and cleanup. The forms root is not
acquired because this backend rejects UserForms and never mutates form artifacts.

Source-tree locks use crash-released, descriptor-owned operating-system file
locks with shared and exclusive modes on Windows and Unix. The explicit file
backend does not acquire a workbook lock;
therefore it works without Excel and remains local under WSL. Two pulls sharing
a managed root fail with `source_tree_busy` by default or wait under the normal
`--wait`/`--wait-timeout` contract. The timeout starts only after the first
authoritative contention result; uncontended lock setup does not consume the
wait budget. Disjoint managed roots do not contend.

If a future command requires both resource kinds, it must acquire all workbook
identities first, then all source-tree identities, sorting each group by LockID
and releasing in reverse order.

## Cross-backend parity corpus

`testdata/pull-parity/cases.json` is consumed by both the Go file-backend tests
and the .NET Excel-backend tests. It fixes the shared tracked-source contract
for standard/class/document normalization, empty documents, folder annotations,
generated Erl removal, mixed line endings, Unicode, and terminal newlines.
Backend-specific inputs may differ, but their normalized tracked output must
match the same expected fixture.

## Real-Excel release gate

After `task install`, run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-file-pull-e2e.ps1 -KeepWorkspace
```

The gate maintains the Excel pull baseline → file pull → normalized comparison
→ pack → real Excel execution chain. Its fixture includes multiple standard,
class, and document modules, nested folders, CP932/Japanese text, a non-ASCII
component name, cross-module execution, and ordinary project references. A
separate UserForm-bearing copy must fail without changing the tracked source.
The script prints both absolute workspace paths, the sentinel, and the Excel/OS
identity needed for release evidence.

## Stable failures

| Error code                      | Meaning                                                |
| ------------------------------- | ------------------------------------------------------ |
| `pull_args_invalid`             | Unknown backend or incompatible `--session`.           |
| `workbook_format_unsupported`   | File backend was requested for a non-`.xlsm` workbook. |
| `pull_file_not_found`           | The configured workbook does not exist.                |
| `pull_file_unreadable`          | The saved workbook cannot be read.                     |
| `pull_vba_project_missing`      | The package has no `xl/vbaProject.bin`.                |
| `pull_vba_project_malformed`    | The package or VBA project is malformed.               |
| `pull_protected_project`        | The saved VBA project is protected.                    |
| `pull_userform_unsupported`     | At least one UserForm is present.                      |
| `pull_source_path_unsafe`       | A component cannot be mapped inside its managed root.  |
| `vba_line_number_safety_failed` | Generated line numbers cannot be removed safely.       |
| `pull_source_publish_failed`    | Transactional source publication failed.               |
| `source_tree_busy`              | Another process owns a managed source root.            |
| `source_tree_busy_timeout`      | Waiting for a managed source root timed out.           |
| `source_tree_busy_cancelled`    | Waiting for a managed source root was cancelled.       |

`--formulas` runs only after successful VBA publication and retains its
existing saved-workbook formula authority and partial-failure contract.
