# xlflow fix

Preview or apply the supported Quick Fix corrections for explicit VBA source
files, using one diagnostic rule. The command is source-only and never opens
Excel.

## Usage

```powershell
xlflow fix <path> [<path>...] --rule VB044|VB080|VB084|VB086 [--apply] [--json]
```

Without `--apply`, the command previews the edits. Add `--apply` to write
changed results for the selected file or files. Each invocation captures and
analyzes its current inputs, so a prior preview is not an apply plan.

## Options and Arguments

| Option / argument | Description                                                                     | Default      |
| ----------------- | ------------------------------------------------------------------------------- | ------------ |
| `<path>...`       | Explicit existing files under eligible configured source or test roots.         | Required     |
| `--rule`          | One rule: `VB044`, `VB080`, `VB084`, or `VB086`.                                | Required     |
| `--apply`         | Apply changed bytes; multiple files use staged rollback and recovery reporting. | Preview only |
| `--json`          | Return the structured xlflow JSON envelope.                                     | false        |

The project root must contain `xlflow.toml`. Eligible files are under the
configured modules, classes, or workbook source roots, or the project test
root. The initial command accepts only `.bas` and `.cls` files. Under the
configured forms root, only `forms/code/*.bas` is eligible, and only when
`[userform].code_source = "sidecar"`. Forms assets, specs, and inactive
sidecar paths are rejected even when their extensions look like source. `.frm`
files are excluded because their `.frx` sidecars need a separate transaction
contract. Source must be valid UTF-8 without a BOM.

## What the command does

`fix` reuses the shared same-rule correction provider, its current rule
configuration, inline suppressions, diagnostic associations, and UTF-16 edit
application. It does not enable rules or change diagnostic defaults. The
command reads the saved file on disk and does not use open editor buffers,
workbooks, Excel COM, or a workbook session. It does not load TypeDB, run
whole-project lint, or construct ProcedureIR/CFG.

Preview output reports the proposed edits and unified diff without writing.
For apply, the command captures and computes every selected file before it
publishes any source. It validates every target, including files with no edits;
a read-only file is rejected only when its proposed bytes change. It stages
replacement bytes and original-byte backups beside every changed target,
preserves permission bits, and rechecks target identity, parent, mode, and
contents after staging and again before each publication. Symlink and reparse
targets and non-regular files are rejected. No-op files create no temporary
files.

If a publication fails or cancellation occurs before all changed files publish,
the command restores already-published files in reverse order and ignores
cancellation during recovery. Before restoring, it verifies the published
target's identity, mode, and bytes and verifies the original-byte backup. It
does not knowingly overwrite an external edit. If a conflict or restore failure
prevents rollback, original recovery material is retained and its paths are
reported. If the parent directory moved or was replaced, or the source-side
backup was altered or became inaccessible, the command treats that backup path
as untrusted. It attempts to persist an independent copy of the captured
original bytes in an OS temporary file created with permission mode `0600`
(subject to host filesystem permission semantics), using the
`.sourceedit-*.tmp` file and reports its absolute path as `recovery_path`
outside the project. An inaccessible or unremoved source-side artifact is
reported separately from the fallback copy. If all persistence attempts fail,
recovery is not guaranteed; the persistence errors and affected files are
reported. Same-directory `.sourceedit-*.tmp` artifacts are removed after a
successful commit or rollback.

With one positional path, changed bytes retain the existing single-file
publication behavior: the command stages beside the target and publishes with
the platform's native atomic replacement after rechecking identity, mode, and
contents. A mismatch returns `fix_source_changed`.

The transaction is best-effort. It does not provide simultaneous atomic
visibility across files, crash durability, or a race-free compare-and-swap
against external writers. Process termination does not trigger automatic
recovery.

The command advertises a source-tree mutation capability to consumers through
policy `ResourceSourceTree`, operation `mutate`, and `ParallelSafe: false`.
This metadata acquires no source-tree lock or workbook lease and provides no
`--wait` behavior. The target recheck remains best-effort against external
writers.

## JSON Output

Successful JSON uses the standard envelope with `command="fix"` and an
`output` object. `mode` is `preview` or `apply`; `path` is project-relative
with `/` separators; `rule` is uppercase. `changed` reports whether proposed
bytes differ from the captured source. `applied` is true only when changed
bytes were written, so previews and no-change applies report false.

`source_sha256` and `result_sha256` cover the captured and proposed bytes.
They identify the current invocation's file state and do not make a previous
preview reusable. `edits` contain ranges against the captured source revision
using zero-based UTF-16 line and character positions, plus `new_text`.
`diagnostics` contains the associated diagnostic metadata from the shared
provider. `diff` is the unified source-to-result diff and is empty when no
change is proposed.

When two or more paths are provided, `output` contains `mode` (`preview` or
`apply`), `rule`, `files` (the complete single-file results in that mode), and
`summary.file_count`, `summary.changed_file_count`, and `summary.edit_count`.
Unchanged files are included and report `applied=false`. Logical path aliases
are deduplicated (case-folded only on Windows), and results are sorted by
project-relative path. Every argument is resolved and validated before logical
deduplication. Batch apply rejects distinct selected paths with the same physical identity, such as
hard links, are rejected before publication. Multiple arguments use the batch
schema even if aliases select one logical file. `target.paths` lists the
selected paths in the same order. One positional path keeps the existing
single-file output shape.

On success, each changed file that was published reports `applied=true`; no-op
files report `false`. The summary counts unique logical files, including
unchanged, suppressed, and disabled-rule results. A failure returns one error
and omits `output`. Directories, globs, project discovery, cross-file edits,
and `.frm` transactions are unsupported. Captures are sequential and do not
form an atomic project snapshot or a reusable apply plan.

Errors from the batch transaction include `error.details.committed`, `error.details.files`, and
`error.details.artifacts`. File entries report `path` and one of the states
`unchanged`, `not_applied`, `applied`, `restored`, or `recovery_required`, with
an optional `recovery_path`. `unchanged` files needed no edits, `not_applied`
files were not published, `applied` files remain committed, `restored` files
were returned to their original bytes, and `recovery_required` files could not
be safely restored. `committed` is true only when all requested changes were
published. Retained recovery material is listed under `artifacts`, and
human-readable errors also show transaction states and recovery paths. CLI paths
are project-relative with `/` separators, except paths outside the project,
which remain absolute.

## Exit Codes

| Code | Meaning                                                                                             |
| ---: | --------------------------------------------------------------------------------------------------- |
|  `0` | Preview or apply succeeded, including when no correction applies.                                   |
|  `1` | Invalid source encoding, source changed before publication, or invalid edit set.                    |
|  `2` | Invalid arguments or missing/invalid project configuration.                                         |
|  `3` | File-system or operational failure, cancellation, rollback failure, or post-commit cleanup failure. |

Structured errors include `fix_args_invalid` (exit `2`), the existing
`source_encoding_invalid` projection (exit `1`), `fix_source_changed` and
`fix_edits_invalid` (exit `1`), `fix_failed` and `fix_cancelled` (exit `3`),
`fix_rollback_failed` (exit `3`), and `fix_cleanup_failed` (exit `3`). Rollback
failure overrides the triggering stale-state or cancellation error. A
post-commit cleanup failure reports `committed=true` with the committed file
states. A missing `xlflow.toml` uses the standard `config_error` projection
(exit `2`).

## Examples

Preview a rule's applicable corrections:

```powershell
xlflow fix src/modules/Loops.bas --rule VB084 --json
```

Preview the same rule in multiple explicit files:

```powershell
xlflow fix src/modules/First.bas src/classes/Second.cls --rule VB080 --json
```

Apply the current result after reviewing the preview:

```powershell
xlflow fix src/modules/Loops.bas --rule VB084 --apply --json
```

Apply the reviewed rule to several explicit files as one recoverable batch:

```powershell
xlflow fix src/modules/First.bas src/classes/Second.cls --rule VB080 --apply --json
```

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow fix` to preview and selectively apply one supported rule's safe
corrections. Preview and apply can select several explicit configured source
files; multi-file application reports transaction state and recovery paths on
failure.

## Prerequisites

Run it from a project containing `xlflow.toml`. Each selected file must be in
the configured modules, classes, workbook, or test roots and use `.bas` or
`.cls`; the only forms path allowed is active sidecar `forms/code/*.bas`.

## What this command reads and changes

The command reads explicitly selected saved source files and project
configuration. With `--apply`, it can replace changed files after staging
original-byte recovery copies and rechecking captured file state. It does not
read or change a workbook or an unsaved editor buffer.

## Effect on source-of-truth state

Preview leaves all source files untouched. Apply writes only when the proposed
bytes differ. Review the unified diff before using `--apply`; a later apply
recomputes against the file as it exists then.

## Common workflows

Run `xlflow fix <path> --rule <VB-rule> --json`, inspect `output.diff`, then
repeat with `--apply` when the current result is intended. For multiple paths,
inspect `output.files[].diff` and `output.summary`, then apply the same explicit
file list with `--apply`. A later invocation recomputes every correction.

## Common failures

Check `error.code`, the exit code, and `error.details.files`. Correct project
configuration, source encoding, file scope, or concurrent file changes as
indicated, then rerun the command to compute a fresh result. For
`fix_rollback_failed`, use the reported recovery paths and inspect affected
files before restoring anything manually. For `fix_cleanup_failed`, the source
changes are committed; `error.details.committed=true` and the artifact paths
identify cleanup that remains.
