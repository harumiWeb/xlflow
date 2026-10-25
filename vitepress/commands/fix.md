# xlflow fix

Preview or apply the supported Quick Fix corrections for one VBA source file
and one diagnostic rule. The command is source-only and never opens Excel.

## Usage

```powershell
xlflow fix <path> --rule VB044|VB080|VB084|VB086 [--apply] [--json]
```

Without `--apply`, the command previews the edits. Add `--apply` to write a
changed result. The command always captures and analyzes the current file for
each invocation, so a prior preview is not an apply plan.

## Options and Arguments

| Option / argument | Description                                                         | Default      |
| ----------------- | ------------------------------------------------------------------- | ------------ |
| `<path>`          | One existing file under an eligible configured source or test root. | Required     |
| `--rule`          | One rule: `VB044`, `VB080`, `VB084`, or `VB086`.                    | Required     |
| `--apply`         | Apply changed bytes to the file.                                    | Preview only |
| `--json`          | Return the structured xlflow JSON envelope.                         | false        |

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
Apply stages changed bytes beside the target and publishes them with the
platform's native atomic replace. Symlink targets and non-regular files are
rejected. A read-only target is rejected in apply mode when the result changes
the file. Before publication, the command rechecks file identity, mode, and
contents. This detects many concurrent changes, but does not provide a
race-free compare-and-swap against arbitrary external writers or crash
durability. Cancellation or failure before publication cleans up the staging
file.

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

## Exit Codes

| Code | Meaning                                                                          |
| ---: | -------------------------------------------------------------------------------- |
|  `0` | Preview or apply succeeded, including when no correction applies.                |
|  `1` | Invalid source encoding, source changed before publication, or invalid edit set. |
|  `2` | Invalid arguments or missing/invalid project configuration.                      |
|  `3` | File-system or operational failure, including cancellation.                      |

Structured errors include `fix_args_invalid` (exit `2`), the existing
`source_encoding_invalid` projection (exit `1`), `fix_source_changed` and
`fix_edits_invalid` (exit `1`), `fix_failed` (exit `3`), and
`fix_cancelled` (exit `3`). A missing `xlflow.toml` uses the standard
`config_error` projection (exit `2`).

## Examples

Preview a rule's applicable corrections:

```powershell
xlflow fix src/modules/Loops.bas --rule VB084 --json
```

Apply the current result after reviewing the preview:

```powershell
xlflow fix src/modules/Loops.bas --rule VB084 --apply --json
```

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow fix` to preview and selectively apply one supported rule's safe
corrections in one configured source file.

## Prerequisites

Run it from a project containing `xlflow.toml`. The selected file must be in
the configured modules, classes, workbook, or test roots and use `.bas` or
`.cls`; the only forms path allowed is active sidecar `forms/code/*.bas`.

## What this command reads and changes

The command reads one saved source file and its project configuration. With
`--apply`, it can replace that file after rechecking the captured file state.
It does not read or change a workbook or an unsaved editor buffer.

## Effect on source-of-truth state

Preview leaves the source file untouched. Apply writes only when the proposed
bytes differ. Review the unified diff before using `--apply`; a later apply
recomputes against the file as it exists then.

## Common workflows

Run `xlflow fix <path> --rule <VB-rule> --json`, inspect `output.diff`, then
repeat with `--apply` when the current result is intended.

## Common failures

Check `error.code` and the exit code. Correct project configuration, source
encoding, file scope, or concurrent file changes as indicated, then rerun the
command to compute a fresh result.
