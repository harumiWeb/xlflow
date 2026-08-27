# xlflow build

Build a release artifact from the configured VBA source. The configured base
workbook is input only and is never modified. Use `build` for a Windows/Excel
release artifact: Excel/VBIDE reconstructs the project and validates VBE
compilation before xlflow publishes the file.

```bash
xlflow build --dry-run
xlflow build --base templates/Production.xlsm --out dist/Product.xlsm --dry-run
```

`--base` defaults to `[excel].path`. `--out` defaults to `build/Release/<base-workbook-name>` and always accepts a complete workbook path, not a directory. Base and output must be different project-local files with the same `.xlsm`, `.xlam`, or `.xlsb` extension.

## Publication

A non-dry build creates a bridge-owned, uniquely named staging directory beside
the requested output. Excel reconstructs and compiles only that staged workbook.
xlflow saves it, closes the workbook, and confirms that Excel exited before
checking that the staged artifact exists, is non-empty, and is readable.

Only then is the artifact published. A new output is moved into place on the
same volume (`atomic_create`); an existing output is replaced through the
operating system's atomic replacement API (`atomic_replace`). xlflow never uses
a delete-then-copy fallback. Therefore, a failed build leaves an existing output
unchanged.

Before starting Excel, xlflow creates and validates the output parent directory
and coordinates both the base and output workbooks. It refuses to build when the
output is locked by Office, owned by a live xlflow session, or cannot safely be
replaced. A matching dirty xlflow session for the base must be saved first.

Successful `--json` results include:

```json
{
  "output": {
    "path": "C:\\project\\dist\\Product.xlsm",
    "replaced_existing": false,
    "publication": "atomic_create",
    "temporary_cleanup": {
      "status": "clean"
    }
  }
}
```

`publication` is `atomic_create` or `atomic_replace`; `replaced_existing`
matches that publication. `temporary_cleanup.status` is `clean` or `failed`; a
failed cleanup can also include `residual_path` and `error`. Text output reports
the same publication method, whether an existing artifact was replaced, and the
staging-cleanup result.

Staging cleanup occurs after publication. If cleanup alone fails, the published
artifact remains successful and xlflow returns warning
`build_temporary_cleanup_failed` (with the retained staging path when
available); `output.temporary_cleanup.status` is `failed`.

## Build manifest

After successful workbook publication, a non-dry build attempts to publish a
versioned companion manifest at `<output>.build.json` (for example
`build/Release/Book.xlsm.build.json`). The same v1 fields appear under `build`
in `--json`: project-relative base/output paths, included and excluded
components, VBE/save/close validation, and publication metadata.
`build.schema_version` is the contract version for integrations.

The workbook artifact remains authoritative. If only the manifest cannot be
published after the workbook was safely published, build succeeds with warning
`build_manifest_publish_failed` and `build.manifest.published=false`; rerun the
build to recreate the companion file.

## Choose the right command

| Command | Purpose                                                                                              | Excel/VBE validation                                             |
| ------- | ---------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------- |
| `build` | Publish a filtered release workbook without changing the development workbook.                       | Required on Windows.                                             |
| `push`  | Synchronize the complete source tree, including development/test code, into the configured workbook. | Runs through Excel/VBIDE.                                        |
| `pack`  | Produce an experimental `.xlsm` with the pure-Go file-level writer.                                  | Never performed by `pack`; release-gate Excel smoke is required. |

`[build].exclude` filters only `build`; it never changes `push` or `pack`.
Projects created by `xlflow new` or `xlflow init` exclude the scaffold's
`src/modules/Tests/**` and `src/modules/Xlflow/XlflowAssert.bas` by default;
edit or clear the list when those components are required in the release.
Supported base/output pairs are `.xlsm`, `.xlam`, and `.xlsb` with matching
extensions. Excel and Trust Center VBA-project access are required for a real
build. Save a matching dirty xlflow session before building.

`build` guarantees a separately staged workbook, VBE compilation, a clean
save/close, and publication only after those checks pass. It does not prove
that every runtime macro path is correct; run representative workbook tests or
smoke macros before distributing a release.

## Dry-run

Use `--dry-run` to validate workbook paths and configured `[build].exclude`
patterns without opening Excel, acquiring workbook coordination, creating a
directory, or writing an artifact. It also remains local when invoked from WSL.
The output lists included and excluded VBA components; `--json` exposes the
same information under `build`.

`build --dry-run` is the explicit Excel-free exception: it returns the v1
manifest shape with `validation.vbe_compile="not_run"`, creates neither an
artifact nor a companion manifest, and remains local under WSL.

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow build --dry-run` to review exactly which source components will be
included in a release artifact. Use `xlflow build` to create the validated,
published artifact. Neither replaces `push`, which synchronizes the full source
tree into the development workbook.

## Common failures

Read the structured `error.code` instead of scraping terminal text. Missing
bases, unsupported or mismatched extensions, equal base/output paths, and
invalid source plans fail before Excel can open. Publication-specific failures
are `build_output_directory_failed`, `build_output_busy`,
`build_temporary_artifact_missing`, and `build_output_replace_failed`; all
preserve a prior output artifact.
