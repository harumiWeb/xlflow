# xlflow build

Build a release artifact from the configured VBA source. The configured base
workbook is input only and is never modified.

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

## Dry-run

Use `--dry-run` to validate workbook paths and configured `[build].exclude`
patterns without opening Excel, acquiring workbook coordination, creating a
directory, or writing an artifact. It also remains local when invoked from WSL.
The output lists included and excluded VBA components; `--json` exposes the
same information under `build`.

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
