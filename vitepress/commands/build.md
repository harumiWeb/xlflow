# xlflow build

Resolve the release build plan for an Excel-backed artifact. The configured base workbook is input only and is never modified.

```bash
xlflow build --dry-run
xlflow build --base templates/Production.xlsm --out dist/Product.xlsm --dry-run
```

`--base` defaults to `[excel].path`. `--out` defaults to `build/Release/<base-workbook-name>` and always accepts a complete workbook path, not a directory. Base and output must be different project-local files with the same `.xlsm`, `.xlam`, or `.xlsb` extension.

## Dry-run

Use `--dry-run` to validate the workbook paths and configured `[build].exclude` patterns without opening Excel, acquiring a workbook lock, creating a directory, or writing an artifact. It also remains local when invoked from WSL. The output lists included and excluded VBA components; `--json` exposes the same information under `build`.

The Excel/VBIDE reconstruction and publication pipeline is not available yet. A validated invocation without `--dry-run` returns `build_not_implemented`. Use `pack` when an Excel-independent artifact workflow is needed.

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow build --dry-run` to review exactly which source components will be included in a future release artifact. It does not replace `push`, which continues to synchronize the full source tree into the development workbook.

## Common failures

Read the structured `error.code` instead of scraping terminal text. Missing bases, unsupported or mismatched extensions, equal base/output paths, and invalid source plans fail before Excel can open.
