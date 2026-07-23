# Inspect workbook content and formulas

Use read-only inspection to turn “I think the macro worked” into evidence. Inspection does not edit VBA or formulas; it reports what is already in the selected workbook.

```bash
xlflow inspect workbook --json
xlflow inspect sheets --json
xlflow inspect range --sheet Result --address A1:F20 --include-style --json
xlflow formulas pull --json
xlflow export-image --sheet Result --range A1:F20 --json
```

`inspect workbook` gives the broad shape; `inspect sheets` helps choose a sheet; `inspect range` proves an exact value, formula, or style; `export-image` is useful when layout matters. Add `--session` when the state exists only in the live workbook. Save first when formulas were changed in a dirty session and the command reads the disk workbook.
