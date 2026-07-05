# formulas

Extract saved workbook formulas into source-controlled snapshots without opening Excel.

## Pull Formula Snapshots

```bash
xlflow formulas pull --json
```

The command reads the workbook configured by `[excel].path` and writes:

```text
formulas/
  manifest.json
  names.jsonl
  sheets/
    001-Sheet1.regions.jsonl
```

`manifest.json` lists workbook sheets and region files. `names.jsonl` contains workbook and sheet scoped defined names. Each sheet JSONL file contains logical formula regions grouped by normalized R1C1-like formula patterns. OOXML shared formula group boundaries are coalesced when they differ only by storage metadata.

`formulas pull` supports `.xlsx` and `.xlsm`, reads OOXML files directly, and does not use Excel COM. Cached calculated values are not included.

Unsupported formula syntax is preserved as raw formula data with `partial` or `failed` parse status so one unsupported formula does not block the snapshot.

## Standalone Use

Use `--src` to point at a workbook directly and `--out` to choose the output directory. In this mode, `xlflow.toml` is not required.

```bash
xlflow formulas pull --src ./Book.xlsx --out ./formula-snapshot --json
```
