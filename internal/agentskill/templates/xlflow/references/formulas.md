# xlflow Formula Snapshot Reference

Use this reference when workbook behavior may depend on worksheet formulas, defined names, formula-driven ranges, or sheet layout.

Formula snapshots expose worksheet-level business logic to Git and AI agents. They complement VBA source files because many Excel workbooks keep important calculations in cells, defined names, and table references rather than in macros.

## When To Use

Read or refresh formula snapshots when:

- implementing or changing VBA that reads or writes worksheet ranges
- changing workbook layout, inserted columns, named ranges, or calculation sheets
- modifying columns used by formulas
- debugging calculation-related behavior
- reviewing formula changes in Git
- generating tests for workbook behavior
- explaining workbook business logic to a user or another agent

## Commands

From an xlflow workspace:

```bash
xlflow formulas pull --json
```

For standalone analysis without `xlflow.toml`:

```bash
xlflow formulas pull --src ./Book.xlsx --out ./formulas --json
```

`formulas pull` reads the saved `.xlsx` / `.xlsm` file directly as OOXML. It does not open Excel, use COM, evaluate formulas, recalculate, or write formulas back.

If a live Excel session has unsaved workbook changes, run `xlflow save --json` before `formulas pull`; formula extraction reads the saved file, not unsaved session state.

Inspect existing snapshots before opening raw JSONL files:

```bash
xlflow formulas inspect
xlflow formulas inspect --sheet Invoice
xlflow formulas inspect --cell Invoice!E500
xlflow formulas inspect --range Invoice!D2:F1001
xlflow formulas inspect --json
```

Use `--dir <path>` when the snapshot was generated outside the default `formulas/` directory:

```bash
xlflow formulas inspect --dir ./formula-snapshot --summary --json
```

`formulas inspect` reads existing snapshot files only. It does not open Excel, evaluate formulas, recalculate, or mutate workbook/source files.

To apply a formula pattern to a live workbook range, use `edit formula` during an xlflow session:

```bash
xlflow edit formula --sheet Invoice --range D2:D1001 --formula-r1c1 '=RC[-2]*RC[-1]' --session --json
```

Prefer `--formula-r1c1` for repeated regions because it maps directly to `formula_r1c1` in `*.regions.jsonl`. Use `--formula` only when you intentionally want Excel's normal A1 range assignment semantics. Save the session before `formulas pull` if you need the refreshed snapshot to include the edit.

## Output Layout

```text
formulas/
  manifest.json
  names.jsonl
  sheets/
    001-Invoice.regions.jsonl
```

Do not edit these generated files manually unless the user explicitly asks for a fixture or documentation example. Refresh them with `xlflow formulas pull`.

## Manifest

`manifest.json` lists sheets in workbook order and includes parse summaries:

```json
{
  "version": 1,
  "workbook": "Book.xlsm",
  "parse_status_summary": {
    "ok": 12,
    "partial": 1,
    "failed": 0
  },
  "sheets": [
    {
      "index": 1,
      "name": "Invoice",
      "sheet_id": "1",
      "path": "sheets/001-Invoice.regions.jsonl",
      "formula_region_count": 4,
      "parse_status_summary": {
        "ok": 3,
        "partial": 1,
        "failed": 0
      }
    }
  ]
}
```

Use `parse_status_summary` to decide how much of the workbook was normalized versus preserved as raw formula text.

## Region Records

Each `*.regions.jsonl` line represents a logical formula region, not one cell.

```json
{
  "range": "D2:D1001",
  "formula_r1c1": "=RC[-2]*RC[-1]",
  "example_cell": "D2",
  "example_formula": "=B2*C2",
  "count": 1000,
  "parse_status": "ok",
  "refs": ["B2:B1001", "C2:C1001"]
}
```

Important fields:

- `range`: worksheet cells covered by this formula pattern
- `formula_r1c1`: normalized R1C1-like pattern for copied formulas
- `formula`: raw formula text when normalization is partial or failed
- `example_cell`: representative cell used for `example_formula`
- `example_formula`: original A1-style formula from the example cell
- `count`: number of formula cells represented by the region
- `parse_status`: `ok`, `partial`, or `failed`
- `features`: notable conditions such as `structured_reference`, `external_reference`, or `outlier`
- `refs`: best-effort referenced A1 ranges expanded to the logical region shape
- `depends_on_sheets`: sheet names from sheet-qualified references
- `functions`: Excel function names used by the example formula
- `storage_kinds` / `storage_group_count`: OOXML storage details only; do not treat these as semantic differences by themselves

## How To Interpret R1C1 Patterns

`formula_r1c1` is a normalized pattern, not the exact formula text from every cell.

- `RC[-1]` means the cell one column to the left on the same row.
- `RC[-2]` means two columns to the left on the same row.
- `R[-1]C` means one row above in the same column.
- `R2C1` means absolute row 2, absolute column 1.
- `Config!R2C2` means absolute row 2, column 2 on the `Config` sheet.

Example:

```json
{
  "range": "F2:F1001",
  "formula_r1c1": "=RC[-2]+RC[-1]",
  "example_formula": "=D2+E2",
  "refs": ["D2:D1001", "E2:E1001"]
}
```

This means column F is calculated from columns D and E on the same row.

## Partial And Unsupported Formulas

Do not treat `parse_status:"partial"` as a failed extraction. It means xlflow preserved the formula safely but did not fully normalize every construct.

Example:

```json
{
  "range": "G4",
  "formula": "=SUM(SalesTable[Amount])",
  "example_cell": "G4",
  "example_formula": "=SUM(SalesTable[Amount])",
  "count": 1,
  "parse_status": "partial",
  "features": ["structured_reference"],
  "functions": ["SUM"]
}
```

Use the raw `formula`, `features`, and `functions` fields for reasoning. Structured references, external references, 3D references, spill references, and implicit intersection may intentionally stay raw in early implementations.

## Defined Names

`names.jsonl` contains workbook-scoped and sheet-scoped defined names.

```jsonl
{"name":"TaxRate","scope":"workbook","refers_to":"=Config!$B$2","kind":"formula"}
{"name":"InvoiceTotal","scope":"Invoice","refers_to":"=Invoice!$G$12","kind":"formula"}
```

Always check `names.jsonl` when formulas reference constants, workbook parameters, named ranges, or sheet-scoped aliases. A named value such as `TaxRate` may explain a formula that otherwise looks disconnected from its source cell.

## Inspect Command

Prefer `xlflow formulas inspect` when you need a quick, human-readable or agent-readable view of snapshot contents.

Summary view:

```bash
xlflow formulas inspect --summary
```

Use it to answer:

- which sheets contain formulas
- how many formula regions and formula cells each sheet has
- whether formulas parsed as `ok`, `partial`, or `failed`
- which notable features appear, such as `structured_reference`
- which sheets are referenced by formulas
- which defined names exist

Sheet view:

```bash
xlflow formulas inspect --sheet Invoice
```

Use it to list formula regions on one sheet, including each region's pattern or raw formula, example formula, cell count, parse status, features, and dependencies.

Cell view:

```bash
xlflow formulas inspect --cell Invoice!E500
```

Use it to find which formula region contains a specific cell. When the region has a supported `formula_r1c1` pattern, inspect also shows the expanded A1 formula at the requested cell, such as `=D500*Config!$B$2`. If the cell is not part of a formula region, the command reports that clearly.

Range view:

```bash
xlflow formulas inspect --range Invoice!D2:F1001
```

Use it to find all formula regions overlapping a range before changing columns, VBA range writes, named constants, or sheet layout.

JSON output:

```bash
xlflow --json formulas inspect --cell Invoice!E500
```

The JSON envelope uses `command:"formulas inspect"` and places the view payload under `output.formulas_inspect`. Use this form when another agent or script needs structured fields such as `regions`, `region`, `expanded_formula`, `defined_names`, `features`, or `depends_on_sheets`.

## Agent Workflow

Before editing VBA or workbook layout:

1. Read `xlflow.toml` and identify the workbook and source directories.
2. If the workbook has live session changes, save first with `xlflow save --json`.
3. Run `xlflow formulas pull --json` if the snapshot may be stale, or use existing `formulas/` snapshots if they are known current.
4. Run `xlflow formulas inspect --summary --json` to identify formula-heavy sheets, parse statuses, defined names, features, and sheet dependencies.
5. Use `xlflow formulas inspect --sheet <name> --json`, `--cell <sheet!A1> --json`, or `--range <sheet!A1:B2> --json` to narrow the affected formula regions.
6. Open raw `manifest.json`, `names.jsonl`, or relevant `sheets/*.regions.jsonl` files only when inspect output is not detailed enough for the task.
7. Identify whether the VBA or layout change touches formula input ranges, formula output ranges, defined names, or dependent sheets.
8. Make the code or workbook-source change.
9. Run xlflow lint/tests/macros as appropriate.
10. Re-run `xlflow formulas pull --json` after workbook formula or layout changes and review the JSONL diff.

## Review Checklist

When reviewing formula snapshot diffs:

- Large copied formula changes should usually appear as one compact region diff, not thousands of cell-level changes.
- One-cell regions near large repeated regions may be intentional exceptions or outliers; inspect `features:["outlier"]` when present.
- A changed `formula_r1c1` usually means calculation logic changed across a copied region.
- A changed `example_formula` with the same `formula_r1c1` may be harmless if only the representative cell moved, but verify the range and refs.
- `refs` and `depends_on_sheets` are best-effort indexes, not a full Excel calculation graph.
- `storage_kinds` and `storage_group_count` describe OOXML storage. They should not drive business-logic conclusions by themselves.
- `partial` rows still matter. Read the raw `formula` and `features`.

## Safety Rules

- Do not assume VBA is the only source of workbook behavior.
- Do not ignore formula regions when changing worksheet columns, ranges, or named constants.
- Do not interpret one JSONL line as one cell; it may represent thousands of cells.
- Do not treat `formula_r1c1` as exact A1 formula text.
- Do not treat `partial` as extraction failure.
- Do not hand-edit generated `formulas/` outputs as the source of truth.
- Do not rely on saved-file formula snapshots when the live session workbook has unsaved changes.
