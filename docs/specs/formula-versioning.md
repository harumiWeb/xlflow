# Formula Versioning

This spec defines the formula snapshot contract for `xlflow formulas pull`.

## Command

`xlflow formulas pull` reads the workbook configured by `[excel].path` in `xlflow.toml` and writes formula metadata under `formulas/`.

The command:

- supports `.xlsx` and `.xlsm` files;
- reads workbook files directly as OOXML zip packages;
- does not launch Excel, use COM, evaluate formulas, recalculate, or write formulas back;
- replaces the generated `formulas/` directory on each successful run;
- omits volatile metadata such as generation timestamps.

Unsupported workbook files or extraction failures return `formulas_pull_failed`. A configured workbook path that is not `.xlsx` or `.xlsm` returns `formulas_pull_args_invalid`.

## Output Layout

```text
formulas/
  manifest.json
  names.jsonl
  sheets/
    001-Invoice.regions.jsonl
    002-Summary.regions.jsonl
```

Sheet region filenames use the workbook sheet index and a sanitized sheet name. If sanitized names collide, xlflow appends a deterministic numeric suffix.

## Manifest

`manifest.json` has schema version `1`:

```json
{
  "version": 1,
  "workbook": "Book.xlsm",
  "sheets": [
    {
      "index": 1,
      "name": "Invoice",
      "sheet_id": "1",
      "path": "sheets/001-Invoice.regions.jsonl",
      "formula_region_count": 3
    }
  ]
}
```

The manifest is ordered by workbook sheet order.

## Defined Names

`names.jsonl` contains workbook-scoped and sheet-scoped defined names:

```jsonl
{"name":"TaxRate","scope":"workbook","refers_to":"=Config!$B$2","kind":"formula"}
{"name":"InvoiceTotal","scope":"Invoice","refers_to":"=Invoice!$G$12","kind":"formula"}
```

Names are sorted workbook scope first, then sheet order, then name. `refers_to` always includes a leading `=`.

## Formula Regions

Each sheet JSONL file contains one formula region per line. Normal formulas are grouped vertically when contiguous cells in the same column have the same normalized R1C1-like pattern and parse status.

```jsonl
{
  "range": "G2:G1000",
  "kind": "normal",
  "formula_r1c1": "=RC[-2]*RC[-1]",
  "example_cell": "G2",
  "example_formula": "=E2*F2",
  "count": 999,
  "parse_status": "ok"
}
```

Single-cell deviations stay as one-cell regions. When xlflow can identify a one-cell deviation between larger matching regions in the same column, it adds `outlier` to `features`.

Shared formulas are represented from the shared anchor and OOXML `ref` range instead of expanding child cells:

```jsonl
{
  "range": "C2:C1000",
  "kind": "shared",
  "shared_index": "0",
  "anchor": "C2",
  "formula_r1c1": "=RC[-2]*RC[-1]",
  "example_cell": "C2",
  "example_formula": "=A2*B2",
  "count": 999,
  "parse_status": "ok"
}
```

If a formula cannot be fully normalized, xlflow keeps raw formula text and continues:

```jsonl
{
  "range": "D10",
  "kind": "normal",
  "formula": "=Table1[Amount]",
  "example_cell": "D10",
  "example_formula": "=Table1[Amount]",
  "count": 1,
  "parse_status": "partial",
  "features": [
    "structured_reference"
  ]
}
```

Cached calculated values from `<v>` elements are not part of the canonical snapshot.

## Initial OOXML Scope

The command parses:

- `xl/workbook.xml`
- `xl/_rels/workbook.xml.rels`
- `xl/worksheets/sheet*.xml`

Later issues may add tables, external links, calculation chain metadata, connections, or richer dependency summaries.
