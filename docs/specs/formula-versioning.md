# Formula Versioning

This spec defines the formula snapshot contract for `xlflow formulas pull`.

## Command

`xlflow formulas pull` reads a workbook and writes formula metadata under `formulas/`.

By default, the command reads the workbook configured by `[excel].path` in `xlflow.toml`:

```bash
xlflow formulas pull
```

It can also run outside an xlflow workspace by specifying the source workbook directly:

```bash
xlflow formulas pull --src path/to/Book.xlsx --out path/to/formulas
```

For workspace source sync, `xlflow pull --formulas` is a convenience path that runs the normal VBA `pull` first and then refreshes the default `formulas/` snapshot from the saved configured workbook if the VBA pull succeeds. The standalone `xlflow formulas pull` command remains the canonical formula extraction command and should be used for `--src` / `--out` workflows.

Formula snapshots are read-only artifacts. To apply one explicit formula pattern to a live workbook range, use `xlflow edit formula --session --sheet <name> --range <A1:B2> --formula-r1c1 <formula>`. That command is a low-level live-session edit primitive, not a snapshot push/apply workflow.

The command:

- supports `.xlsx` and `.xlsm` files;
- rejects `.xlsb` with `workbook_format_unsupported` because Excel Binary Workbook storage is not OOXML XML;
- accepts `--src <workbook>` to override the configured workbook and skip `xlflow.toml` loading;
- accepts `--out <dir>` to choose the output directory; the default is `formulas`;
- reads workbook files directly as OOXML zip packages;
- does not launch Excel, use COM, evaluate formulas, recalculate, or write formulas back;
- replaces the generated `formulas/` directory on each successful run;
- omits volatile metadata such as generation timestamps.

Extraction failures for valid `.xlsx` or `.xlsm` files return `formulas_pull_failed`. A source workbook path that is not `.xlsx` or `.xlsm` returns `formulas_pull_args_invalid`, except project workbook formats that are not supported by the formula snapshot backend, such as `.xlam` and `.xlsb`, which return `workbook_format_unsupported` with `workbook.format` and `workbook.capability = "formulas pull"`.

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
  "parse_status_summary": {
    "ok": 2,
    "partial": 1,
    "failed": 0
  },
  "sheets": [
    {
      "index": 1,
      "name": "Invoice",
      "sheet_id": "1",
      "path": "sheets/001-Invoice.regions.jsonl",
      "formula_region_count": 3,
      "parse_status_summary": {
        "ok": 2,
        "partial": 1,
        "failed": 0
      }
    }
  ]
}
```

The manifest is ordered by workbook sheet order. `parse_status_summary` counts emitted formula regions by parse status at workbook and sheet level.

## Defined Names

`names.jsonl` contains workbook-scoped and sheet-scoped defined names:

```jsonl
{"name":"TaxRate","scope":"workbook","refers_to":"=Config!$B$2","kind":"formula"}
{"name":"InvoiceTotal","scope":"Invoice","refers_to":"=Invoice!$G$12","kind":"formula"}
```

Names are sorted workbook scope first, then sheet order, then name. `refers_to` always includes a leading `=`.

## Formula Regions

Each sheet JSONL file contains one logical formula region per line. Normal and shared OOXML formulas are treated as storage forms of the same semantic formula region. Adjacent regions are grouped vertically when contiguous cells in the same column have the same normalized R1C1-like pattern, parse status, and compatible features.

```jsonl
{
  "range": "G2:G1000",
  "formula_r1c1": "=RC[-2]*RC[-1]",
  "example_cell": "G2",
  "example_formula": "=E2*F2",
  "count": 999,
  "parse_status": "ok",
  "refs": [
    "E2:E1000",
    "F2:F1000"
  ]
}
```

Single-cell deviations stay as one-cell regions. When xlflow can identify a one-cell deviation between larger matching regions in the same column, it adds `outlier` to `features`.

Shared formulas are read from their OOXML anchor and `ref` range, but `shared_index`, anchor, and shared-group boundaries are storage metadata and do not define canonical regions. If Excel stores one copied formula range as multiple adjacent shared groups with the same normalized pattern, xlflow coalesces them:

```jsonl
{
  "range": "D2:D101",
  "formula_r1c1": "=RC[-2]*RC[-1]",
  "example_cell": "D2",
  "example_formula": "=B2*C2",
  "count": 100,
  "parse_status": "ok",
  "refs": [
    "B2:B101",
    "C2:C101"
  ],
  "storage_kinds": [
    "shared"
  ],
  "storage_group_count": 2
}
```

When available, region rows include lightweight formula intelligence fields:

- `refs`: referenced A1 ranges expanded to the logical formula region shape;
- `depends_on_sheets`: sheet names from sheet-qualified references;
- `functions`: Excel function names used by the example formula.

These fields are best-effort indexes for review and AI-agent context. They do not attempt full Excel dependency graph resolution.

`storage_kinds` identifies non-default OOXML storage forms that contributed to a canonical region. `storage_group_count` is emitted only when multiple storage-level groups were coalesced. Plain normal formula regions omit storage metadata.

If a formula cannot be fully normalized, xlflow keeps raw formula text and continues:

```jsonl
{
  "range": "D10",
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

## Inspecting Snapshots

`xlflow formulas inspect` reads the generated snapshot files without opening Excel or recalculating the workbook. It defaults to the `formulas/` directory and can inspect another snapshot directory with `--dir`.

Supported views:

- `xlflow formulas inspect` or `xlflow formulas inspect --summary` prints a workbook-level overview.
- `xlflow formulas inspect --sheet <name>` prints all formula regions for one sheet.
- `xlflow formulas inspect --cell <sheet!A1>` finds the formula region containing a cell.
- `xlflow formulas inspect --range <sheet!A1:B2>` lists formula regions overlapping a range.

The inspect command treats `partial` and `failed` parse statuses as inspectable snapshot data. They do not fail the command unless the snapshot file itself is missing or malformed. Defined names from `names.jsonl` are included in summary output when present.

For cell inspection, xlflow attempts a best-effort expansion of supported normalized R1C1 patterns back to A1 references at the requested cell. Supported forms include relative and absolute cell references such as `RC[-2]`, `R1C1`, `R[1]C1`, ranges such as `R[-6]C[-2]:R[-1]C[-2]`, and sheet-qualified references including quoted sheet names. Absolute R1C1 rows and columns render with `$` in A1 output. If a formula pattern cannot be expanded safely, inspect still shows the matching region and raw formula fields without `expanded_formula`.

## Initial OOXML Scope

The command parses:

- `xl/workbook.xml`
- `xl/_rels/workbook.xml.rels`
- `xl/worksheets/sheet*.xml`

Later issues may add tables, external links, calculation chain metadata, connections, or richer dependency summaries.
