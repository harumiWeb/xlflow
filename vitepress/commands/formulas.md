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

`manifest.json` lists workbook sheets, region files, formula region counts, and `ok` / `partial` / `failed` parse status summaries. `names.jsonl` contains workbook and sheet scoped defined names. Each sheet JSONL file contains logical formula regions grouped by normalized R1C1-like formula patterns. Region rows also include best-effort `refs`, `depends_on_sheets`, and `functions` indexes when xlflow can derive them from the formula text. OOXML shared formula group boundaries are coalesced when they differ only by storage metadata.

`formulas pull` supports `.xlsx` and `.xlsm`, reads OOXML files directly, and does not use Excel COM. Cached calculated values are not included. `.xlsb` workbooks are rejected with `workbook_format_unsupported` because Excel Binary Workbook storage is not an OOXML XML package.

Unsupported formula syntax is preserved as raw formula data with `partial` or `failed` parse status so one unsupported formula does not block the snapshot.

When you are already syncing workspace VBA source, `xlflow pull --formulas --json` runs the normal VBA pull first and then refreshes the default `formulas/` snapshot from the configured saved workbook.

## Standalone Use

Use `--src` to point at a workbook directly and `--out` to choose the output directory. In this mode, `xlflow.toml` is not required.

```bash
xlflow formulas pull --src ./Book.xlsx --out ./formula-snapshot --json
```

## Inspect Formula Snapshots

Use `formulas inspect` to read existing formula snapshots without opening Excel.

```bash
xlflow formulas inspect
```

The default view is a workbook summary. It includes formula region counts, formula cell counts, parse status counts, notable features, sheet dependencies, and defined names when `names.jsonl` exists.

```text
Formula summary

Sheets:
  Invoice
    formula regions: 3
    formula cells: 3000
    parse: ok 3, partial 0, failed 0
    depends on sheets:
      Config

Defined names:
  TaxRate -> =Config!$B$2
```

Inspect one sheet to list its formula regions:

```bash
xlflow formulas inspect --sheet Invoice
```

```text
Invoice formulas

D2:D1001
  pattern: =RC[-2]*RC[-1]
  example: D2 = B2*C2
  cells: 1000
  parse: ok
```

Inspect a specific cell to find the containing formula region. When the region has a supported `formula_r1c1` pattern, xlflow expands it to the requested cell.

```bash
xlflow formulas inspect --cell Invoice!E500
```

```text
Invoice!E500

Region:
  E2:E1001

Formula pattern:
  =RC[-1]*Config!R2C2

Expanded formula at Invoice!E500:
  =D500*Config!$B$2
```

Inspect a range to list overlapping formula regions:

```bash
xlflow formulas inspect --range Invoice!D2:F1001
```

Snapshots generated somewhere other than `formulas/` can be inspected with `--dir`:

```bash
xlflow formulas inspect --dir ./formula-snapshot --summary
```

Use `--json` for agent-friendly output. The inspect payload is under `output.formulas_inspect`.

```bash
xlflow --json formulas inspect --cell Invoice!E500
```

```json
{
  "status": "ok",
  "command": "formulas inspect",
  "output": {
    "formulas_inspect": {
      "view": "cell",
      "dir": "formulas",
      "workbook": "Book.xlsm",
      "cell": "Invoice!E500",
      "region": {
        "sheet": "Invoice",
        "range": "E2:E1001",
        "kind": "formula",
        "formula_r1c1": "=RC[-1]*Config!R2C2",
        "example_cell": "E2",
        "example_formula": "=D2*Config!$B$2",
        "count": 1000,
        "parse_status": "ok",
        "depends_on_sheets": ["Config"]
      },
      "expanded_formula": "=D500*Config!$B$2"
    }
  }
}
```

`partial` and `failed` formula rows remain inspectable. They are shown with their parse status and any snapshot features, such as `structured_reference`, instead of failing the command.

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow formulas` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

## Prerequisites

Check the project configuration and run `xlflow doctor --json` before workbook-backed operations. Source-only commands can run without Excel; commands that read or mutate a workbook require Windows Excel and VBIDE access.

## What this command reads and changes

The command reads the inputs and configuration described in its syntax and examples. Treat source files, the saved workbook, and a live session as separate states; add `--session` when the live workbook is authoritative. Any mutation is reversible only when a backup or explicit session save boundary exists.

## Effect on source-of-truth state

Use `xlflow status --json` before and after the command. A source edit normally requires `push`; a workbook edit normally requires `pull`; a dirty live session requires `save --session` or an intentional discard.

## Common workflows

Combine this command with the relevant [source/workbook/session workflow](../concepts/workbook-session-source), and use `--json` in scripts and agent loops.

## Common failures

Read the structured `error.code`, exit code, and recovery metadata instead of scraping terminal text. The [symptom-oriented troubleshooting guide](../help/troubleshooting) maps installation, execution, session, VS Code, and WSL failures to recovery steps.
