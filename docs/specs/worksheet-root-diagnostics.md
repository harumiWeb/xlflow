# Worksheet Root Diagnostics

## Scope

`analyze` provides two procedure-local reliability diagnostics for Excel range
expressions. They run in batch analysis and in the editor's shared real-time
analysis path. The rules are independent of `VBA205`: general active-sheet
dependency remains a separate concern from last-row boundary reliability.

## `VBA216`: Proven Worksheet Root Mismatch

`VBA216` is enabled by default, has error severity, blocks source preflight,
and cannot be suppressed inline. It is emitted only when an outer `Cells` or
`Range` expression and an object inside its arguments have statically proven,
distinct worksheet identities.

The analyzer recognizes only these comparable identity forms:

- two discovered worksheet codenames;
- two `ThisWorkbook.Worksheets` / `ThisWorkbook.Sheets` selectors whose sole
  arguments are string literals.

`Worksheets("Data")` and `Sheets("Data")` canonicalize to the same name
identity. Index and dynamic selectors, comparisons between selector forms, an
unknown variable, `ActiveSheet`, and bare Excel members are not proof and must
not emit `VBA216`. A worksheet variable inherits a proven identity from `Set`;
once it receives more than one distinct proven identity in a procedure, it is
ambiguous and no longer participates in blocking comparisons.

Worksheet codenames come only from the project workbook document-module source
root (excluding `ThisWorkbook`), which is the source-layout contract for
worksheet document modules. A name that merely resembles `Sheet1` is unknown.

The analyzer evaluates complete VBA logical statements, including explicit
line continuations, and evaluates `With` headers before adding their root to
the `With` stack. A finding is attached to the first physical line of the
logical statement and includes a fully qualified replacement example when one
can be formed.

## `VBA217`: Unstable Last-Row Boundary

`VBA217` is enabled by default, warning severity, non-blocking, and supports
normal inline suppression. It is limited to last-row contexts: assignments to
`lastRow`, `last_row`, `endRow`, or `end_row`, and expressions ending in
`.End(...).Row`.

It reports implicit worksheet roots, `.End(xlDown).Row`, an unadjusted
`UsedRange.Rows.Count`, and `CurrentRegion.Rows.Count`. A fully qualified
`UsedRange.Row + UsedRange.Rows.Count - 1` expression is accepted. Guidance
prefers a stable column-based calculation such as
`ws.Cells(ws.Rows.Count, column).End(xlUp).Row` where appropriate.

## Configuration and Compatibility

The rules use `[analyze].disabled_rules` entries `VBA216` and `VBA217`, and
the compatible boolean keys `detect_worksheet_root_mismatch` and
`detect_unstable_last_row_patterns`. Disabling either rule removes it from
batch analysis and real-time editor diagnostics. `VBA216` retains its
non-suppressible preflight contract when enabled.
