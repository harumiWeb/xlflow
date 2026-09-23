# Excel semantic inspections

This specification defines opt-in diagnostics for Excel-specific source forms
whose binding or maintenance cost is hidden by VBA shorthand.

<!-- xlflow-rule-contract: {"id":"VBA260","family":"analyze","category":"maintainability","default_severity":"warning","scope":"procedure-local","realtime":false,"configuration_key":"detect_worksheet_string_access","inline_suppressible":true,"preflight_blocking":false} -->

## VBA260: worksheet string access

VBA260 reports an explicit `ThisWorkbook.Worksheets("name")` or
`ThisWorkbook.Worksheets.Item("name")` access only when authoritative workbook
metadata maps the visible worksheet name to one document-module CodeName and
that module exists in the analyzed project. Dynamic selectors, external or
unknown workbooks, `Sheets` entries that are not proven worksheets, missing or
ambiguous metadata, and callers that do not supply workbook metadata fail open.
The filesystem-backed analyzer may derive the metadata from the configured
saved OOXML workbook. `AnalyzeProject` never reads a workbook implicitly.

<!-- xlflow-rule-contract: {"id":"VBA261","family":"analyze","category":"reliability","default_severity":"information","scope":"procedure-local","realtime":true,"configuration_key":"detect_application_worksheet_function_dispatch","inline_suppressible":true,"preflight_blocking":false} -->

## VBA261: Application worksheet-function dispatch

VBA261 reports a call whose receiver resolves exactly to `Excel.Application`
and whose member exists in the complete generated `Excel.WorksheetFunction`
member set. It includes explicit `Application.Member`, typed Application
variables, and `With Application`; it excludes shadowed identifiers,
`Object`/`Variant` receivers, bare calls, explicit `Application.WorksheetFunction`
calls, non-worksheet-function members, and incomplete TypeLib views. The
diagnostic explains that explicit `WorksheetFunction` binding can have a
different return or error contract; it does not synthesize an Application
signature or replace existing VBA218/VBA251 ownership.

<!-- xlflow-rule-contract: {"id":"VBA262","family":"analyze","category":"maintainability","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_host_bracket_expressions","inline_suppressible":true,"preflight_blocking":false} -->

## VBA262: Excel host bracket expression

VBA262 reports a bare Excel host bracket expression such as `[A1]` from the raw
procedure expression projection and recommends an explicit object-model access.
The rule does not treat array indexing, array bounds, bracketed declarations,
strings, comments, or qualified members such as `Me.[Member]` as host bracket
expressions. Because xlflow analyzes Excel VBA projects, the rule is an explicit
Excel-host policy and remains disabled unless configured.
