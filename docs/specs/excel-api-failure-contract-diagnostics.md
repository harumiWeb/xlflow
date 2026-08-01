# Excel API Failure-Contract Diagnostics

## Scope

`VBA218` is a default-enabled, high-precision, interprocedural `analyze`
warning. It runs in batch analysis and the shared real-time editor path,
supports normal inline suppression, and does not block `push` or `run` source
preflight.

The rule resolves the receiver and member before classifying a call. It does
not report late-bound or unresolved calls, and it does not infer API identity
from text alone. A finding says that the API *can* fail when no result exists;
it does not assert that an error is guaranteed.

## Failure Contracts

The initial exception-raising API set is:

- `Range.SpecialCells`
- `WorksheetFunction.Match`
- `WorksheetFunction.VLookup`
- `WorksheetFunction.XLookup`
- `WorksheetFunction.Index`

Calls in this set require either an enclosing `On Error GoTo <label>` error
path, or a narrow `On Error Resume Next` probe that immediately checks `Err`
and restores `On Error GoTo 0`.

The initial `Variant/Error` API set is:

- `Application.Match`
- `Application.VLookup`
- `Application.XLookup`

These calls require `IsError` before their returned value is consumed. The
rule recognizes direct `IsError(...)` checks and assignments whose successful
branch is dominated by an `IsError` guard. Direct consumption, an unchecked
assignment, and use in a nested expression remain findings.

## Local Wrappers

The analyzer recognizes conservative local function summaries. A function
that catches an exceptional API result and returns `CVErr(...)` is treated as
handling the exception, while its callers still need to guard the resulting
`Variant/Error`. A Boolean local function that returns an `IsError` result is
treated as a guard alias.

Same-module uniquely resolved wrappers participate in batch and real-time
analysis. Cross-module uniquely resolved wrappers participate in batch
analysis only. Dynamic, unresolved, and name-shaped calls do not suppress a
finding.

## Configuration

Disable the rule with `[analyze].disabled_rules = ["VBA218"]`. The legacy
`[analyze].detect_excel_api_failure_contracts` Boolean remains accepted for
compatibility and emits the standard deprecation warning. Disabling the rule
removes it from both batch and real-time analysis.
