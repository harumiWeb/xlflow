# Unsafe VBA Select Operations

<!-- xlflow-rule-contract: {"id":"VBA250","family":"analyze","category":"runtime-safety","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_unsafe_select_operations","inline_suppressible":true,"preflight_blocking":false} -->

`VBA250` reports intentional Excel UI selections whose required active
workbook or worksheet cannot be proven on every reachable path. It is a
default-enabled, warning-level, non-blocking analyzer rule available to batch
and realtime/LSP diagnostics. It is independent of `VB003`, which reports the
style and maintainability risk of using `Select` or `Activate` at all.

## Public contract

The procedure-local abstract state tracks two facts:

```text
ActiveWorkbook  = known workbook identity | unknown
ActiveWorksheet = known worksheet identity | unknown
```

Procedure entry and any uncertain state-changing operation start with unknown
facts. A branch join retains a fact only when the same identity is established
on every incoming reachable path. Conditional activation therefore does not
make a later selection safe.

The rule recognizes resolved Excel object-model references and common local
aliases, including `ThisWorkbook`, worksheet variables, `Worksheets(...)`,
`Range`/`Cells` roots, direct `Set` aliases, and `With` blocks. Unresolved,
late-bound, non-Excel, and cross-procedure effects remain unknown rather than
being treated as proof of safety.

The supported state transitions and checks are:

- `Workbook.Activate` establishes that workbook as active and clears the
  active worksheet fact.
- `Worksheet.Activate` establishes its parent workbook and that worksheet as
  active.
- `Worksheet.Select` requires its parent workbook to be known active. A
  successful select may establish the selected worksheet only when its
  selection mode is known to replace the current selection.
- `Range.Select` requires its owning worksheet to be known active. Selecting a
  range does not establish a new workbook or worksheet fact.

`On Error Resume Next`, unresolved calls, dynamic dispatch, and other
unmodeled state changes do not establish an activation guarantee. A finding
is emitted at the unsafe `Select` call with the missing workbook or worksheet
precondition and a suggestion to activate the required object first.

## Configuration and suppression

The compatibility configuration key is:

```toml
[analyze]
detect_unsafe_select_operations = false
```

Projects may instead use the shared rule policy:

```toml
[analyze]
disabled_rules = ["VBA250"]
```

An intentional local exception can use `xlflow:disable-line VBA250` or
`xlflow:disable-next-line VBA250`. Suppression only removes the finding; it
does not change the modeled active-state facts or make the source operation
safe. Because `preflight_blocking` is false, an unsuppressed warning does not
stop workbook source preflight.

## Compatibility boundaries

The initial rule is procedure-local. Activation performed by a helper
procedure is not assumed to establish state in its caller. `Application.Goto`
and other state-changing APIs without a dedicated transfer model remain
unknown. `VBA250` does not replace, weaken, or deduplicate `VB003` findings.

## Verification requirements

Focused coverage must include direct qualified references, local aliases,
`With` blocks, unconditional and conditional activation, branch joins,
`Worksheet.Select`, `Range.Select`, and suppression/configuration behavior.
Batch and realtime projections must agree on the same source revision and
stable source range. The rule is runtime-safety evidence and does not require
VBE-oracle promotion.

## Related

- Issue #792
- `docs/adr/ADR-0052-unsafe-select-runtime-state.md`
- `docs/specs/cli-contract.md`
