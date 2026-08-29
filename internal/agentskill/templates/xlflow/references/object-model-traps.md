# Excel Object-Model Traps Reference

Load this reference when correctness depends on Excel's runtime behavior rather
than on the source: lookups over data whose order you do not control, criteria
strings, inserting or deleting cells under existing formulas or merges, sorting
mixed-type columns, or any call that exists under both `WorksheetFunction` and
VBA.

Every trap here shares one property: the source is not wrong on its face. Static
analysis cannot separate the correct use from the silent defect, because the
answer depends on the data. Where a lint rule can catch a mistake, you do not
need to remember it. These are what is left.

Each entry states when to care, what fails quietly, what to do instead, how to
prove it, and how far the statement reaches.

## Approximate Lookup Over Unsorted Data

**Trigger.** `Match`, `VLookup`, `HLookup` or `Lookup` with approximate matching
— the default — against a range whose sort order you have not established in
the same procedure.

**Risk.** A key below every value in the range returns `#N/A` in a formula, and
raises at the `WorksheetFunction` call — that case is at least loud. The
dangerous one is a key inside the range: a position comes back with no error and
no `#N/A`, and it is not the nearest value. The result may also be correct by
accident on the data you tested with, and wrong on the data your user has.

**Safe rule.** Approximate matching is only defined when the required order is
established *and* verified. If you cannot show that it holds, pass exact matching
explicitly: `Match(key, rng, 0)`, `VLookup(key, tbl, 2, False)`. Prefer exact
matching unless a banded or bucketed lookup is genuinely intended, and do not
rely on any particular wrong answer.

Sorting to satisfy the rule is a structural edit, not a lookup detail. Sort the
whole associated dataset, never the key column alone — sorting a `VLookup` key
range by itself breaks its rows away from the columns it returns — and treat
reordering the user's data as a change that needs its own justification.

**Proof.** Assert the behavior you are relying on, not the one you are avoiding.
Fill a range out of order and assert that exact matching finds the key wherever
it sits; or establish the required order in the procedure, assert the order
itself, and then assert the approximate result. Run it through the session proof
loop in SKILL.md (`session start` → `push --fast --session --no-save` → `test
--session --no-save` → `save --session` → `session stop`). Do not assert the
position an unsorted approximate lookup happens to return: that is the observed
behavior below, and it is not a contract to hold a test against.

**Scope / provenance.** Excel documents approximate matching as requiring
ascending order, and documents nothing about unsorted input. Observed on Excel
16 (Windows 11, ja-JP): with `A1:A5` holding `10, 50, 20, 40, 30`,
`WorksheetFunction.Match(35, Range("A1:A5"), 1)` returns `3` — the cell holding
`20`, not the nearest value at or below the key. That is consistent with a
binary search reporting where it stopped, but treat it as an observation of one
implementation, not as a contract. Descending approximate matching over shuffled
input was not stable enough to state at all.

## Criteria Comparisons Are Not Symmetric

**Trigger.** `CountIf`, `SumIf`, `AverageIf` and their `*Ifs` forms, or
`AutoFilter`, given a criteria string — especially against a column that mixes
numbers with text that spells numbers.

**Risk.** `"<>20"` does not count what `"=20"` leaves out. The two overlap, so
their counts can exceed the number of cells; both are right at once and the
totals will not reconcile, with no error to notice.

**Safe rule.** Do not derive one criterion from another by negation. When a
column may hold text that looks numeric, either normalize the column before
counting, or state the comparison you actually mean and verify both halves
against the row count.

**Proof.** Build a range holding the text `"20"`, the number `20`, and the text
`"20.0"`, then assert `CountIf` for `"=20"`, `"<>20"` and `">=20"` in one test
run through the session proof loop. The signal is that `"=20"` and `"<>20"`
together account for five matches over three cells: they overlap rather than
partition.

**Scope / provenance.** Observed on Excel 16 (Windows 11, ja-JP): `"=20"` counts
3, `"<>20"` counts 2, `">=20"` counts 1. Equality coerces text that spells a
number; the ordering operators do not. Excel documents criteria loosely enough
that this is worth verifying on your own version before depending on either
number.

## Insert And Delete Move Absolute References

**Trigger.** `Range.Insert`, `Range.Delete`, `Rows.Insert`, `Columns.Delete`, or
any structural edit above or across a range that existing formulas refer to.

**Risk.** "Absolute means fixed" is a reasonable rule carried over from copying
formulas, and it is wrong here. `$A$2` is unchanged by a copy and becomes `$A$3`
when a row is inserted above it. Formulas keep pointing at the value they meant,
which is usually what you want — but if your code assumed the address text was
stable, it is now reading a different cell.

**Safe rule.** Never assume an address literal survives a structural edit. Re-read
addresses after inserting or deleting, or refer to the data through a defined
name or a table column, which are maintained across the edit.

**Safe rule for ranges.** A range reference may shift or resize, depending on
where the structural edit meets it. An insert *above* `SUM(A1:A3)` shifts it to
`SUM(A2:A4)`; an insert *inside* it resizes it to `SUM(A1:A4)`; a deletion across
it gives `SUM(A1:A2)`; an insert past its end leaves it exactly as it was; only a
range removed in its entirety becomes `#REF!`. Either way the reference is
adjusted silently, so do not treat the absence of `#REF!` as evidence that a
formula still covers what it used to.

**Proof.** Snapshot formulas with `xlflow formulas pull --json` before and after
the structural edit and compare the regions that matter. The snapshot reads the
saved file, so run `xlflow save --json` first.

**Scope / provenance.** This is documented Excel behavior for reference
adjustment; it is listed here because the sensible reading of "absolute" points
the other way, not because it is undocumented.

## A Partial Shift Divides Formulas And Merges Oppositely

**Trigger.** `Range.Insert` or `Range.Delete` with `xlShiftDown` or `xlShiftRight`
over a band narrower than the region it crosses — that is, any partial shift on a
sheet holding merged cells or multi-column formula ranges.

**Risk.** The same containment test produces opposite outcomes. A range that
reaches past the shifted band is left alone; a merge that the band only half
covers is taken apart. Code that reasons about one will be wrong about the other.

**Safe rule.** Before a partial shift, establish which merges the band intersects
and decide explicitly whether losing them is acceptable. If it is not, shift whole
rows or columns instead, or unmerge and re-merge deliberately. Verify merges after
the edit rather than assuming they survived.

**Proof.** After the shift, assert both halves in one test: the formula that
should be untouched and the merge that should not be. `xlflow formulas pull
--json` covers the formula side; read `MergeArea` for the merge side.

**Scope / provenance.** Observed on Excel 16 (Windows 11, ja-JP):
`Range("B2").Insert xlShiftDown` rewrites `B3`, leaves `C3` alone, and leaves
`SUM(A1:C3)` alone because that range reaches past column B — while a merge only
half covered by the same band is broken. The formula half follows from documented
reference adjustment; the merge half is an observation.

## Sorting Orders By Kind Before Value

**Trigger.** `Range.Sort` or `AutoFilter` sorting over a column that may hold more
than one type, or any `Sort` call that omits `Header`.

**Risk.** Two silent failures. Values do not interleave, they group by kind: in
ascending order all numbers come before all text, which comes before Booleans,
which come before error values; descending reverses that order of kinds as well
as the values inside each one. So a "sorted" column can look shuffled to a user
reading values. And `Header` defaults to `xlNo`, so an omitted argument sorts the
header row in with the data.

**Safe rule.** Pass `Header:=xlYes` explicitly whenever the range includes a
header row — do not rely on Excel guessing. Normalise a mixed column to one type
before sorting if the user expects values to interleave.

**Safe rule for blanks.** Blanks sit outside the ordering of kinds: they sort to
the bottom in both directions. A descending sort does not bring them to the top,
so "the last row is blank" is not a reliable end-of-data test after sorting.

**Proof.** Sort a column holding a number, a text value, a Boolean, an error and
a blank — ascending and descending — and assert the resulting order cell by cell.
Assert both directions: a test that only covers ascending will not notice code
that assumes the kinds stay in the same order when the direction flips.

**Scope / provenance.** Excel documents the type ordering and the `Header`
default. Observed on Excel 16 (Windows 11, ja-JP), over a column holding `text`,
`5`, `TRUE`, a blank, `#DIV/0!`, `apple`, `100` and `FALSE`: ascending gives
`5, 100, apple, text, FALSE, TRUE, #DIV/0!, [blank]` and descending gives
`#DIV/0!, TRUE, FALSE, text, apple, 100, 5, [blank]`. Locale affects text
collation, so assert on the grouping by kind rather than on the order of specific
strings if the workbook may run elsewhere.

## The Same Name Is Two Different Functions

**Trigger.** Any call that exists in both places: `Round`, `Trim`, and others
reachable as both `WorksheetFunction.X` and a VBA built-in.

**Risk.** Both answers are correct in their own language, which is what makes the
pair dangerous in one file. A refactor that adds or removes the
`WorksheetFunction` qualifier changes the result without changing the intent, and
nothing in the diff looks like a behavior change.

| expression          | `WorksheetFunction`        | VBA                    |
| ------------------- | -------------------------- | ---------------------- |
| `Round(2.5, 0)`     | `3` — half away from zero  | `2` — half to even     |
| `Trim("  a   b  ")` | `a b` — inner runs collapse | `a   b` — ends only   |

**Safe rule.** Choose the rounding rule the requirement states, and write it so
the choice is visible: qualify the call and say why in a comment, or wrap it in a
named helper. Never change a qualifier during unrelated cleanup. For whitespace,
`Trim` and `WorksheetFunction.Trim` are not interchangeable at all — pick by
whether inner runs must collapse.

**Proof.** Assert both boundary cases directly. `Round(2.5, 0)` tells you which
rounding rule is in force — `3.5` agrees under both rules and proves nothing. For
whitespace, assert `Trim("  a   b  ")` against
`WorksheetFunction.Trim("  a   b  ")`: one keeps the inner run and the other
collapses it.

**Scope / provenance.** Both behaviors are documented in their own references;
they are listed here because the collision is easy to miss and expensive to find.
Observed on Excel 16 (Windows 11, ja-JP).

## Safety Rules

- Treat any approximate lookup as unverified until the required order is both
  established and asserted; otherwise pass exact matching.
- Do not derive a criteria string by negating another one.
- Re-read addresses after a structural edit; do not trust an address literal
  across `Insert` or `Delete`.
- Check merges after a partial shift, separately from checking formulas.
- Pass `Header` explicitly to every `Sort` over a range that has one, and assert
  a mixed-type sort in both directions rather than only ascending.
- Prove the boundary case whenever a rounding or trimming rule matters.
- When an entry above is marked as an observation, verify it on the target Excel
  version before writing code that depends on the specific value.
