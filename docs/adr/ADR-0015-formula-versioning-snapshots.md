# ADR-0015: Pure-Go Formula Versioning Snapshots

## Status

Accepted

## Context

xlflow already makes workbook VBA source reviewable, but many workbooks also carry important business logic in worksheet formulas. A cell-level formula export would be technically simple but noisy: copied formulas across thousands of rows would create large files and unreadable diffs.

ADR-0008 and ADR-0012 keep a clear boundary between Excel automation and pure file-level workbook operations. Formula extraction is an inspection and source snapshot task. It can read OOXML package parts directly and does not need COM, VBIDE, formula evaluation, or recalculation.

## Decision

Add `xlflow formulas pull` as a pure-Go command that reads the configured `.xlsx` or `.xlsm` workbook and writes deterministic formula snapshots under `formulas/`.

The command is separate from VBA `xlflow pull`:

- `pull` remains the Excel/VBIDE-backed VBA export command.
- `formulas pull` is a file-level OOXML reader and never launches Excel.
- The command writes region-based JSONL by default, grouping contiguous same-column formulas that share the same normalized R1C1-like pattern.
- Workbook and sheet-scoped defined names are written to `formulas/names.jsonl`.
- Unsupported formulas are preserved as raw formula data with `partial` or `failed` parse status instead of aborting extraction.

The public contract lives in `docs/specs/formula-versioning.md`.

## Consequences

- Positive: workbook formula logic becomes visible to humans, source control, and AI agents without opening Excel.
- Positive: region JSONL keeps large copied-formula sheets reviewable.
- Positive: `pull` semantics stay unchanged for VBA source export.
- Positive: the implementation can run in non-Windows environments because it reads OOXML directly.
- Negative: the snapshot does not evaluate formulas, recalculate workbooks, or prove formula correctness.
- Negative: initial region grouping is vertical only; rectangular grouping is deferred.
- Negative: unsupported Excel formula syntax may appear as raw partial or failed entries until the normalizer grows broader support.

## Alternatives Considered

1. **Add `xlflow pull --formulas`** - Rejected because `pull` already means VBA export through the Excel bridge. Mixing pure file-level formula extraction into the same command would blur backend and source-of-truth expectations.
2. **Cell-level JSON snapshots** - Rejected because large copied-formula sheets would produce noisy diffs and poor agent context.
3. **Use Excel COM to read formulas** - Rejected for the initial feature because saved workbook OOXML already contains formula metadata and COM would make the command Windows/Excel-bound.
4. **Implement a full Excel formula AST first** - Rejected as too broad. The reference-aware normalizer is sufficient for the first region grouping contract, while unsupported syntax can remain raw.

## Related

- `docs/specs/formula-versioning.md`
- `docs/specs/cli-contract.md`
- `docs/adr/ADR-0008-dotnet-excel-bridge.md`
- `docs/adr/ADR-0012-pack-command.md`
- xlflow issue #227
