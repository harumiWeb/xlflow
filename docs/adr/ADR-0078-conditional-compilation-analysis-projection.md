# ADR-0078: Conditional-compilation analysis projection

## Status

Accepted

## Context

Issue #966 contains ordinary prose inside an inactive `#If` branch. VBE
accepts the module, but parsing every branch makes lint emit VB014 and analyze
fail before push can deliver the source. Waiving VB014 does not make an
incomplete syntax tree suitable for analysis.

Analysis must remain deterministic without opening Excel. A host OS or process
architecture does not identify the target Excel installation, and project
conditional arguments are not necessarily available from source files.

## Decision

Use one byte-preserving source projection before diagnostic parsing. Only
provably inactive content becomes spaces; newlines and active source positions
remain. Resolved conditional wrappers also become spaces so split declaration
headers join their shared bodies; unknown wrappers remain. Preserve the original source for editing,
formatting, fingerprints and workbook delivery. Raw structural parsers keep
their existing behavior.

Module `#Const` directives form a sequential, module-local environment. Explicit
`[vba.conditional_constants]` values provide source-controlled project arguments
and target built-ins. Undefined ordinary names follow VBA Empty semantics.
Unspecified built-ins and expressions outside the supported evaluator subset
remain unknown. Unknown branches remain available to conservative existing IR
and diagnostic consumers, and conditional assignments are merged before later
directives are evaluated.

Conditional evaluation extends ADR-0041's shared expression parser through an
explicit entry point; it does not widen declaration-constant coercions. Editor
snapshots must keep original source identity separate from projected analysis
identity. A conditional change can affect source outside the edited range, so
incremental reuse must not assume that only that range changed.

## Consequences

- Valid inactive prose no longer prevents lint, analyze or source preflight.
- Batch and editor analysis share the same environment and projection.
- Analysis remains Excel-free; configuration never changes workbook compiler
  arguments. Users must align explicit arguments with their target workbook.
- Unknown conditions can still expose parser recovery in candidate branches.
  This is preferable to silently dropping potentially active code.
- Retaining unknown wrappers preserves conditional uncertainty rules. A
  resolved branch supplies ordinary active syntax, including shared bodies.
- Conservative fallbacks trade some coverage and conditional-edit reuse for
  correctness. Malformed directive structures are never hidden by projection.

## Alternatives Considered

1. Suppress VB014 only. Rejected because analyze still requires a usable tree
   and text-based diagnostics would continue examining inactive code.
2. Assume the analysis host is the target Excel platform. Rejected because
   x86/x64, Windows/Mac and VBA versions are properties of the target.
3. Change the raw parser or rewrite sources before push. Rejected because
   formatting and delivery must preserve every authored branch.
4. Query Excel during analysis. Rejected because offline CLI and LSP snapshots
   must be deterministic and independent of workbook automation.

## Evidence

- `internal/vba/conditional/project_test.go`
- `internal/vba/constexpr/conditional_test.go`
- `docs/specs/vba-conditional-compilation.md`
- Issue #966 and the `issue966-*` local VBE oracle fixtures.
