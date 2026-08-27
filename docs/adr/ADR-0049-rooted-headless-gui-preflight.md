# ADR-0049: Rooted Headless GUI Preflight

## Status

Accepted

## Context

`run --headless` previously rejected a run whenever the configured source tree
contained any GUI boundary. This made an unrelated legacy procedure such as a
file-picker helper block a non-interactive entrypoint that could not call it.
The repository already has a conservative VBA call graph and rooted
reachability implementation for `VB021`, but its public/event roots describe
project-wide runtime entrypoints and are not suitable as the single root for a
requested macro.

## Decision

Add a headless-only GUI analysis that uses the requested macro as a confirmed
root and reuses the shared calls snapshot and call graph. Filter only boundaries
whose owning procedures are definitely unreachable. Follow statically resolved
dynamic targets as possible reachability. If source parsing, procedure
ownership, or project dispatch is uncertain, retain the existing project-wide
rejection behavior. External and builtin-like calls do not by themselves make
the whole project uncertain. Standard-module Property Get/Let/Set accessors are
possible roots because VBA can invoke them implicitly without a syntax-local
call edge.

Keep `inspect-gui`, `lint`, and `doctor` project-wide. Preserve the existing
`gui_boundary_detected` error code, exit status, JSON boundary shape, and
non-headless behavior. Run the GUI check before a `run --push` import so a
rejected headless command does not mutate the workbook first.

## Consequences

- Unreachable GUI code in legacy modules no longer blocks an unrelated safe
  headless macro.
- Reachable, ambiguous, dynamic, malformed, or ownerless cases remain blocked,
  preserving the safety property of headless preflight.
- Headless preflight performs a source parse and graph construction in addition
  to the existing source preflight; this cost is paid only for headless runs.
- Standard-module property accessors remain possible roots because VBA can call
  them implicitly without a syntax-local call edge.
- The analysis remains limited to configured source files and cannot prove
  safety for code outside that source tree or unsupported late-bound dispatch.

## Alternatives considered

1. Keep project-wide rejection and add a manual allowlist - rejected because it
   shifts call-graph knowledge to configuration and can silently become stale.
2. Reuse the VB021 root set unchanged - rejected because public and host roots
   intentionally model project-wide entrypoints and recreate the reported bug.
3. Treat every unresolved call as a project-wide fallback - rejected because
   ordinary external and builtin calls such as `Debug.Print` and `Range` would
   defeat the useful precision without improving safety.

## Evidence

- GUI boundary scanner: `internal/gui/analyzer.go`.
- Shared call extraction and dynamic references: `internal/vba/calls`.
- Rooted graph traversal: `internal/vba/callgraph` and
  `internal/vba/reachability`.
- CLI headless preflight: `internal/cli/root.go`.
- Contract: `docs/specs/vba-call-graph-reachability.md` and
  `docs/specs/runtime-debugging.md`.

## Related

- ADR-0025: Rooted VBA Call-Graph Reachability for VB021.
- Issue #746.
