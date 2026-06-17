# ADR-0012: Cross-Platform pack Command for Source-to-Artifact Workbook Generation

## Status

Accepted

## Context

xlflow treats VBA projects as source-controlled code. Most commands already operate on the source tree and are cross-platform. `push` is the one source-to-workbook command that still requires Windows and Excel: writing VBA back into a workbook means regenerating `xl/vbaProject.bin`, and today that happens through the Excel/VBIDE bridge as a live read-modify-write of the loaded workbook.

CI pipelines, containers, release packaging, and headless or agent-driven workflows need to produce an `.xlsm` artifact from source without a Windows-plus-Excel host. Everything else inside an `.xlsm` is OOXML XML that any platform can already produce. `xl/vbaProject.bin` — an OLE/CFB container holding MS-OVBA-compressed sources — is the single opaque piece that keeps that capability Windows-bound.

ADR-0008 already drew the relevant boundary. It placed source-tree management, lint/format/analysis, and release packaging in the Go core, and stated that source-only commands remain cross-platform while Windows/Excel automation stays in the bridge. Generating `xl/vbaProject.bin` from on-disk `.bas`/`.cls` sources is a source-to-artifact transform of the same kind: it works purely at the file level, never touches COM, VBIDE, or Win32, and never launches Excel.

A pure-Go, file-level writer cannot, however, provide what the live `push` path provides. It performs no VBE compile validation, no runtime validation, and gives no guarantee that every host-specific project state is interpreted exactly as Excel would. It is therefore not a drop-in replacement for `push`.

Feasibility is established. `ovba-writer` (https://github.com/kay-ws/ovba-writer), a standalone pure-Go reference implementation, reads and writes source-only `vbaProject.bin`, round-trips real Excel-compiled workbooks (cross-checked against `olevba`), and produces bins that real Excel recompiles and runs.

## Decision

Introduce a new, experimental, cross-platform command `xlflow pack` that builds an `.xlsm` artifact from the source tree plus a workbook template, entirely in Go at the file level. It regenerates `xl/vbaProject.bin` from source and replaces that one entry inside the workbook zip.

`pack` is a separate command, not a mode or backend of `push`:

- `push` is unchanged. It remains the Excel/VBIDE-backed live-session path and still requires Windows and Excel.
- The backend a command uses must not depend on project configuration. xlflow will not add an `xlflow.toml` switch that turns `push` into a file-level writer, because that would make a command's meaning depend on config and be confusing in CI logs, in support and debugging, and for agents.

The initial MVP boundary is deliberately narrow and fail-loud:

- the command is gated behind `--experimental`;
- `--out` is required; `pack` never overwrites the template or configured source workbook in place;
- `pack` operates only on closed workbook files, never on active sessions or live workbooks;
- `--template` is optional and falls back to the source workbook configured in `xlflow.toml`;
- standard and class modules are supported first; document modules only where they map safely against the template;
- JSON output identifies the backend as the experimental pure-Go packer and reports that VBE compile validation was not performed;
- the pure-Go packaging path has Linux tests;
- unsupported cases fail with specific errors rather than best-effort behavior.

Unsupported in the MVP, each a loud and specific error: active sessions or live workbooks, in-place overwrite of the template/source workbook, protected VBA projects, signed VBA projects, full UserForm/`.frx` generation, and unknown or ambiguous VBA project layouts.

The contract for command shape, the JSON envelope, and exit codes lives in `docs/specs/pack-command.md`.

### Ownership and attribution

`pack`'s implementation, test harness, and compatibility behavior live in an xlflow-internal package, `internal/pack`. xlflow owns the public behavior and the maintenance responsibility, because this is a core artifact-generation path.

`ovba-writer` is a reference implementation and feasibility proof, not a long-term dependency. xlflow does not take a direct module dependency on it; the logic is reimplemented in-repo and the reference is credited by link. Both `ovba-writer` and xlflow are MIT-licensed, so reuse is license-compatible. Because `pack` is not a `go.mod` dependency, no entry in `THIRD_PARTY_LICENCES.md` is required.

### Staged scope

`pack` matures before it loses the experimental gate:

- **Initial MVP** — as described above.
- **Before non-experimental status** — broader document-module fixtures; hardened signed/protected project detection beyond the MVP's baseline reject-on-detect; non-ASCII/Japanese source fixtures; Windows/Excel smoke tests that open the generated workbook and compile/run a minimal macro; a documented UserForm preservation/update strategy; and docs stating that `pack` does not compile or run VBA.
- **UserForms, staged in three steps** — (1) preserve the template's existing designer streams unchanged, generating no forms; (2) update form code-behind while keeping the template's designer state; (3) full reconstruction from exported `.frm`/`.frx` as a separate, higher-risk phase. Only step (1)'s "carry existing designer streams through untouched" is compatible with the MVP, and only when the forms already exist in the template.

### No VBE validation

"No VBE validation" is a permanent semantic boundary, not a temporary limitation. `pack` never compiles or executes VBA. Its output is a file artifact whose correctness against the VBE has not been verified; Excel compiles it from source on first open. Every `pack` run reports this in its output. Consumers that need compile or runtime validation must use the Excel/VBIDE-backed `push` path on Windows.

## Consequences

- Positive: `push` semantics are unchanged; the live-session workflow does not regress.
- Positive: a cross-platform path exists for CI, containers, release packaging, and headless or agent artifact generation, with no Windows-plus-Excel host.
- Positive: keeping the backend independent of configuration means a command's meaning stays stable and legible in CI logs and to agents.
- Positive: ownership in `internal/pack` keeps the artifact-generation path's maintenance and compatibility behavior inside xlflow.
- Negative: `pack` output lacks VBE compile and runtime validation; consumers must treat it as unvalidated, and the contract must keep saying so.
- Negative: two commands (`pack`, `push`) with overlapping inputs but different guarantees increase the surface to document and explain.
- Negative: xlflow takes on maintenance of MS-OVBA and OLE/CFB compatibility behavior in-repo rather than delegating it to an external module.
- Negative: the experimental surface needs staged hardening — fixtures, detection, smoke tests — before it can be considered stable.

## Alternatives Considered

1. **Make `push` itself cross-platform / Excel-free** — Rejected because it would change `push`'s validation guarantees and conflate the live-session, VBE-validated semantics with an unvalidated file-level transform under one command name.
2. **Add an `xlflow.toml` switch that changes `push`'s backend** — Rejected because a command's behavior would then depend on project configuration, which is confusing in CI logs, in support and debugging, and for agents.
3. **Implement Excel automation in Go** (ADR-0008 Alternative #2) — Not applicable and rejected. `pack` never touches COM, VBIDE, or Win32 and never launches Excel. It is a file-level transform, not automation, so it does not recreate interop complexity in the CLI process.
4. **Take a long-term direct dependency on `ovba-writer`** — Rejected because a core artifact-generation path should not depend on an external module whose maintenance xlflow cannot control. The implementation is internalized under `internal/pack`, with the reference credited.
5. **Include full UserForm/`.frx` generation in the MVP** — Rejected as higher-risk. UserForm support is staged, starting with carrying existing template designer streams through untouched.

## Related

- `docs/adr/ADR-0008-dotnet-excel-bridge.md`
- `docs/adr/ADR-0001-agent-ready-vba-cli-architecture.md`
- `docs/specs/pack-command.md`
- `docs/specs/cli-contract.md`
- https://github.com/kay-ws/ovba-writer
- xlflow issue #143
