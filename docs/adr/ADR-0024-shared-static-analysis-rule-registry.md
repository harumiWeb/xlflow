# ADR-0024: Shared Static-Analysis Rule Registry

## Status

Accepted

## Context

xlflow exposes VBA static-analysis diagnostics through several surfaces: `lint`,
`analyze`, source preflight, the LSP server, the VS Code extension, project
configuration, and user documentation. Rule identity and behavior were
previously repeated across Go maps and switches, TypeScript allowlists, and
hand-maintained Markdown tables. Adding a rule therefore required coordinated
edits that were difficult to verify, and consumers could disagree about whether
a rule was configurable, suppressible, enabled by default, or safe for real-time
editor feedback.

ADR-0013 assigns semantic runtime-risk checks to `analyze`, while ADR-0014 keeps
VBA analysis protocol-neutral and makes the VS Code extension a thin LSP client.
Neither decision defines a shared source of rule metadata. A registry is needed
to preserve those boundaries while allowing CLI, LSP, editor, configuration,
and documentation adapters to project the same facts.

## Decision

Create a protocol-neutral registry under `internal/staticanalysis/rules`. Its
canonical data is an embedded JSON file owned by that package. Each normal
`VB...` or `VBA...` static-analysis rule records its stable ID, title,
description, family, category, severity, scope, precision, default-enabled
state, configuration binding, inline-suppression policy, preflight policy,
real-time eligibility, fix availability, and documentation URL.

The Go package validates the complete registry at load time and exposes sorted,
defensive projections through lookup, enumeration, and family filtering. It
does not import CLI, LSP, or configuration types. Configuration keeps a small
adapter that maps registry configuration keys to `LintConfig` and
`AnalyzeConfig` getters and setters; tests require the adapter and configurable
registry entries to match exactly.

The registry is authoritative for these projections:

- `lint`, `analyze`, and source preflight rule identity and policy;
- inline suppression eligibility and preflight-blocking behavior;
- LSP diagnostic severity and `codeDescription.href` documentation links;
- the source-only, parallel-safe `xlflow rules` command;
- VS Code suppression Quick Fix eligibility; and
- the generated static-analysis diagnostic catalog.

`xlflow rules --json` publishes schema version 1 and the sorted registry items
inside the normal success envelope. Version 1 evolves additively: consumers
must ignore unknown fields and rule IDs, and breaking field changes require a
new `schema_version`. The VS Code extension caches successful metadata by
resolved CLI path and fails closed when retrieval fails, the schema is
unsupported, or a diagnostic ID is unknown. It must not offer an inline
suppression Quick Fix without affirmative registry metadata.

The JSON file is embedded so the installed binary remains self-contained, while
the documentation generator reads the same repository file directly and does
not require a Go build. Generated documentation is checked byte-for-byte.

`VBA000` remains outside the registry because it represents an analysis pipeline
failure rather than a static-analysis rule. UserForm `FRM...` and `UFY...`
diagnostics are also outside this decision. Existing `disabled_rules`, legacy
per-rule booleans and warnings, nested `VB044` configuration, and inline
suppression syntax retain their compatibility behavior.

## Consequences

Positive consequences:

- Rule metadata has one reviewable source and deterministic ordering.
- Configuration, preflight, LSP, VS Code, CLI, and documentation drift can be
  detected with focused tests and generated-file checks.
- Editor integrations can discover capabilities from the installed CLI instead
  of shipping a duplicated rule allowlist.
- The registry stays reusable by protocol adapters and documentation tooling.

Negative consequences:

- Every normal `VB...` or `VBA...` diagnostic must be registered before it can
  pass validation and integration tests.
- The embedded JSON and public schema become compatibility surfaces that need
  deliberate additive evolution.
- Configuration still requires an explicit adapter because the registry must
  not depend on concrete config structures.
- Older xlflow binaries may not provide `rules`; editor consumers must preserve
  fail-closed fallback behavior rather than guessing.

## Alternatives Considered

1. **Keep independent Go, TypeScript, and Markdown lists.** Rejected because the
   existing duplication is the source of policy drift and cannot be made
   reliable through reviewer discipline alone.
2. **Make Go structs the canonical source and generate JSON by running Go.**
   Rejected because documentation generation and freshness checks should remain
   portable Node-based repository operations that do not require CGO or a Go
   toolchain.
3. **Make the public CLI response or LSP protocol types canonical.** Rejected
   because rule ownership is protocol-neutral; transport envelopes are adapters,
   not the domain model.
4. **Let the VS Code extension infer suppressibility from diagnostic prefixes.**
   Rejected because preflight-blocking and non-suppressible rules do not follow
   a safe prefix-only rule, and unknown future diagnostics must fail closed.

## Rationale

- Specs: `docs/specs/cli-contract.md` defines the public command, JSON schema,
  configuration compatibility, and suppression contract.
- Code: `internal/staticanalysis/rules`, `internal/config`, `internal/lspserver`,
  `internal/cli`, and `editors/vscode` consume registry projections through
  their own adapter boundaries.
- Docs: `scripts/docs/generate-reference-inventory.mjs` and
  `vitepress/reference/diagnostics.md` prove that the canonical JSON can drive a
  deterministic catalog without building Go.
- Tests: registry validation, config-adapter completeness, CLI/LSP contracts,
  VS Code fail-closed behavior, and generated-document freshness cover the
  cross-surface decision.

## Supersedes

- None

## Superseded by

- None

## Related

- `docs/adr/ADR-0013-analyze-runtime-risk-ownership.md`
- `docs/adr/ADR-0014-reusable-vba-lsp-server.md`
- `docs/specs/cli-contract.md`
- `vitepress/commands/rules.md`
- `vitepress/reference/diagnostics.md`
