# Local Excel/VBE oracle

The VBE oracle is an optional developer tool for answering focused VBA
compile-time questions with a real local Excel/VBE instance. It is evidence for
static-analysis maintenance; it is not part of the `xlflow` production command
surface, a normal test dependency, or a GitHub Actions check.

## Prerequisites and safety

Oracle execution requires Windows, a locally installed Excel desktop client,
and **Trust access to the VBA project object model** enabled in Excel's Trust
Center. The command creates one unsaved disposable workbook per case and must
be run sequentially. Do not run two oracle processes at the same time: Excel,
the VBE command bars, modal dialogs, and cleanup ownership are single-user
resources.

The command only terminates the Excel process it created and recorded. If a run
reports an unconfirmed cleanup or an infrastructure failure, inspect Task
Manager before starting another run and do not promote a fixture.

Dialog detection is tied to the target Excel process/VBE owner chain through
the existing Win32/UI Automation watcher. The oracle does not use
`SendKeys` or keyboard-focus scripting.

## Running cases

From the repository root, run one case through the PowerShell wrapper:

```powershell
./scripts/dev/test-vbe-oracle.ps1 --case byref-parenthesized-variable
```

The direct equivalent is:

```powershell
go run ./cmd/xlflow-vbe-oracle --case byref-parenthesized-variable
```

Omit `--case` to select all cases in the deterministic order in
`testdata/vbe-oracle/manifest.json`. Repeat `--case` to select a focused set.
The runner always executes the known-accept and known-reject controls first.
Use `--timeout 5m` (the default) to set the per-case timeout. Progress is sent
to stderr; stdout is a schema-versioned JSON result suitable for scripts.

To run the health controls explicitly in strict mode:

```powershell
./scripts/dev/test-vbe-oracle.ps1 `
  --case known-compile-accept `
  --case known-compile-reject `
  --strict
```

## Fixture format

The manifest has `schema_version: 1`, a `controls` object with `accept` and
`reject` case IDs, and an ordered `cases` list. The two control entries also
carry `control_role` metadata for reviewers. Each case directory contains a
`case.json` and ordinary `.bas` sources. V1 accepts standard modules only and
requires `probe.mode: "compile"`; runtime probes and class/UserForm modules are
future extensions.

Case expectations start as observational fixtures:

```json
{
  "vbe": {
    "expected": "observe",
    "evidence_phase": "unknown",
    "diagnostic_meaning": "observation"
  },
  "provenance": { "status": "pending", "verified_on": [] }
}
```

Source is kept in normal module files so the same fixture can be consumed by
parser, lint, analyze, type-analysis, LSP, and oracle tests. The analyzer
contract uses `analysis.expected_diagnostics` and
`analysis.forbidden_diagnostics`; entries may specify `code`, `severity`, an
optional source range, and (when needed) `surfaces` from `lint`, `analyze`, and
`lsp`.

## Outcomes and strict mode

Only `accepted` and `rejected` are VBA evidence. A compile dialog associated
with the target Excel/VBE process is `rejected`. A successful Compile command
with no delayed compile dialog and confirmed cleanup is `accepted`. A disabled
Compile command is `infrastructure_failure` with `compile_invoked: false`
because no VBE Compile evidence was observed.

Timeouts, Excel startup/VBOM/import/project activation failures, inability to
find or invoke Compile, worker crashes or malformed output, unknown/unrelated
modals, COM disconnection, and unconfirmed cleanup are
`infrastructure_failure`. They are never converted to VBA acceptance or
rejection. A cleanup failure overrides any earlier accepted/rejected result.

Default observational runs print what Excel reported and do not change fixture
files. `--strict` fails when a case is still `observe` or when an asserted
expectation does not match. The known controls are a harness health gate: if
the accept control is rejected, or the reject control is accepted, the run
stops and the batch is unhealthy.

Exit codes are stable for automation: `0` success/observation, `1` strict
expectation mismatch, `2` CLI or fixture validation error, and `3` oracle
infrastructure failure.

## Promoting observed evidence

Promotion is explicit and must select one or more case IDs:

```powershell
./scripts/dev/test-vbe-oracle.ps1 `
  --case byref-parenthesized-variable `
  --promote-observed
```

Promotion is atomic for the selected batch. It refuses an already asserted
fixture, an implicit all-case selection, infrastructure failures, or an
unconfirmed cleanup. A rejected compile is classified as `compile-error`.
Accepted compile evidence requires an explicit diagnostic meaning of
`specification`, `policy`, or `maintainability`; `runtime-error` cannot be
proven by a compile probe. Supply meanings with the command's
`--diagnostic-meaning <case-id>=<meaning>` option. Promotion records the
observed phase and Excel version/build, process bitness, UI locale, and UTC
verification timestamp in `provenance.verified_on`.

## Static-analysis workflow

VBE is authoritative for whether this source compiles, but not for policy or
maintainability advice. A rule must distinguish compile-equivalent errors from
deterministic runtime-risk warnings and non-VBE policy/maintainability
warnings. Normal Go and .NET tests consume committed fixture expectations and
never launch Excel.

When changing static-analysis semantics (including call/argument validation,
type inference, object/`Set` diagnostics, parser interpretation, severity, or
LSP projections):

1. Decide whether the behavior depends on actual VBE semantics.
2. Add or update an `observe` fixture if there is no existing evidence.
3. Run both known controls and the focused oracle case IDs locally.
4. Promote only accepted/rejected observations with confirmed cleanup.
5. Run deterministic Go and .NET tests.
6. Include the executed oracle case IDs, or a reason for non-applicability, in
   the pull request description.

Oracle infrastructure failures are stop-the-line failures for agents and
contributors. Do not change analyzer behavior or promote fixtures based on a
timeout, unknown modal, failed Compile invocation, or unconfirmed Excel
cleanup.

## Pull request checklist

```markdown
- [ ] This change does not alter static-analysis behavior.
- [ ] Relevant VBE oracle cases were run locally:
      `<case-id-1>`, `<case-id-2>`
- [ ] A VBE oracle run was not applicable because:
      `<reason>`
```

The oracle remains local-only until an appropriate Excel-installed self-hosted
runner exists. No CI workflow may install, launch, or automate Excel for this
tool.
