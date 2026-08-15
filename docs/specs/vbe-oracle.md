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

The command only terminates the Excel process it created and recorded. Other
Excel instances may already be running or start while the oracle executes;
they are neither attached to nor terminated. Concurrent instances discovered
during cleanup are reported under `cleanup.concurrent_processes` for
observability, but do not invalidate cleanup after every process whose oracle
ownership is proven by PID plus start time has exited. If a run reports an
unconfirmed owned-process cleanup or another infrastructure failure, inspect
Task Manager before starting another run and do not promote a fixture.

The oracle acquires a per-user Windows cross-process lock for the complete
batch, including the known controls, selected cases, and any promotion. A
second local oracle process fails immediately with the structured
`oracle_already_running` infrastructure error. The lock is an OS-owned file
lock, so normal process termination and crash termination release it without
stale-lock cleanup or user confirmation. The lock does not participate in
normal xlflow workbook coordination.

Dialog detection is tied to the target Excel process/VBE owner chain through
the existing Win32/UI Automation watcher. The oracle does not use
`SendKeys` or keyboard-focus scripting.

After injecting each fixture module, the oracle reads the `CodeModule` back
and requires it to match the sanitized fixture source, apart from line-ending
normalization and trailing line breaks. If VBE rewrites the source during
injection, the run fails with `oracle_source_mutated`; compile evidence from
the rewritten source must not be used to classify the committed fixture.

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

The per-case timeout applies to the VBE probe. The runner reserves an
additional bounded cleanup margin for the bridge to close the disposable
workbook, quit Excel, release COM/UI references, drain owned-process shutdown,
and emit structured output before the outer process deadline.

## Cleanup diagnostics and recovery

Each bridge observation includes a `cleanup` object. It reports the owned
Excel process identity, whether exit was confirmed for each owned process, the
number of drain attempts, remaining dialog/process counts, and the failure
stage. Only `owned_processes` contains processes proven to belong to the
oracle; unrelated Excel processes are excluded from that list, reported under
`concurrent_processes`, and never terminated.

For example:

```json
{
  "cleanup": {
    "confirmed": false,
    "owned_processes": [
      {
        "pid": 1234,
        "start_time": "2026-08-06T10:00:00.0000000Z",
        "exit_confirmed": false,
        "ownership_basis": "captured-process-pid"
      }
    ],
    "concurrent_processes": [],
    "drain_attempts": 3,
    "remaining_windows": 0,
    "remaining_processes": 1,
    "failure_stage": "process-exit-confirmation"
  }
}
```

A process listed under `concurrent_processes` could not be proven to belong to
the oracle. It is intentionally left alone and does not affect the result. A
`process-exit-confirmation` failure means an owned process could not be proven
to have exited within the bounded drain. After that failure, inspect Task
Manager for the recorded owned PID and any dialogs, confirm that no
oracle-owned Excel remains, and rerun the controls. Do not remove lock files
manually and do not promote fixtures from that report.

`failure_stage` is `null` after confirmed cleanup. Infrastructure responses
that did not start Excel use `not-attempted`; cleanup failures use
`process-exit-confirmation` or `temporary-directory`.
The execution stage for platform, startup, import, and compile failures is
reported separately in `last_stage`.

## Fixture format

The manifest has `schema_version: 1`, a `controls` object with `accept` and
`reject` case IDs, and an ordered `cases` list. The two control entries also
carry `control_role` metadata for reviewers. Each case directory contains a
`case.json` and ordinary `.bas` sources. V1 requires `probe.mode: "compile"`
and supports standard, class, UserForm, and document modules; runtime probes
remain a future extension.

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
parser, lint, analyze, type-analysis, LSP, and oracle tests. Schema v1 accepts
`modules[].kind` values `standard`, `class`, `form`, and `document`; document
modules must also set `document_target` to `workbook` or `worksheet`. Existing
standard fixtures omit these fields and remain valid. The bridge injects the
sanitized `.bas` body into the corresponding component (document probes use
the disposable workbook's ThisWorkbook or a worksheet); UserForm designer and
`.frx` artifacts are not part of the oracle contract. The analyzer
contract uses `analysis.expected_diagnostics` and
`analysis.forbidden_diagnostics`; entries may specify `code`, `severity`, an
optional source range, and (when needed) `surfaces` from `lint`, `analyze`, and
`lsp`.

Every fixture must also declare its diagnostic binding state and a non-empty
`evidence_role` under `analysis`:

```json
{
  "analysis": {
    "binding_status": "unbound",
    "evidence_role": "language-observation"
  }
}
```

The supported states are:

- `unbound`: VBE evidence exists, but no analyzer rule contract is connected.
  It cannot declare expected or forbidden diagnostics. Use an evidence role to
  distinguish a pending compile-equivalent result from a language, policy, or
  maintainability observation; a non-empty `binding_note` may explain the
  missing contract.
- `partially-bound`: at least one rule is connected, but coverage is incomplete;
  `evidence_role` must be `compile-equivalent`, and `rule_codes` plus a
  non-empty `binding_note` are required.
- `bound`: the fixture is fully connected to declared rules. It requires at
  least one rule code, `evidence_role` must be `compile-equivalent`, and every
  declared code must appear in the contract that matches the VBE result
  (`expected` for rejected cases, `forbidden` for accepted cases).
- `not-applicable`: the fixture is a harness control. It must use the
  `harness-control` evidence role and cannot declare rule codes or diagnostic
  contracts. A language observation is `unbound`, not `not-applicable`.

`analysis.evidence_role` records what the fixture can support independently of
the VBE outcome and must not be used to rewrite `vbe.expected`,
`vbe.evidence_phase`, `vbe.diagnostic_meaning`, or `provenance`:

- `compile-equivalent`: the VBE result is relevant to a compile-equivalent
  diagnostic. A bound or partially-bound fixture has a connected contract;
  an unbound fixture records a pending compile-equivalent rule with no current
  analyzer contract.
- `language-observation`: an unbound fixture retaining VBE language behavior
  for a rule that has no connected analyzer contract yet.
- `policy-observation`: an unbound fixture where accepted VBA may still
  receive a safety or policy warning, such as a missing-`Set` assignment.
- `maintainability-observation`: an unbound fixture where accepted VBA may
  still receive a maintainability warning, such as a parenthesized `Sub` call.
- `harness-control`: a `not-applicable` known accept/reject control used to
  verify the oracle itself.

The validator requires `compile-equivalent` for `bound` and
`partially-bound`, `harness-control` for `not-applicable`, and rejects
`harness-control` on `unbound` fixtures. This keeps compile-equivalent binding
coverage separate from accepted-language, policy, and maintainability
observations.

Every declared rule code is resolved against the authoritative static-analysis
rule registry. The registry validates the canonical diagnostic ID, supported
`lint`/`analyze`/`lsp` surfaces, and supported severities; fixture validation
rejects unknown or non-canonical codes, unsupported surfaces, and unsupported
severities. A `compile-equivalent` fixture binding must reference registry
rules with `compile_equivalent: true`; rejected bound expectations for those
rules must use `severity: "error"`. Bound and partially-bound fixtures require
every compile-equivalent contract code to be listed in `rule_codes`; bound
fixtures additionally require every declared
compile-equivalent rule code to appear in the contract that matches the VBE
result (`expected` for rejected cases, `forbidden` for accepted cases). A
bound accepted compile-equivalent fixture may additionally list a registered
non-compile-equivalent, non-blocking style or maintainability rule (for
example, `VB066`) and put it in the explicit `expected` warning contract. That
supplemental rule is not treated as VBE compiler evidence or a
preflight-blocking binding. Until a fixture is
connected to an implemented rule, keep it `unbound` and do not alter the VBE
expectation to satisfy an analyzer test. Binding notes must be omitted when
unnecessary; an explicitly empty or whitespace-only note is invalid.

## Positive/negative binding pairs

Rejected fixtures may identify the accepted fixtures that protect their rules
from false positives:

```json
{
  "analysis": {
    "binding_status": "bound",
    "evidence_role": "compile-equivalent",
    "rule_codes": ["VBAxxx"],
    "negative_controls": ["optional-argument-omitted", "known-named-argument"],
    "expected_diagnostics": [{ "code": "VBAxxx", "surfaces": ["analyze", "lsp"] }]
  }
}
```

`negative_controls` is valid only on rejected `partially-bound` or `bound`
fixtures. Each ID must name a distinct, existing, VBE-accepted, bound fixture;
self-references, rejected controls, duplicate IDs, and cycles are invalid. The
accepted fixture must forbid the same rule code. A bound rejected fixture is
complete only when the union of its referenced controls' forbidden surfaces
covers every surface in its positive contract. When a contract omits surfaces,
the rule registry's supported surfaces are used for this comparison.

The corpus validator also checks that every bound rule has rejected positive
evidence and an accepted negative control. Existing `unbound` and
`partially-bound` fixtures remain visible in the report and do not fail CI by
themselves; malformed relationships and incomplete `bound` rules do fail. A
valid pair attached to a `partially-bound` rejected fixture is informational
until that parent fixture becomes `bound`, but its controls must still forbid
one of the parent's declared rule codes.

## Binding coverage report

Run the Excel-free deterministic report with:

```powershell
rtk go test ./internal/oracle -run TestOracleBindingCoverage -v
```

The test reports fixture state counts, complete and incomplete rule counts,
sorted fixture IDs for every state (`bound`, `partially-bound`, `unbound`, and
`not-applicable`), and sorted rule-to-case and surface coverage. The committed
corpus currently reports 91 asserted fixtures: 75 `bound`, 0
`partially-bound`, 15 `unbound`, and 2 `not-applicable`; the fifteen bound rules
have complete positive/negative coverage. The report is emitted before a
validation failure so missing positive/negative evidence and state changes
remain visible in CI logs.

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

### Procedure terminator compatibility audits

Procedure-boundary audits keep three observations separate: the parser/CST and
xlflow boundary tracker describe source structure, a real VBE Compile probe
answers whether Excel accepts that source, and the registry describes xlflow's
style/preflight policy. A parser mismatch is therefore not sufficient evidence
for a compile-equivalent error, and VBE acceptance does not prohibit an
independent non-blocking style warning.

Issue #597 audits the complete 5-by-3 opener/terminator matrix for each of the
four supported module targets. The five openers are `Sub`, `Function`,
`Property Get`, `Property Let`, and `Property Set`; the three terminators are
`End Sub`, `End Function`, and `End Property`. The module targets are a
standard module, a class module, `ThisWorkbook`, and a worksheet document
module. Each matrix cell is an independent fixture with exactly one procedure
boundary, so a second syntax or declaration error cannot decide the compile
result.

Before Excel is run, every matrix fixture is `vbe.expected = "observe"`,
`provenance.status = "pending"`, and `analysis.binding_status = "unbound"`
with `analysis.evidence_role = "compile-equivalent"`. The fixture records the
parser/CST result and current diagnostic projection separately from the oracle
result. Known accepted and rejected harness controls run first; the focused
cases then run sequentially. Only confirmed `accepted` or `rejected` outcomes
may be promoted. A timeout, unknown modal, failed Compile invocation,
source mutation, COM/worker failure, or unconfirmed cleanup is infrastructure
failure and cannot be promoted or used to change analyzer severity.

Promotion and binding follow the result, rather than the parser's expectation:

| Evidence result                | Contract after promotion                                                                                                                                                          |
| ------------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| VBE rejects the mismatch       | Bind the rejected fixture to the compile-equivalent error contract and require `error` severity; it may block source preflight.                                                   |
| VBE accepts the mismatch       | Forbid the compile-equivalent error on `lint` and `lsp`; any retained parser/style concern must be an explicitly non-blocking policy or maintainability observation.              |
| Mixed accepted/rejected matrix | Split only the confirmed rejected subset into the compile-equivalent contract and use a separately registered, non-blocking style rule for any retained accepted-form preference. |
| Infrastructure failure         | Keep the fixture observational and stop the batch; do not infer VBA validity.                                                                                                     |

The promoted issue #597 run used Excel 16.0, build 17932, x64, `ja-JP`, and
completed all 60 cases after the known controls. All four module targets had
the same result: matching `Sub`, `Function`, and `Property Get/Let/Set`
terminators were accepted; `Property Get` with `End Sub` or `End Function`
was also accepted; every other mismatched pair was rejected with the expected
`End Sub`, `End Function`, or `End Property` compile dialog. Rejected fixtures
are bound to `VB012` with `error` contracts on `lint` and LSP and matching
accepted baselines are negative controls. Accepted mismatches forbid `VB012`
and are covered by `VB066` style regressions; they are not compile-equivalent
bindings. The accepted source contracts also verify that the confirmed
structure does not leave a duplicate blocking `VB014` recovery finding.

The registry and generated diagnostics reference now reflect that split:
`VB012` remains compile-equivalent and preflight-blocking for VBE-rejected
forms, while `VB066` is a default-enabled, high-precision, file-local,
inline-suppressible warning that is not compile-equivalent and never blocks
preflight. Parser structure, compiler validity, and xlflow style preference
must remain documented and tested as separate contracts.

When changing static-analysis semantics (including call/argument validation,
type inference, object/`Set` diagnostics, parser interpretation, severity, or
LSP projections):

1. Search existing oracle fixtures before creating new evidence.
2. Add or update a minimal rejected `observe` fixture when needed.
3. Add or identify one or more adjacent accepted controls.
4. Run both known controls and the focused oracle case IDs locally.
5. Promote only accepted/rejected observations with confirmed cleanup.
6. Implement the diagnostic using confirmed VBA behavior.
7. Add `expected_diagnostics` with registry `severity: "error"` to rejected
   compile-equivalent fixtures.
8. Add `forbidden_diagnostics` and `negative_controls` to accepted/rejected
   binding pairs, respectively.
9. Assign `evidence_role`: use `compile-equivalent` only for registry
   `compile_equivalent` bindings, and keep policy, maintainability, and
   language observations unbound.
10. Mark fixtures `bound` only after all declared surfaces are covered; keep
    incomplete compile-equivalent work `partially-bound` with a note.
11. Run Excel-free contracts and the coverage report, then record rule codes
    and all executed case IDs in the pull request.

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
- [ ] Rejected fixtures contain expected diagnostics and `negative_controls`.
- [ ] Accepted controls forbid the bound rule codes on all declared surfaces.
- [ ] Rule codes:
      `<VBAxxx>`
- [ ] Rejected cases:
      `<case-id>`
- [ ] Accepted controls:
      `<control-id>`
```

The oracle remains local-only until an appropriate Excel-installed self-hosted
runner exists. No CI workflow may install, launch, or automate Excel for this
tool.
