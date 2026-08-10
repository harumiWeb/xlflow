# ADR-0033: Command-Execution Safety Ownership and Risk Taxonomy

## Status

Accepted

## Context

VBA can launch operating-system processes through several syntactic shapes:
the `Shell` statement, `WScript.Shell.Run`/`Exec`, and Win32 shell APIs. A
single command may also embed another interpreter such as `cmd.exe /c` or
PowerShell. External worksheet, file, user, and environment values can reach
the executable path, ordinary arguments, or interpreter command text. A
diagnostic that only reports “tainted data reaches an API” cannot distinguish
shell injection from a benign fixed argument, an unquoted path, an exposed
secret, or a launch whose result is ignored.

ADR-0025 established a conservative procedure-local source-to-sink data-flow
layer and exposed it through `VBA224`. That generic rule is useful for many
sinks, but process execution needs sink-specific argument-role and Windows
quoting knowledge. Keeping process findings under `VBA224` would either lose
that distinction or duplicate diagnostics when a specialized rule is added.

Issue #466 also requires a stable machine-readable risk taxonomy and the same
behavior in batch and LSP surfaces. The change is therefore a public diagnostic,
configuration, registry, documentation, and ownership decision rather than a
private implementation detail.

## Decision

Add default-enabled, warning-level, non-blocking `VBA236`,
`detect_unsafe_command_construction`, as the owner of process-launch safety
findings. It is procedure-local, available in batch and real-time analysis, and
inline suppressible. The registry is authoritative for its public identity,
configuration, severity, surfaces, and suppression policy.

Reuse ADR-0025's data-flow states and propagation paths, but classify each
recognized sink by launcher and argument role: executable path, ordinary
arguments, shell command text, URL/document target, window style, wait flag,
and result observation. The v1 sink catalog includes VBA `Shell`,
`WScript.Shell.Run`/`Exec`, `Shell.Application.ShellExecute`, and Win32
`ShellExecute` variants. `FollowHyperlink` and `ShellExecuteEx` remain out of
scope until their navigation/structure semantics are modeled.

The finding context uses an additive `command_execution` object containing
`risk_class`, `risk_kind`, `launcher`, optional `interpreter`,
`command_role`, and `origin_state`. Existing `data_flow` source/sink/path
context is retained for known origins. Risk classes are `injection` and
`process_launch`; stable risk kinds are `tainted_command_text`,
`unknown_origin`, `credential_exposure`, `observability`, and
`unquoted_executable_path`:

- `injection` is used only with `tainted_command_text` when known tainted input
  reaches executable or interpreter command text;
- `process_launch` is used with `unknown_origin` for uncertain origins or
  ordinary arguments, and with the independent path, credential, or
  observability kinds when those policy observations apply.

Only taint with a known source and role may be called a potential injection
risk. Unknown transformations and origins remain general process-launch risk;
the analyzer must not claim a confirmed injection vulnerability. Windows
guidance documents executable quoting, nested `cmd /c` metacharacters,
PowerShell `-File` over `-Command`, fixed script-host paths, and moving secrets
off command lines.

Transfer process-launch sink ownership from `VBA224` to `VBA236`. `VBA224`
continues to report generic source-to-sink flows for non-process APIs. The two
rules must not emit duplicate findings for the same process-launch call.
When `VBA236` is disabled, `VBA224` remains the compatibility fallback for
process-launch flows. Existing `VBA224` suppressions do not suppress the new
`VBA236` findings, so projects that intentionally exempt a process launch must
migrate the suppression or disable/suppress both rule IDs during the transition.

## Consequences

- Process findings carry enough role and risk metadata for CLI, LSP, and agent
  consumers to distinguish injection from general launch hygiene.
- Constant commands remain accepted for injection while independent path,
  credential, and observability risks stay visible.
- Shared data-flow traversal remains reusable, and process-specific policy does
  not leak into the generic source catalog.
- Existing users of `VBA224` may see process findings move to `VBA236`; the
  stable configuration and suppression mechanisms continue to work through the
  new rule ID and key.
- Conservative unknown-state handling intentionally leaves some launches as
  general process-launch warnings until richer origin or sanitizer summaries
  exist.
- Windows quoting and interpreter behavior add a platform-specific policy
  surface that must be kept in the command-execution safety specification and
  registry-generated reference in sync.

## Alternatives Considered

1. **Keep all process launches under `VBA224`.** Rejected because the generic
   sink cannot reliably distinguish executable paths, shell text, URL targets,
   and ordinary arguments without making its contract process-specific.
2. **Report every dynamic command as injection.** Rejected because unknown
   origin or role is not proof of attacker-controlled input and would violate
   the conservative evidence contract.
3. **Use a separate rule per launcher or interpreter.** Rejected for v1
   because launcher-specific IDs would fragment configuration and duplicate the
   same data-flow and suppression behavior; launcher/interpreter names remain
   structured context instead.
4. **Treat quoting helpers as universal sanitizers.** Rejected because Windows
   quoting is role- and interpreter-dependent, and user-defined helpers lack a
   verified summary contract.

## Evidence

- Issue #466 requirements for command sinks, risk distinctions, quoting
  guidance, and non-confirmatory unknown-origin behavior.
- Shared data-flow contract: `docs/specs/vba-source-sink-dataflow.md` and
  `docs/adr/ADR-0025-conservative-vba-source-sink-dataflow.md`.
- Registry ownership and generated projections:
  `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`,
  `internal/staticanalysis/rules/registry.json`, and
  `scripts/docs/generate-reference-inventory.mjs`.
- Public configuration and diagnostic contracts:
  `docs/specs/cli-contract.md` and
  `docs/specs/vba-command-execution-safety.md`.

## Supersedes

- None

## Superseded by

- None

## Related

- Issue #466
- `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`
- `docs/adr/ADR-0025-conservative-vba-source-sink-dataflow.md`
- `docs/specs/vba-source-sink-dataflow.md`
- `docs/specs/vba-command-execution-safety.md`
