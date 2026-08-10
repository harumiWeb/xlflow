# VBA Command Execution Safety

<!-- xlflow-rule-contract: {"id":"VBA236","family":"analyze","category":"security","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_unsafe_command_construction","inline_suppressible":true,"preflight_blocking":false} -->

`VBA236` is a default-enabled, procedure-local warning for process-launch
commands assembled from worksheet, file, user, or other external input. It is
available in batch analysis and real-time LSP diagnostics, is inline
suppressible, and never blocks source preflight. Disable it project-wide with
`[analyze].disabled_rules = ["VBA236"]`; the legacy
`detect_unsafe_command_construction` Boolean remains accepted with the normal
deprecation and disabled-rules precedence behavior.

## Scope and sink recognition

The rule classifies calls through the shared ProcedureIR/data-flow layer. It
recognizes the VBA `Shell` statement and these Windows process-launch APIs:

- `WScript.Shell.Run` and `WScript.Shell.Exec`, including a receiver alias
  established by `CreateObject("WScript.Shell")`;
- `Shell.Application.ShellExecute`;
- Win32 `ShellExecute`, `ShellExecuteA`, and `ShellExecuteW` declarations.

`FollowHyperlink` and `ShellExecuteEx` are outside the v1 contract. The former
needs document-navigation semantics and the latter requires structure tracking
that is not available in the procedure-local IR.

The classifier separates the executable path, ordinary arguments, shell
command text, URL/document target, window style, wait flag, and result value
where the call shape makes those roles knowable. A user-defined procedure named
`Shell`, or an unrelated object member named `Run`/`Exec`, is not treated as a
process sink.

When a command starts `cmd.exe /c`, `powershell.exe`/`pwsh`, or
`wscript.exe`/`cscript.exe`/`mshta.exe`, the interpreter and its command-text
argument are retained in the finding context. This distinguishes shell-text
injection from an ordinary executable argument.

## Risk classes

Findings use an additive `command_execution` JSON context with
`risk_class`, `risk_kind`, `launcher`, optional `interpreter`,
`command_role`, and `origin_state`. Existing `data_flow` source/sink/path
context is preserved when a known source reaches the sink.

The rule reports two `risk_class` values:

- `injection` class: a known tainted value is placed in the executable path,
  `cmd /c` text, PowerShell `-Command` text, or script-host command text;
- `process_launch` class: a launch occurs with an unknown origin or role, an
  external value chooses a URL/document target, or an external value is passed
  as an ordinary process argument;

The stable `risk_kind` values are:

- `tainted_command_text`: the `injection` risk kind for a known tainted command
  or executable role;
- `unknown_origin`: the `process_launch` risk kind when origin or role is not
  precise enough to claim injection;
- `credential_exposure`: a credential-like literal or known secret binding is
  placed in command-line arguments;
- `observability`: a constant hidden window style is used and the process ID
  or `Run` exit code is discarded; and
- `unquoted_executable_path`: a statically reconstructed executable path
  contains spaces without an outer pair of quotes.

Only `StateTainted` with a known source origin and command role is described as
`potential command-injection risk`. Unknown transformations, unresolved
origins, and `StateUnknown` are reported as general process-launch risk; the
rule never claims a confirmed injection vulnerability from an unknown input.

Constant commands, quoted executable paths, and fixed arguments are accepted
for injection purposes. Independent findings still apply to unquoted paths,
credential exposure, and hidden or unobserved launches in otherwise constant
commands.

## Windows guidance

Suggestions should quote the complete executable path before appending fixed
arguments. For `cmd.exe /c`, avoid shell text when possible; if it is required,
quote nested executable and argument text and treat `&`, `|`, `<`, `>`, `(`,
`)`, `^`, `%`, and `!` as shell metacharacters. For PowerShell, prefer
`-File` with fixed arguments over `-Command`; do not interpolate external input
into a script string. Script hosts should receive a fixed script path and
validated arguments. Keep secrets out of command lines by using environment
variables, standard input, or a protected store instead.

## Ownership and compatibility

`VBA236` owns process-launch-specific findings. `VBA224` remains the generic
procedure-local source-to-sink data-flow rule for SQL, file, workbook, HTTP,
and other non-process sinks; it must not duplicate a process-launch diagnostic.
The shared data-flow states and path format remain defined in
[VBA source-to-sink dataflow](vba-source-sink-dataflow.md).

The finding is policy evidence, warning severity, and non-blocking. The rule's
metadata, configuration key, LSP eligibility, and suppression behavior are
authoritative in the shared static-analysis registry.
