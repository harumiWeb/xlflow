# Invariants

## Public Contracts

- Public CLI behavior, JSON schemas, exit codes, and documented flags must remain backward compatible unless the task explicitly changes them.
- YAML and TOML schema compatibility must be preserved or documented with migration guidance.
- Human output must distinguish a real empty result from data unavailable due to failure.
- Persisted spec/review artifacts must not absorb transient operational warnings such as `save_required` unless that is their explicit contract.

## Workbook and Session Safety

- Workbook-backed commands must act on the configured workbook or the matching live session workbook, not an accidental hidden copy.
- Dirty session state and save-required warnings must remain visible on success and failure paths.
- Execution commands must keep macros executable; inspection commands may disable automation macros.
- `push`, `run`, `test`, `pull`, and `save` multi-command validation should prefer session-first workflows.

## Source Fidelity

- Exported VBA must be deterministic.
- AST transforms and source normalizations must be reversible or conservatively preserve executable code.
- Document modules must not retain exported `VERSION/BEGIN/MultiUse/END` headers when reinserted into VBIDE.
- Malformed or partially missing headers must not trigger aggressive truncation of body code.
- UserForm snapshot/build round-trips must preserve designer dimensions without artificial offsets.

## Preflight and Error Safety

- Predictable source problems should fail with structured CLI issues before Excel opens.
- Missing xlflow trace helpers are compile-time source problems except where configured `run --trace` can inject helpers temporarily.
- Known Excel object/member mismatches and invalid VBA quote escapes should be caught before VBE modal dialogs.
- Unsupported CLI/script action enums should fail before Excel COM starts.

## Cleanup

- Temporary import/export workspaces must be removed after command completion.
- Temporary VBIDE helpers must not leave persistent workbook state changed unless that is the requested mutation.
- File handles opened by generated VBA must be closed on success and failure.
