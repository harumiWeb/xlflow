# ADR-0019: Opt-in Workbook Coordination Waiting

## Status

Accepted

## Context

ADR-0018 makes immediate `workbook_busy` failure the safe default and provides
context-aware polling in the Windows workbook lock. Automation and AI-agent
workflows also need a bounded way to wait for short operations without retrying
the workbook command itself or treating an unbounded queue as a hung process.

Waiting affects every non-parallel-safe workbook command, its public CLI flags,
structured diagnostics, cancellation behavior, multi-workbook acquisition, and
WSL delegation. The timeout therefore needs one cross-command contract rather
than command-specific retry loops.

## Decision

xlflow exposes global `--wait` and `--wait-timeout <duration>` options. Waiting
is opt-in; without `--wait`, contention continues to fail immediately with
`workbook_busy`. `--wait` uses a finite 30-second timeout unless the caller
overrides it with a positive `--wait-timeout` value.

The timeout is one budget for acquiring every required workbook lock in stable
lock-ID order. It starts immediately before lock acquisition and ends once all
leases are held. It does not limit validation, the command body, Excel or VBA
execution, saves, or cleanup. A failure releases any earlier leases in reverse
order, and the workbook command body is never retried.

Waiting is allowed only when the authoritative command policy uses workbook
scope, is non-parallel-safe, and is retryable when busy. Unsupported commands
and invalid option combinations fail before Excel or the bridge starts.

The CLI initially attempts non-blocking acquisition. Only real contention
enters the polling wait and emits one human-readable waiting message. JSON mode
emits no progress text. Timeout and cancellation return stable
`workbook_busy_timeout` and `workbook_busy_cancelled` diagnostics with
operational exit code 3. Ctrl+C is intercepted only during lock acquisition;
normal signal behavior resumes after the leases are acquired.

WSL continues to delegate the complete invocation to Windows under ADR-0011,
including both wait options. The delegated Windows process owns the timeout and
lock lifecycle.

## Consequences

- Positive: automation can wait for short contention without duplicating retry
  policy or re-running a partially started command.
- Positive: every wait is finite by default and cancellation-aware.
- Positive: JSON output remains a single machine-readable envelope.
- Negative: polling does not guarantee FIFO fairness and a waiter may lose a
  handoff to another process.
- Negative: one slow earlier target consumes budget available to later targets
  in a multi-workbook operation.
- Negative: callers must distinguish immediate busy, wait timeout, and wait
  cancellation codes.

## Alternatives Considered

1. **Wait indefinitely when only `--wait` is provided** - Rejected because an
   executing macro can run without a bound and make automation appear hung.
2. **Require `--wait-timeout` with every `--wait`** - Rejected because a safe
   finite default gives interactive and agent callers a simpler common path.
3. **Retry the complete command after `workbook_busy`** - Rejected because only
   pre-execution lock acquisition is known to be safe to retry.
4. **Maintain a persistent FIFO queue** - Rejected because it requires a daemon
   and recovery protocol outside the initial coordination model.

## Related

- `docs/adr/ADR-0016-workbook-operation-coordination.md`
- `docs/adr/ADR-0018-windows-workbook-lock-and-owner-metadata.md`
- `docs/adr/ADR-0020-workbook-recovery-quarantine.md`
- `docs/specs/workbook-coordination.md`
- `docs/specs/cli-contract.md`
- xlflow issues #311 and #322
