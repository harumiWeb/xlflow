# Workbook Recovery Reference

Load this reference whenever a command returns a recovery-required error or
top-level `recovery.required=true`.

## Fail Closed

Recovery means xlflow cannot safely determine the outcome of a prior workbook
operation. Stop normal workbook work immediately.

- Do not retry the failed operation, add a wait option, or start another
  workbook-bound command.
- Do not save the quarantined workbook or assume that a free workbook lock makes
  it safe.
- Read the structured recovery metadata and follow its suggested actions.

## Recovery Workflow

1. Run `xlflow status --json` and inspect the recovery state, reason, prior
   operation, and recorded Excel PID when present.
2. For an xlflow-managed session, run `xlflow session stop --discard --json`.
   This discards uncertain unsaved state; never replace it with a save.
3. If the affected Excel PID remains, run `xlflow process cleanup <pid> --json`
   and confirm that the result reports termination before clearing recovery.
4. For an external user-owned session, do not close, discard, or save the
   workbook automatically. Ask the user to close it in Excel without saving.
5. After the recorded process has stopped, run `xlflow recovery clear --json`.
   Leave the marker in place when verification fails.
6. Use `xlflow recovery clear --force --json` only after manual Excel recovery is complete and the user explicitly accepts that it removes xlflow's marker without stopping VBA, closing Excel, discarding changes, or proving workbook safety.
7. Run `xlflow status --json` again. Resume normal workbook work only when the
   recovery-required state is clear.

## Process Cleanup Boundaries

Use process cleanup only for the affected process or an explicitly selected safe
cleanup mode. `xlflow process cleanup --all` can terminate every local Excel
process, including workbooks with unsaved work; require explicit user direction
before using it.

If recovery metadata cannot be published or checked, continue to treat the
workbook as unsafe. Stop or have the user close Excel manually before any further
workbook operation.

## After Recovery

Re-orient after recovery: establish source authority again, start or attach a
fresh session as appropriate, and repeat the normal proof loop. Do not assume
that source, disk, or live workbook state survived the interrupted operation.
