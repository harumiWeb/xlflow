# Push/Run Keepalive Spec

## Goal

Prevent AI agents and task runners from treating long-running Excel COM operations as stalled when `xlflow push` or `xlflow run` produces no final result for several seconds.

## CLI Contract

- `xlflow push --keepalive [--keepalive-interval <duration>]`
- `xlflow run [macro] --keepalive [--keepalive-interval <duration>]`
- `--keepalive-interval` defaults to `5s`.
- When `--keepalive` is enabled, a non-positive interval is a CLI argument error with exit code `2`.
- Keepalive output is written only to stderr.
- Stdout remains unchanged, including pure JSON output when `--json` is set.

## Output Contract

Heartbeat lines start immediately and repeat at the configured interval:

```text
xlflow: push still running... elapsed=0s
xlflow: run still running... elapsed=5s
```

Completion markers are written after the command result is known:

```text
XLFLOW_DONE status=success command=push
XLFLOW_DONE status=failed command=run code=macro_timeout
```

`code` is present only when xlflow has a structured error code.

## Agent Rule

Agents should use `--keepalive --json` for long `push` and `run` calls, wait for process exit, and not start the next workbook-dependent command until stderr contains `XLFLOW_DONE`.
