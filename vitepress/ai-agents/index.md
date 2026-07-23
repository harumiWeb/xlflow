# AI Agents

xlflow gives coding agents a stable interface for Excel VBA projects: source files, JSON output, explicit exit codes, diagnostics, sessions, and visual inspection commands. An agent does not need to click through Excel to make progress; it can change source, run a controlled command, and inspect an observable result.

The safety boundary matters: agents edit `src/`, `push` places that source into Excel, and `save --session` is the explicit point where a live experiment becomes the workbook file on disk.

Every command can return a stable JSON envelope, so coding agents can parse results without scraping terminal text. See [JSON Output](../reference/json-output) and [Error Codes](../reference/error-codes).

Before editing, confirm workbook and source state. This command changes nothing and prevents an agent from overwriting VBE work by accident:

```bash
xlflow status --json
```

Recommended loop (replace `Main.Run` with a name returned by `macros`):

```bash
xlflow doctor --json
xlflow status --json
xlflow session start --json
xlflow pull --session --json
xlflow push --fast --session --no-save --json
xlflow lint --json
xlflow analyze --json
xlflow test --session --json
xlflow macros --session --json
xlflow run Main.Run --headless --session --json
xlflow save --session --json
```

When the agent runs inside WSL, keep the repository under `/mnt/<drive>/...` and install xlflow on both WSL and Windows. The same commands transparently delegate Excel operations to Windows while `lint`, `fmt`, and `analyze` remain in WSL. Prefer the session-first sequence:

```bash
xlflow doctor --json
xlflow session start --json
xlflow push --fast --session --no-save --json
xlflow run Main.Run --diagnostic --session --json
xlflow inspect workbook --session --json
xlflow save --session --json
xlflow session stop --json
```

Read the loop in four phases: **understand** (`doctor`, `status`, `pull`), **edit and check** (`lint`, `analyze`), **prove** (`push`, `test`, `run`, `inspect`), then **persist** (`save`, `stop`). `--no-save` makes the proof reversible until the agent has inspected the result.

If a run fails, read `error.code` first. A source error normally means fix source and repeat the checks; a `workbook_recovery_required` error means stop the loop. Do not blindly repeat an Excel command. For ordinary failures, check state and recover with a short loop:

```bash
xlflow status --json
xlflow lint --json
xlflow analyze --json
xlflow run Main.Run --diagnostic --session --json
xlflow inspect workbook --json
xlflow pull --session --json
```

Use `run --diagnostic` when you need structured compile or runtime failure details instead of a blind Excel/VBE dialog. Re-run `pull` after workbook-side investigation when you need the latest saved workbook state back in source.

## Recovery-required workflow

If a result contains `error.code: "workbook_recovery_required"` or top-level
`recovery.required: true`, stop the normal loop. Do not retry with `--wait` and
do not save the quarantined session.

```bash
xlflow status --json
xlflow session stop --discard --json
# or, when an affected PID is reported:
xlflow process cleanup <pid> --json
xlflow recovery clear --json
xlflow status --json
```

For external user-owned Excel sessions, do not close or discard the workbook
automatically. Ask the user to close it without saving, then run verified
`recovery clear`. Use `recovery clear --force` only with explicit acceptance
that it clears the marker without stopping VBA or proving workbook safety.

Install the bundled skill when supported by your agent:

```bash
xlflow skill install --agent codex
```
