# Run repeated commands in a live session

Use a managed session when you will edit and run more than once. It keeps one known Excel workbook open, which makes iteration faster and avoids wondering whether a command touched a different Excel instance.

Use a normal `push`/`run` for a one-off change. Use a session when you want a repeatable loop. The only extra rule is important: `--no-save` makes the **live** workbook newer than the file, so finish by saving or intentionally discard it.

```bash
xlflow session start --json
xlflow pull --session --json
xlflow push --fast --session --no-save --json
xlflow run Main.Run --diagnostic --session --json
xlflow inspect workbook --session --json
xlflow save --session --json
xlflow session stop --json
```

`--no-save` leaves the live workbook newer than disk. A dirty session requires `save --session` or an intentional discard. `session attach` adopts a workbook already opened by a user; stopping an external session detaches metadata and does not close Excel. See [recovery](../commands/recovery) when termination is uncertain.

**What success looks like:** `session start` reports a running session, commands with `--session` report the same workbook, and `save --session` clears the dirty/save-required state. If you accidentally stop before saving, do not guess—run `xlflow status --json` and use [recovery](../commands/recovery) when it asks you to.
