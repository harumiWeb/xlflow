# Quickstart

This loop is the smallest useful xlflow workflow.

```bash
xlflow new Book.xlsm
xlflow doctor --json
xlflow pull --json
xlflow lint --json
xlflow macros --json
xlflow run Main.Run --headless --json
```

For iterative development, keep a live Excel session open:

```bash
xlflow session start --json
xlflow pull --session --json
xlflow push --fast --session --no-save --json
xlflow test --session --json
xlflow run Main.Run --headless --session --json
xlflow save --session --json
xlflow session stop --json
```

Use `--json` for scripts and agents. Human output is intentionally optimized for terminal reading and should not be parsed as a stable API.
