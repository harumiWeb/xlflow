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

After `xlflow new Book.xlsm`, the project contains:

- `build/Book.xlsm` as the workbook managed by xlflow
- `xlflow.toml` as the project config
- `src/` as the exported VBA source tree you edit

Edit files under `src/`, then import those changes back into the workbook with:

```bash
xlflow push --json
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

Use `--json` for scripts and agents. Every command returns a stable JSON envelope, and failures use stable error codes so callers do not need to scrape terminal text. See [JSON Output](./reference/json-output) and [Error Codes](./reference/error-codes).
