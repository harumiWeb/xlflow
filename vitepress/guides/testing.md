# Test and analyze VBA

Use each layer to answer one different question. Passing `lint` does not prove a macro produces the right cell value; passing a test does not mean the source is formatted. Run cheap source checks before commands that open Excel.

| Tool            | Answers                                            |
| --------------- | -------------------------------------------------- |
| `fmt --check`   | Is source formatted?                               |
| `lint`          | Is the VBA structurally safe and style-compliant?  |
| `analyze`       | Are runtime-risk patterns present?                 |
| LSP diagnostics | What does the current unsaved editor buffer imply? |
| `test`          | Does executable VBA behavior satisfy the test?     |

```bash
xlflow fmt --check
xlflow lint --json
xlflow analyze --json
xlflow test --json
```

Keep source-only checks before Excel-backed checks. Use `--session` for fast iteration and `--no-save` when the test is exploratory. After a behavior test, use `inspect` or `export-image` for the one workbook result a human would recognise; then `save --session` only when that result should persist.
