# Backup and Rollback

The default `push` path is conservative. It backs up the workbook before replacing VBA components.

```bash
xlflow push --json
```

Fast development loops may use:

```bash
xlflow push --fast --session --no-save --json
```

That skips some safety cost for speed and leaves the live session dirty until `xlflow save --session`.

For review, use `diff` to compare workbook files and optional exported VBA trees:

```bash
xlflow diff before.xlsm after.xlsm --vba-before before-src --vba-after after-src --json
```
