# Quickstart: one complete workflow

This walkthrough turns a new workbook into a checked, executed result. For an existing `.xlsm`, use [Import an existing workbook](./tutorials/existing-workbook).

## 1. Create and check the project

```powershell
xlflow new Book.xlsm
xlflow doctor --json
xlflow status --json
```

`new` creates `build/Book.xlsm`, `xlflow.toml`, and the editable `src/` tree. `doctor` confirms that Excel and VBIDE can be automated; `status` records the initial source/workbook/session state.

## 2. Export and check VBA

```powershell
xlflow pull --json
xlflow fmt --check --json
xlflow lint --json
xlflow analyze --json
```

`pull` makes the workbook's VBA visible as source files. The remaining commands are source-only checks, so fix their findings before opening Excel.

## 3. Push, run, and verify

```powershell
xlflow push --json
xlflow macros --json
xlflow run Main.Run --diagnostic --json
xlflow inspect workbook --json
```

Use the `qualified_name` returned by `macros` when the scaffolded entry point differs. A successful run returns exit code `0` and a JSON envelope with `status: "ok"`.

## 4. Iterate without reopening Excel

```powershell
xlflow session start --json
xlflow push --fast --session --no-save --json
xlflow test --session --json
xlflow run Main.Run --diagnostic --session --json
xlflow save --session --json
xlflow session stop --json
```

`--no-save` keeps the live workbook newer than disk until the explicit save. If a step fails, inspect `error.code` and follow [Troubleshooting](./help/troubleshooting); do not retry a recovery quarantine with `--wait`.

## Command cheat sheet

```text
source edit -> fmt/lint/analyze -> push -> run/test -> inspect -> save --session
workbook edit in VBE -> pull -> review git diff
uncertain state -> status -> backup -> choose pull or push
```
