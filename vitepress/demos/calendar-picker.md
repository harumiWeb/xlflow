# Calendar Picker

A richer UserForm interface that benefits from Designer snapshot and export-image checks.

![ Calendar Picker ](/images/calendar-picker.png)

View the sample project on GitHub:
[example/calendar-pick](https://github.com/harumiWeb/xlflow/tree/main/example/calendar-pick)

<!-- xlflow-demo-case-study -->

## What it does

Calendar Picker is a small workbook application that can be inspected, executed, and reviewed as source.

## Why it is a useful xlflow example

It demonstrates UserForm Designer artifacts, sidecar code-behind, and controlled form snapshots.

## Project structure

The repository keeps VBA under `src/`, the workbook under `build/`, and project behavior in `xlflow.toml`.

## xlflow features used

- `doctor`, `status`, and `pull` for setup and source synchronization;
- `fmt`, `lint`, and `analyze` before Excel operations;
- sessions, `run --diagnostic`, `inspect`, and `export-image` for verification.

## Verification strategy

Run the source checks, push into a disposable workbook, execute the documented entry point, inspect the affected cells or form, export an image when layout matters, and review the Git diff.

## Commands to reproduce

```bash
xlflow doctor --json
xlflow pull --json
xlflow lint --json
xlflow push --json
xlflow run --diagnostic --json
xlflow inspect workbook --json
xlflow export-image --json
```

## Repository

See the linked example repository above for the workbook and source files.
