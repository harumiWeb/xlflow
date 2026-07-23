# World News

A workbook that summarizes world news in Excel using NewsAPI-style data retrieval.

![World News](/images/world-news.png)

View the sample project on GitHub:
[example/world-news](https://github.com/harumiWeb/xlflow/tree/main/example/world-news)

<!-- xlflow-demo-case-study -->

## What it does

World News is a small workbook application that can be inspected, executed, and reviewed as source.

## Why it is a useful xlflow example

It demonstrates HTTP data retrieval, source-first modules, and structured workbook inspection.

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
