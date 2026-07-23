# Getting Started

xlflow turns an Excel VBA workbook into a project you can understand and change outside the VBA editor. You edit ordinary files in `src/`, check them before Excel opens, then deliberately copy them into the workbook when you are ready.

If that is all new to you, start with one idea: **the source folder is where you work; the workbook is what Excel runs.** `pull` copies VBA out of Excel, and `push` copies your edited source back in. A live session is only a faster, still-unsaved copy of the workbook.

## Requirements

- Windows
- Microsoft Excel
- PowerShell
- Trust access to the VBA project object model enabled in Excel

Commands that only inspect source files or saved workbook files, such as `lint`, `analyze`, and `formulas pull`, can run without Excel. Workbook-backed automation commands use Excel COM.

## Create a project

Create a new macro-enabled workbook:

```bash
xlflow new Book.xlsm
```

Omit the extension to default to `.xlsm`, or use an explicit `.xlam` filename for an Excel add-in project or `.xlsb` for an Excel Binary Workbook VBA project.

Or initialize from an existing workbook or add-in:

```bash
xlflow init Book.xlsm
xlflow init ExistingAddin.xlam
xlflow init ExistingModel.xlsb
```

`new` is for a blank workbook. `init` is for a workbook you already have; it copies that workbook into the project so the original is left alone. If your goal is to put a real `.xlsm` under Git, use the [existing-workbook tutorial](./tutorials/existing-workbook) rather than continuing with the abbreviated commands below.

Install the bundled agent skill during scaffolding when you want an AI coding agent to follow xlflow workflows:

```bash
xlflow new Book.xlsm --with-skill --agent codex
```

## Choose a workflow

Use [Choose your workflow](./choose-workflow) to select a complete path. The most common adoption path is [Import an existing workbook](./tutorials/existing-workbook); new projects can start with [First xlflow project](./tutorials/first-project).

## First workflow

```bash
xlflow doctor --json
xlflow pull --json
xlflow formulas pull --json
xlflow lint --json
xlflow macros --json
xlflow run Main.Run --headless --json
```

Read the commands as a story, not a required incantation: `doctor` proves the Excel setup, `pull` exposes workbook VBA as files, `lint` catches source problems, `macros` tells you a safe runnable name, and `run` executes it. A successful JSON response has `"status": "ok"`; a failed one tells you where to look through `error.code`.

Once you begin changing VBA repeatedly, use `xlflow session start` so Excel and the workbook stay open between commands. The [quickstart](./quickstart) shows the whole edit → verify → save loop.

## Next pages

- [Installation](./installation)
- [Quickstart](./quickstart)
- [Choose your workflow](./choose-workflow)
- [Tutorials](./tutorials/)
- [VS Code](./vscode/)
- [Troubleshooting](./help/troubleshooting)
- [Command reference](./commands/)
- [AI agent workflow](./ai-agents/)
