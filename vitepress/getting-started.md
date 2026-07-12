# Getting Started

xlflow turns an Excel VBA workbook, add-in, or binary workbook into a source-controlled project that can be edited, checked, and executed from the command line.

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

Install the bundled agent skill during scaffolding when you want an AI coding agent to follow xlflow workflows:

```bash
xlflow new Book.xlsm --with-skill --agent codex
```

## First workflow

```bash
xlflow doctor --json
xlflow pull --json
xlflow formulas pull --json
xlflow lint --json
xlflow macros --json
xlflow run Main.Run --headless --json
```

Use `xlflow session start` for repeated edit loops so Excel and the configured workbook stay open between commands.

## Next pages

- [Installation](./installation)
- [Quickstart](./quickstart)
- [Command reference](./commands/)
- [AI agent workflow](./ai-agents/)
