# Create your first xlflow project

This tutorial creates a harmless new workbook and takes it through one complete loop: create → inspect → run → save. Use it when you do not yet have a workbook to import.

**You will finish with:** a folder containing `xlflow.toml`, editable VBA under `src/`, and `build/Book.xlsm`, the workbook that Excel runs. Nothing outside this folder is modified.

## 1. Install and check Excel

Follow [Installation](../installation), then run:

```powershell
xlflow version
xlflow doctor --json
```

`doctor` checks the project-independent bridge, Excel, VBIDE access, and WSL boundary. Fix any error before opening a workbook-backed command.

## 2. Create the project

```powershell
mkdir first-xlflow-project
cd first-xlflow-project
xlflow new Book.xlsm
```

The command creates `build/Book.xlsm`, `xlflow.toml`, and the editable `src/` tree. The workbook is the persisted artifact; `src/` is the normal edit surface.

Open the folder now if you want to look around. You do not need to understand every generated file; `xlflow.toml` tells xlflow where the workbook is, while `src/` is the only place you normally edit VBA.

## 3. Inspect and run

```powershell
xlflow status --json
xlflow pull --json
xlflow lint --json
xlflow macros --json
xlflow run Main.Run --headless --json
```

Success means the command exits with code `0`, the JSON envelope has `status: "ok"`, and the macro result is visible in the output. If the macro name differs, use the `qualified_name` returned by `macros`.

## 4. Iterate with a session

```powershell
xlflow session start --json
xlflow push --fast --session --no-save --json
xlflow run Main.Run --headless --session --json
xlflow save --session --json
xlflow session stop --json
```

Use [Sessions and recovery](../guides/sessions) when a run times out or the session is dirty. Do not overwrite a newer workbook with `push` until `status` or a backup confirms which side is authoritative.

You have now used the basic vocabulary: **pull** gets workbook code into source, **push** puts source into a workbook, and **save** makes session work permanent. Next, either [edit this project in VS Code](./vscode-development) or learn the [existing-workbook path](./existing-workbook).
