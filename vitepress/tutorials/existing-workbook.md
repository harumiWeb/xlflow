# Import an existing macro-enabled workbook

This is the common adoption path for a workbook that already contains VBA. It creates a project copy, makes the VBA reviewable as files, changes one small thing, and proves that Excel used the change.

**Before you begin:** keep your original `.xlsm` somewhere safe and close it in Excel. `init` does not modify that original—it copies it into the project—but closing it avoids uncertainty about which workbook you are looking at.

## 1. Prepare and diagnose

Close the workbook in Excel, make a copy, and run:

```powershell
xlflow doctor --json
xlflow init C:\path\to\Existing.xlsm
```

`init` copies the workbook into `build/` without changing its filename or format and writes `xlflow.toml`. It does not replace the original file.

**Expected result:** your current folder now contains `xlflow.toml` and `build/Existing.xlsm`. From this point on, use the copy under `build/` for xlflow work; leave the original as your fallback.

## 2. Inspect the project

```powershell
xlflow status --json
xlflow pull --json
```

Review `src/` in VS Code. Standard modules are source files such as `src/modules/Main.bas`; document modules and UserForms retain their tracked artifacts. Commit this baseline before editing.

At this point no VBA behavior has changed. `pull` only exported the VBA from the project copy into `src/`, so `git diff` gives you a trustworthy baseline before you edit anything.

## 3. Edit and check source

Edit one procedure, then run:

```powershell
xlflow fmt --check
xlflow lint --json
xlflow analyze --json
```

These commands do not require Excel. A lint or preflight error should be fixed in source before `push` opens Excel.

Choose a change you can recognise—for example, alter a constant or write a clear value to a result cell. This makes the later verification meaningful. Do not edit the same procedure in VBE and VS Code at the same time: if you do edit in VBE, stop and `pull` first.

## 4. Push, run, and verify

```powershell
xlflow push --json
xlflow macros --json
xlflow run Main.Run --diagnostic --json
xlflow inspect workbook --json
xlflow export-image --sheet Result --range A1:F20 --json
```

`push` imports source, compiles VBA, and saves by default. `run --diagnostic` turns compile/runtime failures into structured output instead of leaving an agent behind a modal dialog. `inspect` and `export-image` verify workbook state without relying on visual memory.

**How to tell it worked:** `push` and `run` return `"status": "ok"`; `inspect` shows the expected cell or object state; the image gives you a visual record. If the macro is not named `Main.Run`, copy its `qualified_name` from `xlflow macros --json`.

## 5. Review and recover

```powershell
git diff -- src xlflow.toml
xlflow diff <before-workbook> <after-workbook> --json
```

If the workbook was edited in VBE, run `pull` first. If a session is dirty, run `save --session` before stopping it. For uncertain termination, follow [recovery](../commands/recovery) and [Troubleshooting](../help/troubleshooting).

The important habit is simple: review source changes in Git, then treat `push` as the deliberate moment those changes enter Excel.
