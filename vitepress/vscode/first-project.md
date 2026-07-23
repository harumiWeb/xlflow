# Open your first project

Create or import a project from the terminal, then open the directory containing `xlflow.toml`. The extension does not create a hidden second project: it reads the same source and workbook configuration as the CLI.

```powershell
xlflow new Book.xlsm
code .
```

The xlflow activity bar appears when the workspace is recognized. Refresh the project, open a module, run `xlflow: Pull` when the workbook is authoritative, and use `xlflow: Push` after source edits. The status view shows the workbook, source, session, and save-required state.

For a first verification, edit a comment or a clearly visible result cell, wait for the editor to show no errors, then use **Push**. If a procedure has a CodeLens **Run** action, run it and inspect the expected workbook result. A “save required” badge means you used a live session: choose **Save session** before closing Excel if you want to keep that run's changes.
