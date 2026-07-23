# Develop VBA in VS Code

Install [the VS Code extension](../vscode/installation), open an xlflow project, and use the editor as the source-first workspace. VS Code does not replace Excel: it is where you safely edit and understand VBA; Excel remains the place where the workbook runs.

**You need:** a project folder containing `xlflow.toml`. Create one with `xlflow new Book.xlsm` or import one with `xlflow init Existing.xlsm` before opening VS Code.

1. Open the folder containing `xlflow.toml`, not just an individual `.bas` file.
2. Run `xlflow pull` once when the workbook may be newer than `src/`.
3. Edit `.bas`, `.cls`, `.frm`, or UserForm sidecar files under `src/`.
4. Use completion, hover, definition navigation, and real-time diagnostics while editing.
5. Run lint/analyze from the command palette or terminal.
6. Use CodeLens to run a no-argument procedure or test.
7. Use the Testing view for discovered `Test*` procedures.
8. Push and save deliberately with the sidebar actions.

The first time, make a tiny edit and wait for a diagnostic or completion suggestion. Then choose **Push** only after `lint` is clean. If you use a live session, **Save session** is the final, explicit step that writes the live workbook back to disk. The [first VS Code project](../vscode/first-project) explains where those actions appear.

The extension delegates analysis and execution to the installed CLI; configure `xlflow.path` when the executable is not on `PATH`. See the [VS Code guide](../vscode/) for each feature and its limitations.
