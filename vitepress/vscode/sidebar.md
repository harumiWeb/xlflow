# Use the project sidebar

The sidebar groups project actions by intent. It is a shortcut for the CLI, not a separate source of truth: the status view tells you which copy is newer before an action changes it.

- **Project**: refresh, doctor, open workbook, and open documentation;
- **Modules/UserForms**: create, rename, delete, reveal source, and open procedures;
- **Execution**: pull, push, run, tests, save, and session start/stop;
- **Formulas**: pull and open formula snapshots;
- **Recovery**: inspect status, roll back, prune backups, and clear quarantine.

Actions operate on the configured project. Check the status view before destructive or source-of-truth-changing actions. In particular, choose **Pull** after edits in Excel/VBE, choose **Push** after edits in VS Code, and choose **Save** only when a live session has changes you mean to keep. If the status is uncertain, open the [state model](../concepts/workbook-session-source) instead of guessing.
