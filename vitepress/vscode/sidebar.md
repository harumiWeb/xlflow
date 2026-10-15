# Use the project sidebar

The sidebar groups project actions by intent. It is a shortcut for the CLI, not a separate source of truth: the status view tells you which copy is newer before an action changes it.

- **Project**: refresh, doctor, open workbook, and open documentation;
- **Modules/UserForms**: create, rename, delete, reveal source, open procedures,
  and build individual or all UserForms;
- **Execution**: pull, push, run, tests, save, and session start/stop;
- **Formulas**: pull and open formula snapshots;
- **Recovery**: inspect status, roll back, prune backups, and clear quarantine.

Actions operate on the configured project. Check the status view before destructive or source-of-truth-changing actions. In particular, choose **Pull** after edits in Excel/VBE, choose **Push** after edits in VS Code, and choose **Save** only when a live session has changes you mean to keep. If the status is uncertain, open the [state model](../concepts/workbook-session-source) instead of guessing.

Use a UserForm row's build button or context menu to build its FormSpec, or use
**Build All UserForms** in the view toolbar. Confirm the target workbook and forms
before existing workbook forms are overwritten. If target source or Designer
changes are pending, choose **Save and Build** to save those changes first.
Forms without a Spec cannot be built; duplicate Specs must be resolved first.

Builds run sequentially in one Excel session and save the workbook once after
all selected forms succeed. An existing session remains open. If a build fails,
later forms are not run and the session remains available with its pending
changes; the result lists completed, failed and remaining forms. The batch does
not roll back earlier successful builds.
