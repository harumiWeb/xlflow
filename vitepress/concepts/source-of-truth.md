# Source of Truth

“Source of truth” simply means the copy whose changes you intend to keep. For normal development, that is source files under `src/`: edit them first, check them, then run `push` to put them into Excel.

The workbook can become newer than source when a user edits in Excel, when a session is dirty, or when a command intentionally mutates workbook state. In that case, run `pull` before editing source.

Use `xlflow status` to quickly check whether source, workbook, and session are in sync. This is safe to run before any decision:

```bash
xlflow status
```

If `src_newer_than_workbook` is `true`, source has an unpushed edit: run `xlflow push` when you want Excel to receive it. If the session is dirty, run `xlflow save --session` when you want the live result to survive. If Excel/VBE is newer, run `pull` before editing source again.

UserForms have two tracked concerns:

- Designer structure: `src/forms/specs/*.yaml`
- Code-behind: `src/forms/code/*.bas` in sidecar mode, or embedded `.frm` code in compatibility mode

Imported projects default to compatibility `frm` mode. Run `xlflow form migrate sidecar` when you want `src/forms/specs/*.yaml` and `src/forms/code/*.bas` to become the primary editing surface for existing UserForms.

Do not treat `.frm` / `.frx` files as the primary Designer source in sidecar mode. `.frx` files are binary companion artifacts, and `.frm` files remain generated artifacts used by xlflow workflows.
