# Source of Truth

For normal development, source files under `src/` are the authority. Agents should edit files first, then run `push`, `lint`, `test`, and `run`.

The workbook can become newer than source when a user edits in Excel, when a session is dirty, or when a command intentionally mutates workbook state. In that case, run `pull` before editing source.

UserForms have two tracked concerns:

- Designer structure: `src/forms/specs/*.yaml`
- Code-behind: `src/forms/code/*.bas` in sidecar mode, or embedded `.frm` code in compatibility mode

Do not treat `.frx` files as reviewable text. They are binary companion artifacts.
