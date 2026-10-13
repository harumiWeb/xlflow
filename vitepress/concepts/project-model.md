# Project Model

An xlflow project is a normal source tree wrapped around a configured Excel workbook, add-in, or binary workbook. You can read it like any other project: configuration says what to use, source is what you edit, and generated folders contain results and local working state.

```text
xlflow.toml
src/modules/**/*.bas
src/classes/*.cls
src/forms/*.frm
src/forms/*.frx
src/forms/code/*.bas
src/forms/specs/*.yaml
src/workbook/*.cls
build/*.xlsm, build/*.xlam, or build/*.xlsb
.xlflow/
```

`xlflow.toml` defines the workbook path, source roots, lint behavior, and UserForm source mode. `src/` is the normal editing surface—commit it. `build/` contains the workbook, add-in, or binary workbook file—the artifact Excel opens. `.xlflow/` contains generated state, backups, and session metadata—inspect it for recovery, but do not hand-edit it.

For a first Git commit, include `xlflow.toml` and the relevant `src/` files. Keep the workbook according to your team's review and artifact policy; [Manage VBA source with Git](../guides/source-control) gives a practical baseline.

New and initialized projects default UserForm code authority to sidecar code files; `init` imports canonical Designer specs and code sidecars. Use `xlflow init --userform-code-source frm` explicitly when imported `.frm` artifacts must remain the code authority, including custom-control artifact workflows.
