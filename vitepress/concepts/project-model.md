# Project Model

An xlflow project is a normal source tree wrapped around a configured Excel workbook.

```text
xlflow.toml
src/modules/*.bas
src/classes/*.cls
src/forms/*.frm
src/forms/*.frx
src/forms/code/*.bas
src/forms/specs/*.yaml
src/workbook/*.cls
build/*.xlsm
.xlflow/
```

`xlflow.toml` defines the workbook path, source roots, lint behavior, and UserForm source mode. `src/` is the normal editing surface. `build/` contains the workbook file, and `.xlflow/` contains generated state, workbook-file backups with metadata, and session metadata.

New projects default UserForm code authority to sidecar code files. Initialized projects preserve the imported `.frm` code model until the user intentionally migrates.
