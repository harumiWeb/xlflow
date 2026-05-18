# Config File

xlflow reads `xlflow.toml` from the project root.

```toml
[project]
name = "sample"
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"
visible = false
display_alerts = false

[src]
modules = "src/modules"
classes = "src/classes"
forms = "src/forms"
workbook = "src/workbook"

[vba]
folders = true
folder_annotation = "update"
default_component_folders = true

[userform]
code_source = "sidecar"
```

See `docs/specs/cli-contract.md` for the stable contract.
