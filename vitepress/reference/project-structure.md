# Project Structure

Typical xlflow projects contain:

```text
xlflow.toml
src/modules
src/modules/Xlflow
src/modules/Tests
src/classes
src/forms
src/forms/code
src/forms/specs
src/workbook
build
.xlflow
```

The `build` directory contains the managed workbook file, typically `build/<name>.xlsm`. Excel add-in and binary workbook projects use the same source layout and store `build/<name>.xlam` or `build/<name>.xlsb`.

Bundled `Xlflow*.bas` helpers are scaffolded under `src/modules/Xlflow/`; `Main.bas`, `App.bas`, and `Ui.bas` remain at the module root. Tests are recommended to live under `src/modules/Tests/` (or any subdirectory under `src/modules`) so they are naturally imported by `push` alongside production code.

`build` and `.xlflow` are generated state. Source-controlled VBA belongs under `src`.
