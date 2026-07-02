# xlflow type db

Manage generated TypeLib databases used by the VBA LSP.

## Usage

```bash
xlflow type db status
xlflow type db init
xlflow type db refresh
xlflow type db refresh --library all
xlflow type db clean
```

## Options

| Option      | Description                                          | Default             |
| ----------- | ---------------------------------------------------- | ------------------- |
| `--dir`     | Override the generated type DB directory.            | `~/.xlflow/typelib` |
| `--library` | TypeLib library to import. Repeat, comma-separate, or use `all` for every known library present on this machine. | `excel`             |
| `--force`   | Deprecated compatibility flag; `refresh` always regenerates. | `false`             |

`status` reports manifest presence, generated files, library LIBID/version metadata, stale state, and the LSP database search order.

`init` generates the database only when it does not already exist. `refresh` always regenerates the generated database, so it is the one-command equivalent of `clean` followed by `init`. `clean` deletes the generated database directory.

The default importer target is Excel. Use `--library all` to generate databases for every known library present on the machine, including Office, MSForms, Scripting, ADODB, and VBIDE when available. Generated entries include TypeLib-derived ProgID mappings when registry metadata can be matched to TypeLib CoClass GUIDs, so the LSP can infer common late-bound expressions such as `CreateObject("Excel.Application")`, `CreateObject("Scripting.FileSystemObject")`, and `CreateObject("ADODB.Connection")`.

Generated entries are loaded by `xlflow lsp` when present; if no generated DB exists, the LSP continues with the embedded built-in database.
