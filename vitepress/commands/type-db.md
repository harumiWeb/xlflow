# xlflow type db

Manage generated TypeLib databases used by the VBA LSP.

## Usage

```bash
xlflow type db status
xlflow type db init
xlflow type db refresh
xlflow type db refresh --force
xlflow type db clean
```

## Options

| Option      | Description                                          | Default             |
| ----------- | ---------------------------------------------------- | ------------------- |
| `--dir`     | Override the generated type DB directory.            | `~/.xlflow/typelib` |
| `--library` | TypeLib library to import. Repeat or comma-separate. | `excel`             |
| `--force`   | Regenerate during `refresh` even when current.       | `false`             |

`status` reports manifest presence, generated files, library LIBID/version metadata, stale state, and the LSP database search order.

`init` generates the database only when it does not already exist. `refresh` regenerates when stale, and `clean` deletes the generated database directory.

The initial importer target is Excel. Generated entries are loaded by `xlflow lsp` when present; if no generated DB exists, the LSP continues with the embedded built-in database.
