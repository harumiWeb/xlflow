# xlflow new

Create a new xlflow project and macro-enabled workbook or Excel add-in.

## Usage

```bash
xlflow new [workbook] [--with-skill] [--agent <provider>] [--no-update-check]
```

## Options and Arguments

| Option / argument    | Description                                                                                                               | Default   |
| -------------------- | ------------------------------------------------------------------------------------------------------------------------- | --------- |
| `workbook`           | Workbook filename or project name. `.xlsm` is appended when omitted. Explicit `.xlsm`, `.xlam`, and `.xlsb` are accepted. | Book.xlsm |
| `--with-skill`       | Install the bundled AI-agent skill after scaffolding.                                                                     | false     |
| `--agent <provider>` | Select the target agent skill provider, such as `codex`.                                                                  | -         |
| `--no-update-check`  | Skip the startup update check.                                                                                            | false     |

## Examples

```bash
xlflow new Sales
xlflow new Sales.xlsm
xlflow new SalesAddin.xlam
xlflow new LargeModel.xlsb
xlflow new Sales.xlsm --with-skill --agent codex --json
```

## Notes

::: tip
Use `--with-skill --agent codex` when the project will be maintained primarily by an AI coding agent.
:::

::: warning
The generated workbook must use `.xlsm`, `.xlam`, or `.xlsb` so Excel can preserve VBA components. Omit the extension only when you want the default `.xlsm` project format.
:::

For `.xlam` projects, `new` creates the add-in file in the project's `build/` directory only. It does not install or register the add-in in Excel.

For `.xlsb` projects, xlflow uses Excel COM/VBIDE for VBA synchronization and creates the workbook with Excel file format `50`. Direct OOXML worksheet features such as formula snapshots, workbook cell diff, and pure-Go `pack` do not support `.xlsb`.

`new` uses the `.NET` bridge on Windows in `auto` mode.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "new",
  "workbook": "build/SalesAddin.xlam",
  "source_root": "src",
  "created": ["xlflow.toml", "src/modules/Main.bas"]
}
```

## Related

- [init](./init)
- [doctor](./doctor)
- [skill](./skill)
