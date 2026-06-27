# xlflow runner

Manage the persistent xlflow runner marker module used by some workbook execution flows.

## Usage

```bash
xlflow runner install
xlflow runner status
xlflow runner remove
```

## Options and Arguments

| Option / argument | Description                                 | Default |
| ----------------- | ------------------------------------------- | ------- |
| `install`         | Install or update the runner marker module. | -       |
| `status`          | Report whether the marker is present.       | -       |
| `remove`          | Remove the marker module.                   | -       |
| `--json`          | Return structured runner state.             | false   |

## Examples

```bash
xlflow runner status --json
xlflow runner install --json
```

## Notes

::: tip
Most users can rely on `xlflow run`; use `runner` directly when debugging persistent runner state.
:::

`runner` uses the `.NET` bridge on Windows in `auto` mode for `install`, `status`, and `remove`.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "runner status",
  "runner": { "installed": true, "module": "XlflowRunner" }
}
```

## Related

- [run](./run)
