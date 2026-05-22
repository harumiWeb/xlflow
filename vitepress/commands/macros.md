# xlflow macros

Discover runnable public workbook macro entrypoints without executing user code.

## Usage

```bash
xlflow macros [--session]
```

## Options and Arguments

| Option / argument | Description                                         | Default |
| ----------------- | --------------------------------------------------- | ------- |
| `--session`       | Read macro metadata from the managed live workbook. | false   |
| `--json`          | Return macro names and module metadata.             | false   |

## Examples

```bash
xlflow macros
xlflow macros --session --json
```

## Notes

::: tip
Run `macros --json` before `run` so agents can choose an exact entrypoint.
:::

> [!IMPORTANT]
> This command inspects workbook state; it does not execute discovered macros.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "macros",
  "macros": [{ "module": "Main", "name": "Run", "entry": "Main.Run" }]
}
```

## Related

- [run](./run)
- [inspect](./inspect)
