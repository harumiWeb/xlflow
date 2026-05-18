# xlflow version

Show xlflow build metadata.

## Usage

```bash
xlflow version [--verbose]
```

## Options and Arguments

| Option / argument | Description                                                   | Default |
| ----------------- | ------------------------------------------------------------- | ------- |
| `--verbose`       | Include additional build and runtime metadata when available. | false   |
| `--json`          | Return machine-readable version metadata.                     | false   |

## Examples

```bash
xlflow version
xlflow version --verbose --json
```

## Notes

::: tip
Include `xlflow version --verbose --json` output in bug reports and CI diagnostics.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "version",
  "version": "0.1.0",
  "commit": "abcdef0"
}
```

## Related

- [installation](../installation)
- [troubleshooting](../reference/troubleshooting)
