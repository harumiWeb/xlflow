# xlflow check

Run lint, analyze, and doctor as a combined preflight.

## Usage

```bash
xlflow check
```

## Options and Arguments

| Option / argument | Description               | Default |
| ----------------- | ------------------------- | ------- |
| `--json`          | Return a combined report. | false   |

## Examples

```bash
xlflow check
```

## Notes

::: tip
Use `check --json` as the default preflight before agent-driven `push` and `run` workflows.
:::

> [!IMPORTANT]
> `check` reports each phase independently, so read the phase list even when the combined status is failure.

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "error",
  "command": "check",
  "phases": [
    { "name": "lint", "status": "ok" },
    { "name": "analyze", "status": "error" },
    { "name": "doctor", "status": "ok" }
  ]
}
```

## Related

- [lint](./lint)
- [analyze](./analyze)
- [doctor](./doctor)
