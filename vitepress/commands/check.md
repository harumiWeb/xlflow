# xlflow check

Run lint, analyze, and doctor as a combined preflight.

## Usage

```bash
xlflow check [--keepalive] [--keepalive-interval <duration>]
```

## Options and Arguments

| Option / argument                 | Description                                    | Default         |
| --------------------------------- | ---------------------------------------------- | --------------- |
| `--keepalive`                     | Reuse the bridge process during doctor checks. | false           |
| `--keepalive-interval <duration>` | Bridge keepalive interval.                     | command default |
| `--json`                          | Return a combined report.                      | false           |

## Examples

```bash
xlflow check
xlflow check --keepalive --json
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
