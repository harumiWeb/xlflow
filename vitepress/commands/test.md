# xlflow test

Discover and run workbook VBA test procedures.

## Usage

```bash
xlflow test [--filter <pattern>] [--session] [--keepalive]
```

## Options and Arguments

| Option / argument    | Description                             | Default |
| -------------------- | --------------------------------------- | ------- |
| `--filter <pattern>` | Run only matching test names.           | -       |
| `--session`          | Run tests in the managed live workbook. | false   |
| `--keepalive`        | Reuse the bridge process.               | false   |
| `--json`             | Return structured test results.         | false   |

## Examples

```bash
xlflow test --json
xlflow test --filter Smoke --session --json
```

## Notes

::: important
`test` executes VBA. Use a controlled workbook state before running tests that mutate sheets or files.
:::

::: tip
Keep VBA assertions simple and scalar so failures are easy for agents to parse.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "test",
  "tests": [{ "name": "TestSmoke", "status": "pass" }],
  "summary": { "passed": 1, "failed": 0 }
}
```

## Related

- [run](./run)
- [check](./check)
- [error handling guide](../guides/error-handling)
