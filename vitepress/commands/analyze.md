# xlflow analyze

Analyze VBA source for runtime-risk patterns without Excel COM.

## Usage

```bash
xlflow analyze
```

## Options and Arguments

| Option / argument | Description                          | Default |
| ----------------- | ------------------------------------ | ------- |
| `--json`          | Return structured analysis findings. | false   |

## Examples

```bash
xlflow analyze
xlflow analyze --json
```

## Notes

::: tip
Use `analyze` for fast source-level feedback before opening Excel.
:::

> [!IMPORTANT]
> Findings that block automation return a failure status and exit code `1`.

## Rules

| Code     | Severity | Description                                                               |
| -------- | -------- | ------------------------------------------------------------------------- |
| `VBA101` | warning  | Object variable assignment likely missing `Set`.                          |
| `VBA102` | warning  | Object-returning project function assignment likely missing `Set`.        |
| `VBA103` | warning  | Object-returning function body likely missing `Set <FunctionName> = ...`. |
| `VBA104` | error    | Known Excel object/member mismatch such as `Worksheet.DisplayGridlines`.  |
| `VBA105` | error    | Removed `XlflowLog` trace helper call.                                    |
| `VBA106` | error    | Removed `XlflowSetTraceFile` trace helper call.                           |
| `VBA201` | warning  | `Range.Find` result is dereferenced before a `Nothing` check.             |
| `VBA202` | warning  | Object variable may be used before an obvious `Set` assignment.           |
| `VBA203` | warning  | `Application` state is changed without an obvious restore path.           |
| `VBA204` | warning  | Normal execution can fall through into an error-handler label.            |
| `VBA205` | warning  | Unqualified or active Excel object access depends on runtime sheet state. |
| `VBA206` | warning  | Likely ByRef argument type mismatch against a project-local procedure.    |
| `VBA207` | warning  | `Dictionary` or `Collection` item access has no obvious existence guard.  |
| `VBA208` | warning  | `ReDim Preserve` is used on a multi-dimensional array.                    |
| `VBA209` | warning  | Object or array is compared with scalar equality.                         |
| `VBA210` | warning  | Function may exit without assigning its return value.                     |
| `VBA211` | error    | Expanded known Excel object/member mismatch.                              |

Rules `VBA101` through `VBA106`, `VBA201` through `VBA205`, `VBA208`, `VBA209`, and `VBA211` are enabled by default. Rules `VBA206`, `VBA207`, and `VBA210` are opt-in through `[analyze]` because they are more dataflow-sensitive.

## JSON Output Example

Failed `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "failed",
  "command": "analyze",
  "error": {
    "code": "analyze_failed",
    "message": "1 analysis finding(s) found"
  },
  "analysis": [
    {
      "code": "VBA201",
      "severity": "warning",
      "file": "src/modules/Main.bas",
      "module": "Main",
      "procedure": "Run",
      "line": 12,
      "message": "Range.Find result found is dereferenced before a Nothing check."
    }
  ]
}
```

## Related

- [lint](./lint)
- [check](./check)
- [inspect-gui](./inspect-gui)
