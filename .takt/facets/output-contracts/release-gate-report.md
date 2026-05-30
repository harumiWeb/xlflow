# Release Gate Report

Output only the report body.

- Do not include preambles, reasoning notes, or self-referential commentary
- The first non-empty line must be exactly `# Release Gate Report`

## Build Verification

| Build     | Result   | Warnings   | Errors   |
| --------- | -------- | ---------- | -------- |
| `{build}` | {result} | {warnings} | {errors} |

## Automated Test Results

| Command     | Result   | Duration   |
| ----------- | -------- | ---------- |
| `{command}` | {result} | {duration} |

## Windows + Excel COM E2E Evidence

### Test Environment

- **OS**: {os}
- **Working Directory**: {working_directory}
- **Command**: {command}
- **Exit Code**: {exit_code}

### Command Output

```json
{json_output}
```

## Contract Verification

| Field     | Expected   | Actual   | Status   |
| --------- | ---------- | -------- | -------- |
| `{field}` | {expected} | {actual} | {status} |

## Known Limitations

| Item     | Status   | Reason   |
| -------- | -------- | -------- |
| `{item}` | {status} | {reason} |

## Verdict

{PASS | FAIL | BLOCKED}
