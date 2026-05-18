# JSON Output

All commands accept the global `--json` flag.

```json
{
  "status": "ok",
  "command": "lint",
  "error": null,
  "logs": []
}
```

Failures set `status` to `failed` and populate `error`:

```json
{
  "status": "failed",
  "command": "run",
  "error": {
    "code": "macro_failed",
    "message": "Main Err 5: inputPath is required",
    "phase": "invoke_macro"
  },
  "logs": []
}
```

Command-specific fields are top-level fields such as `issues`, `analysis`, `macro`, `macros`, `tests`, `diff`, `inspect`, `trace`, `target`, `session`, `warnings`, `hints`, `output`, `forms`, `edit`, `runner`, and `version`.

Source: `docs/specs/cli-contract.md`.
