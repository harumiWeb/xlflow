# Troubleshooting

## Excel or VBIDE access fails

Run:

```bash
xlflow doctor --json
```

Enable Trust access to the VBA project object model in Excel.

## Macro target is unclear

Run:

```bash
xlflow macros --json
```

Use a discovered `qualified_name`.

## Runtime failure is unclear

Run with structured debug output:

```bash
xlflow run Main.Run --json
```

## Workbook reports recovery required

`workbook_recovery_required` means a previous Excel/VBA operation returned
without proving that execution and cleanup finished. Do not retry with `--wait`;
the marker is not lock contention.

1. Run `xlflow status --json` and inspect `coordination.recovery`.
2. For a managed session, use `xlflow session stop --discard --json`.
3. If an affected PID is shown, use `xlflow process cleanup <pid> --json`.
4. After manually closing Excel without saving, use
   `xlflow recovery clear --json`.
5. Use `recovery clear --force` only when you accept that it clears the marker
   without stopping VBA or proving workbook safety.

See [recovery](../commands/recovery).

## Workbook output depends on formatting

Use:

```bash
xlflow inspect range --sheet Result --address A1:F20 --include-style --json
xlflow export-image --sheet Result --range A1:F20 --json
```
