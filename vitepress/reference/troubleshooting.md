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

## Workbook output depends on formatting

Use:

```bash
xlflow inspect range --sheet Result --address A1:F20 --include-style --json
xlflow export-image --sheet Result --range A1:F20 --json
```
