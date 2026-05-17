# xlflow export-image

Export a worksheet range as a PNG image through Excel COM.

## Usage

```bash
xlflow export-image --sheet <name> --range <address> [--out <png>|--output-dir <dir>] [--session]
```

## Options and Arguments

| Option / argument    | Description                                    | Default                 |
| -------------------- | ---------------------------------------------- | ----------------------- |
| `--sheet <name>`     | Worksheet to render.                           | active/configured sheet |
| `--range <address>`  | A1 range to export.                            | used range              |
| `--out <png>`        | Exact output image path.                       | -                       |
| `--output-dir <dir>` | Directory for generated image output.          | artifacts               |
| `--name <name>`      | Base filename when using `--output-dir`.       | derived                 |
| `--format png`       | Output image format.                           | png                     |
| `--overwrite`        | Replace an existing image file.                | false                   |
| `--session`          | Render from the managed live session workbook. | false                   |

## Examples

```bash
xlflow export-image --sheet Dashboard --range A1:K30 --out artifacts/dashboard.png --json
xlflow export-image --sheet QR --range A1:AE31 --output-dir artifacts --overwrite --json
```

## Notes

::: tip
Use image exports in PRs and agent loops to review visual workbook output without opening Excel manually.
:::

::: warning
Without `--overwrite`, existing output files may cause the command to fail instead of replacing the image.
:::

## JSON Output Example

Successful `--json` output uses the xlflow envelope plus command-specific fields.

```json
{
  "status": "ok",
  "command": "export-image",
  "sheet": "Dashboard",
  "range": "A1:K30",
  "output": "artifacts/dashboard.png"
}
```

## Related

- [inspect](./inspect)
- [run](./run)
- [Demos](../demos/)
