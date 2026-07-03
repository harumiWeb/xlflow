# xlflow update

Check whether a newer xlflow release is available.

## Usage

```bash
xlflow update check [--json]
```

## Options and Arguments

| Option / argument | Description                                 | Default |
| ----------------- | ------------------------------------------- | ------- |
| `check`           | Query the latest GitHub Release for xlflow. |         |
| `--json`          | Return machine-readable update information. | false   |

## Examples

```bash
xlflow update check
xlflow update check --json
```

## Notes

`update check` does not require Excel COM or an xlflow project. Development builds such as `dev` are treated as not updateable because they cannot be compared to release versions.

## JSON Output Example

```json
{
  "status": "ok",
  "command": "update check",
  "update": {
    "current_version": "0.16.0",
    "latest_version": "v0.16.1",
    "update_available": true,
    "release_url": "https://github.com/harumiWeb/xlflow/releases/tag/v0.16.1"
  }
}
```

## Related

- [version](./version)
- [installation](../installation)
