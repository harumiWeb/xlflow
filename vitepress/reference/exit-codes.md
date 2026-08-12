# Exit Codes

| Code | Meaning                                                                                                           |
| ---: | ----------------------------------------------------------------------------------------------------------------- |
|  `0` | Success                                                                                                           |
|  `1` | Validation or user-code failure (`fmt --check` found unformatted files, `metrics` thresholds were exceeded, etc.) |
|  `2` | CLI argument or configuration error (including invalid wait combinations and metrics thresholds/exclusions)       |
|  `3` | Operational or environment failure (including workbook busy, wait, and recovery outcomes)                         |

`diff` returns `0` when differences are found. Inspect `diff.summary.total_diffs` in JSON.
