# Exit Codes

| Code | Meaning                                                                                   |
| ---: | ----------------------------------------------------------------------------------------- |
|  `0` | Success                                                                                   |
|  `1` | Validation or user-code failure (`fmt --check` found unformatted files, etc.)             |
|  `2` | CLI argument or configuration error (including invalid wait combinations)                 |
|  `3` | Operational or environment failure (including workbook busy, wait, and recovery outcomes) |

`diff` returns `0` when differences are found. Inspect `diff.summary.total_diffs` in JSON.
