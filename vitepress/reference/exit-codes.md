# Exit Codes

| Code | Meaning                                                                          |
| ---: | -------------------------------------------------------------------------------- |
|  `0` | Success                                                                          |
|  `1` | Validation or user-code failure (`fmt --check` found unformatted files, etc.)    |
|  `2` | CLI argument or configuration error (invalid `fmt` mode combinations, etc.)      |
|  `3` | Environment failure, including Excel, COM, VBIDE, PowerShell, or bridge failures |

`diff` returns `0` when differences are found. Inspect `diff.summary.total_diffs` in JSON.
