# Exit Codes

| Code | Meaning                                                                          |
| ---: | -------------------------------------------------------------------------------- |
|  `0` | Success                                                                          |
|  `1` | Validation or user-code failure                                                  |
|  `2` | CLI argument or configuration error                                              |
|  `3` | Environment failure, including Excel, COM, VBIDE, PowerShell, or bridge failures |

`diff` returns `0` when differences are found. Inspect `diff.summary.total_diffs` in JSON.
