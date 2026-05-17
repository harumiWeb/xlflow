# Error Handling

Use exit codes and JSON error codes before editing source. Environment failures usually call for `doctor`; `invoke_macro` failures call for source, trace, and nearby diagnostic inspection.

For unclear runtime failures, run:

```bash
xlflow run Main.Run --trace --session --json
```
