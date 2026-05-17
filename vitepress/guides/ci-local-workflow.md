# CI and Local Workflow

For docs and source-only changes, run fast checks such as `go test ./...` and `pnpm docs:build` as appropriate.

Workbook behavior changes need Windows + Excel COM validation. CI alone is not enough for release gates that touch VBIDE, sessions, UserForms, or workbook execution.
