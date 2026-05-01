# Runtime Diagnostics Todo

- [x] Add `internal/analyze` with `VBA101`, `VBA102`, and `VBA103`.
- [x] Add `xlflow analyze` and top-level `analysis` output.
- [x] Add `xlflow check` aggregate output.
- [x] Add `run_diagnostic` enrichment for macro failures.
- [x] Add `trace enable/disable/status/clean` while preserving `trace inject`.
- [x] Move traced run logs to `.xlflow/traces` and support temporary helper injection.
- [x] Update scaffolded source modules toward `Main.Run -> App.RunCore`.
- [x] Update README, specs, bundled skill, and working task docs.
- [x] Run full `go test ./...` with an 8-minute timeout.
- [x] Run Excel COM-backed script tests available in the local environment.
