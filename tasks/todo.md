# Push/Run Keepalive Todo

- [x] Add `--keepalive` and `--keepalive-interval` to `push`.
- [x] Add `--keepalive` and `--keepalive-interval` to `run`.
- [x] Emit heartbeat lines from the Go parent process while PowerShell/Excel is still running.
- [x] Emit `XLFLOW_DONE` markers to stderr after success and failure results.
- [x] Keep stdout JSON free of heartbeat and marker lines.
- [x] Update CLI and runtime debugging specs.
- [x] Update bundled xlflow agent skill guidance.
- [x] Add focused unit tests for flags, interval validation, heartbeat, and done markers.
- [x] Run focused Go tests.
- [x] Run full `go test ./...` with an 8-minute timeout.
