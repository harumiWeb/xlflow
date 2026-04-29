# xlflow MVP Todo

## Implementation

- [x] Write working feature spec.
- [x] Write persistent CLI contract spec.
- [x] Write ADR-0001 for the Go CLI + PowerShell bridge architecture.
- [x] Implement Go CLI entrypoint and command registration.
- [x] Implement configuration loading and defaults.
- [x] Implement JSON/human output and exit-code mapping.
- [x] Implement project scaffolding for `init`.
- [x] Implement Excel PowerShell bridge and scripts.
- [x] Implement command use cases.
- [x] Implement MVP lint rules.
- [x] Add Go unit tests and script syntax tests.

## Verification

- [x] Run `go test ./...`.
- [x] Run `go run ./cmd/xlflow --help`.
- [x] Run `go run ./cmd/xlflow lint --json` against a scaffolded sample.
- [ ] Run Excel integration commands when Excel/COM is available.
