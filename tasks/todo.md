# xlflow AI Agent Skill Todo

## Implementation

- [x] Remove scaffolded `prompts/agent.md` generation.
- [x] Add bundled `xlflow` Skill artifact.
- [x] Add provider-aware Skill installer.
- [x] Add `xlflow skill install`.
- [x] Add `new/init --with-skill`.
- [x] Add Bubble Tea provider selector.
- [x] Update CLI contract, README, and ADR.

## Verification

- [x] Add unit tests for scaffold prompt removal, Skill installation, overwrite behavior, CLI flags, `init --with-skill`, JSON non-interactive behavior, and selector model behavior.
- [x] Run `skill-creator` quick validation.
- [x] Run `go test ./...`.
- [x] Run `task verify`.
