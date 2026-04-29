# xlflow MVP Quality Hardening Todo

## Planning

- [x] Define the post-MVP quality-hardening scope in `docs/specs/mvp-quality-hardening.md`.
- [x] Write the implementation plan in `docs/superpowers/plans/2026-04-29-mvp-quality-hardening.md`.
- [x] Refresh session tracking docs for the active phase.

## Next Implementation Work

- [ ] Update `docs/specs/cli-contract.md` and `README.md` to reflect the verified verification baseline.
- [ ] Add regression tests for document-module, class, and userform round-trips.
- [ ] Add one obvious local automated verification command.
- [ ] Re-run `tmp_workspaces` E2E verification after quality-hardening changes.

## Verification

- [x] Review existing specs and task-tracking files before planning.
- [x] Save the new phase spec and implementation plan to the repository.
- [ ] Validate future implementation changes with `go test ./...`.
- [ ] Validate real workbook behavior with the `xlflow-tmp-workspace-e2e` skill.
