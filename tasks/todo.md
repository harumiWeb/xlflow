# Phase 1-2 Implementation Todo

## Phase 1: Scaffold Cleanup + XlflowAssert Expansion

- [x] 1.1 Create feature_spec.md and todo.md
- [ ] 1.2 Remove `tests/` from scaffold directories in `internal/project/scaffold.go`
- [ ] 1.3 Update `TestInitScaffold` in `internal/project/scaffold_test.go` to remove `tests/` assertion
- [ ] 1.4 Expand `defaultAssertModule` in `internal/project/scaffold.go` with new assertions
- [ ] 1.5 Update `TestScaffoldCreatesAssertHelper` in `scaffold_test.go` for expanded assertions
- [ ] 1.6 Update `docs/specs/cli-contract.md` to document `inconclusive` status
- [ ] 1.7 Update `vitepress/commands/test.md` with `src/modules/Tests/` recommendation
- [ ] 1.8 Update `vitepress/reference/project-structure.md` to remove `tests/` mention
- [ ] 1.9 Create ADR for src-located test policy
- [ ] 1.10 Run `go test ./internal/project/...` to verify scaffold changes

## Phase 2: Lifecycle Hooks + Filters

- [ ] 2.1 Add `Find-XlflowModuleHooks` to `internal/excel/scripts/common.ps1`
- [ ] 2.2 Rewrite `New-XlflowTestRunnerCode` in `common.ps1` for per-test-case execution
- [ ] 2.3 Add `RunHook` and `RunTestCase` VBA functions to runner template
- [ ] 2.4 Rewrite `test.ps1` execution loop to group by module and invoke hooks
- [ ] 2.5 Add `BeforeAll`/`AfterAll` failure: mark all tests in module as failed
- [ ] 2.6 Add `BeforeEach` failure: skip test body but still run `AfterEach`
- [ ] 2.7 Add `AfterEach` failure: override test status to `failed`
- [ ] 2.8 Add `AssertInconclusive` mapping to `inconclusive` status (`vbObjectError + 516`)
- [ ] 2.9 Add `--module` flag to `test` command in `internal/cli/root.go`
- [ ] 2.10 Add `--tag` flag to `test` command in `internal/cli/root.go`
- [ ] 2.11 Pass `--module`/`--tag` through `buildTestScriptArgs` in `internal/excel/bridge.go`
- [ ] 2.12 Accept `ModuleFilter` and `TagFilter` in `test.ps1` params
- [ ] 2.13 Update `Select-XlflowTests` in `common.ps1` for module/tag selection
- [ ] 2.14 Update `internal/output/output.go` for `inconclusive` display (`[?]`)
- [ ] 2.15 Update `docs/specs/cli-contract.md` for hook failures and filter flags
- [ ] 2.16 Update `vitepress/commands/test.md` for hooks and filters
- [ ] 2.17 Create ADR for lifecycle hooks design
- [ ] 2.18 Run `go test ./...` for Go-side changes
- [ ] 2.19 E2E verification with `xlflow-tmp-workspace-e2e` skill

## Verification Checklist

- [ ] `xlflow new` does NOT create `tests/` directory
- [ ] `xlflow init` does NOT create `tests/` directory
- [ ] `xlflow module install` writes expanded `XlflowAssert.bas`
- [ ] Hook execution order: `BeforeAll` -> per-test(`BeforeEach` -> `Test` -> `AfterEach`) -> `AfterAll`
- [ ] `BeforeAll` failure marks all module tests as `failed` with `before_all_failed`
- [ ] `AfterAll` failure marks all module tests as `failed` with `after_all_failed`
- [ ] `BeforeEach` failure skips test body but still runs `AfterEach`
- [ ] `AfterEach` failure overrides test status
- [ ] `AssertInconclusive` maps to `inconclusive` status
- [ ] `--module` filter works (exact module name match)
- [ ] `--tag` flag accepted at CLI (selection logic wired, Phase 3 will implement tag extraction)
