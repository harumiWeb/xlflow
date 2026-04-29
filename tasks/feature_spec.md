# xlflow tmp_workspaces E2E Skill Spec

## Goal

Create a repository-local skill that standardizes how agents verify `xlflow` by creating disposable projects under `tmp_workspaces` and running the CLI end-to-end.

## Skill Contract

- Skill name: `xlflow-tmp-workspace-e2e`
- Location: `.agents/skills/xlflow-tmp-workspace-e2e`
- Required files:
  - `SKILL.md`
  - `agents/openai.yaml`

## Trigger Scope

Use the skill when Codex needs to:

- create a temporary `xlflow` project under `tmp_workspaces`
- validate `xlflow` CLI behavior with real Excel COM integration
- run repository-standard E2E checks for `new`, `init`, `doctor`, `pull`, `push`, `run`, or `lint`
- verify workbook state or exported VBA artifacts after CLI execution

## Workflow Requirements

The skill should instruct agents to:

1. review `tasks/lessons.md` before starting
2. create a fresh workspace directory under `tmp_workspaces`
3. use installed `xlflow` from PATH, confirming `Get-Command xlflow`
4. run the blank-project path:
   - `xlflow new --json`
   - `xlflow doctor --json`
   - `xlflow pull --json`
   - `xlflow lint --json`
5. when macro execution matters, add a minimal `src/modules/Main.bas` and run:
   - `xlflow lint --json`
   - `xlflow push --json`
   - `xlflow run Main.Run --json`
   - `xlflow pull --json`
   - `xlflow lint --json`
6. verify workbook effects with Excel COM when behavior depends on saved workbook state
7. exercise `init` in a second fresh workspace when requested or when validating project bootstrapping from an existing workbook
8. treat failures as bugs to investigate and fix, then rerun the affected verification path

## Reporting Requirements

- report exact workspace paths under `tmp_workspaces`
- distinguish blank-workbook verification from macro round-trip verification
- include concrete verification evidence, not assumptions
- call out known limitations when a path was not exercised
