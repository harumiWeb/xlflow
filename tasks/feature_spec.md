# xlflow Scaffold Spec

## Goal

Define the current project scaffold behavior for `xlflow new` and `xlflow init`.

## Behavior

- `xlflow new` and `xlflow init` create or update a project-local `.gitignore`.
- New `.gitignore` files contain only xlflow project artifacts and Excel temporary files:
  - `~$*.xls*`
  - `*.tmp`
  - `.xlflow/`
  - `build/`
- Existing `.gitignore` files are preserved and receive only missing managed entries.
- Existing managed entries are not duplicated.
- `xlflow skill install` installs the bundled `xlflow` Skill.
- `xlflow new/init --with-skill` installs the same Skill during project creation.
- Supported providers are `agents`, `codex`, `claude`, `cursor`, `gemini`, and `copilot`.
- Provider defaults install to `<provider-dir>/skills/xlflow`, for example `.codex/skills/xlflow`.
- `--target <dir>` installs to `<dir>/xlflow`.
- Existing Skill directories are not overwritten unless `--force` is set.
- Interactive terminals use a Bubble Tea selector when no provider or target is specified.
- JSON and non-interactive runs require `--agent` or `--target`.
- `new` and `init` no longer create `prompts/agent.md`.

## Interfaces

- CLI: `xlflow [--json] skill install [--agent <provider> | --target <dir>] [--force]`
- CLI: `xlflow [--json] new [workbook] [--with-skill] [--agent <provider>]`
- CLI: `xlflow [--json] init <workbook> [--with-skill] [--agent <provider>]`
- Skill artifact: bundled `xlflow/SKILL.md` with workflow, validation, trace, lint, test, diff, and final reporting guidance.

## Verification

- Fast gate: `go test ./...` and `task verify`.
- Skill gate: `skill-creator` quick validation for the bundled skill folder.
- Coverage: `.gitignore` creation and append behavior, scaffold prompt removal, provider install paths, overwrite refusal and `--force`, `init --with-skill`, non-interactive JSON failure, and Bubble Tea selector model behavior.
