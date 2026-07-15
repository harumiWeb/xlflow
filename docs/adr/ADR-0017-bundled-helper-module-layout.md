# ADR-0017: Bundled Helper Module Layout

## Status

Accepted

## Context

xlflow has always supported nested directories beneath the configured standard-module root. `push` discovers those sources recursively and, with the default folder configuration, maps the directory structure to VBA `@Folder` annotations. The scaffold nevertheless placed bundled `Xlflow*.bas` helpers beside user-facing entry modules such as `Main.bas`, `App.bas`, and `Ui.bas`. This made the supported organization capability difficult to discover and obscured the boundary between application code and xlflow-provided infrastructure.

## Decision

New and initialized project scaffolds place bundled helper modules under `[src].modules/Xlflow/`:

- `XlflowAssert.bas`
- `XlflowRuntime.bas`
- `XlflowUI.bas`
- `XlflowDebug.bas`

`module install` uses the same location. Existing projects are not migrated automatically. When a legacy root-level helper with the same VBA component name exists, scaffolding and `module install` fail before writing instead of creating a second source file for that component.

## Consequences

- Positive: the scaffold visibly demonstrates nested-module organization and keeps xlflow infrastructure separate from application modules.
- Positive: default `push` behavior maps helpers to the `Xlflow` VBA folder without adding a new import rule.
- Positive: refusing legacy collisions preserves source/component uniqueness for existing projects.
- Negative: users adopting the new layout in an existing project must move helpers deliberately if they want the new folder structure.

## Related

- `docs/adr/ADR-0006-tests-located-in-src.md`
- `docs/specs/cli-contract.md`
- `vitepress/reference/project-structure.md`
