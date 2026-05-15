# UserForm Command Reference

Load this reference when the task depends on `xlflow list forms`, `xlflow inspect form`, `xlflow form snapshot`, `xlflow form build`, or `xlflow form export-image`, especially if you need to:

- choose the right form command for design-time vs runtime inspection
- author or edit a persisted UserForm spec
- validate whether a spec shape is supported by `form build`
- understand overwrite safety, session/save behavior, or known Designer-backed limitations

## Command Selection

- Use `xlflow list forms --session --keepalive --json` to discover workbook UserForm names and expected `.frm` / `.frx` source paths.
- Use `xlflow inspect form <FormName> --designer --session --keepalive --json` for direct VBIDE Designer inspection without running workbook VBA.
- Use `xlflow inspect form <FormName> --runtime --session --keepalive --json` for runtime-populated state from a temporary workbook copy. Add `--initializer <MethodName>` when the form must be explicitly populated before inspection.
- Use `xlflow inspect form <FormName> --both --session --keepalive --json` when you need designer and runtime snapshots in one pass.
- Use `xlflow form snapshot <FormName> --out <path.json|path.yaml|path.yml> --session --keepalive --json` when you need a persisted strict Designer spec suitable for review, diff, or later `form build`.
- Prefer `src/forms/specs/<FormName>.yaml` as that persisted artifact path in normal projects.
- Use `xlflow form build <spec> --session --keepalive --json` to create a new Designer-backed UserForm from a persisted spec.
- Use `xlflow form build <spec> --session --overwrite --keepalive --json` when the intended workflow is to replace an existing UserForm from spec.
- Use `xlflow form export-image <FormName> --out <path.png> --session --keepalive --json` when visual verification depends on the runtime-rendered form.

## Persisted Spec Contract

`form snapshot` persists an `xlflow.userform` spec. `form build` consumes the same contract.

For ongoing maintenance, treat `src/forms/specs/*.yaml` as the canonical source-controlled artifact. Exported `.frm` / `.frx` files can stay in the repository, but they are build or pull artifacts rather than the primary source of truth for Designer-backed behavior.

Required top-level fields:

- `schemaVersion: 1`
- `kind: "xlflow.userform"`
- `basis: "designer"`
- `form`
- `controls`

Typical shape:

```yaml
schemaVersion: 1
kind: xlflow.userform
basis: designer
coordinateSystem: points
form:
  name: CustomerForm
  caption: Customer
  observed:
    width: 240
    height: 180
controls:
  - id: frame_main
    name: FrameMain
    type: Frame
    progId: Forms.Frame.1
    left: 12
    top: 12
    width: 216
    height: 96
    caption: Details
  - id: label_name
    parentId: frame_main
    zIndex: 0
    name: LabelName
    type: Label
    progId: Forms.Label.1
    left: 12
    top: 18
    width: 48
    height: 18
    caption: Name
```

## Canonical Structural Rules

- `controls` is the canonical flat capture array.
- Each control must have a stable `id`.
- `parentId` is optional and case-sensitive. When present, it must reference another control `id`.
- `zIndex` is optional and is used to preserve sibling ordering when present.
- Explicit duplicate `id` values are validation errors. xlflow does not auto-correct them.
- Legacy nested `controls` input may still be accepted and normalized into the flat array, but new specs should use the flat structure directly.

## Supported Build Semantics

`form build` is designed to recreate UserForm structure and common Designer-backed properties, not to guarantee a lossless round-trip for every captured field.

Strongly supported:

- UserForm creation from `form.name`
- top-level vs nested control structure from `id` / `parentId`
- supported control types and ProgID mapping
- common geometry and visual properties such as `caption`, `text`, `value`, `visible`, and `enabled`

Best-effort or observed-only:

- form-level `width` / `height`: best-effort only
- design-time `ComboBox` / `ListBox` `list` / `selectedIndex`: observed-only for round-trip expectations, even though xlflow still attempts to apply them

Expect successful `form build` responses to return contract warnings when the spec depends on those weaker fields.

When `form build` fails before Excel opens, expect structured `spec_parse_failed`, `spec_validation_failed`, or `spec_schema_invalid` errors plus top-level `spec` metadata such as `path`, `format`, optional `line` / `column`, optional `field`, and an optional remediation `suggestion`.

## Overwrite and Safety Rules

- Without `--overwrite`, `form build` fails with `form_already_exists` when a UserForm with the same `form.name` already exists.
- `--overwrite` is only for replacing an existing UserForm. A same-name non-UserForm component must not be deleted.
- Overwrite is implemented as export-backup -> delete -> save -> rebuild.
- If rebuild fails after the delete/save checkpoint, xlflow restores the original UserForm from the temporary export before returning failure.
- `--no-save` is valid only with `--session`.
- `--overwrite --no-save` is invalid because Excel requires an intermediate save after removing the old UserForm and before recreating it.

## Supported Control Types

Current first-class build support covers these controls:

- `Label`
- `TextBox`
- `ComboBox`
- `ListBox`
- `CommandButton`
- `CheckBox`
- `OptionButton`
- `Frame`

When `progId` is present in the spec, xlflow prefers it. Otherwise it falls back to the built-in type-to-ProgID mapping for supported controls.

## Session and Inspection Caveats

- `inspect form --designer` reads the source workbook Designer directly.
- `inspect form --runtime` and `form export-image` execute against a temporary workbook copy and may run `UserForm_Initialize` plus an optional explicit initializer.
- `form snapshot` is stricter than `inspect form --designer`; it opens a temporary workbook copy and runs an injected helper to recover concrete control types for the persisted artifact.
- Disk-backed `inspect workbook|sheets|range|used-range|cell` commands do not reflect unsaved live session changes. Run `xlflow save --json` first if the live workbook may be newer than disk.

## Recommended Agent Workflow

When an agent needs to review or regenerate a UserForm safely:

1. `xlflow list forms --session --keepalive --json`
2. `xlflow inspect form <FormName> --designer --session --keepalive --json`
3. `xlflow form snapshot <FormName> --out src/forms/specs/<FormName>.yaml --session --keepalive --json`
4. edit the persisted spec under `src/forms/specs/`
5. `xlflow form build src/forms/specs/<FormName>.yaml --session --overwrite --keepalive --json`
6. inspect the result with `inspect form` and, when visuals matter, `form export-image`

Treat `form build` as a deterministic scaffold/rebuild command for supported structure and common properties, not as a promise of lossless Designer round-trip fidelity.
