# UserForm Command Reference

Load this reference when the task depends on `xlflow list forms`, `xlflow inspect form`, `xlflow form snapshot`, `xlflow form build`, or `xlflow form export-image`, especially if you need to:

- choose the right form command for design-time vs runtime inspection
- author or edit a persisted UserForm spec
- validate whether a spec shape is supported by `form build`
- understand overwrite safety, session/save behavior, or known Designer-backed limitations

## Command Selection

- Use `xlflow list forms --session --json` to discover workbook UserForm names and expected `.frm` / `.frx` source paths.
- Use `xlflow inspect form <FormName> --designer --session --json` for direct VBIDE Designer inspection without running workbook VBA.
- Use `xlflow inspect form <FormName> --runtime --session --json` for runtime-populated state from a temporary workbook copy. Add `--initializer <MethodName>` when the form must be explicitly populated before inspection.
- Use `xlflow inspect form <FormName> --both --session --json` when you need designer and runtime snapshots in one pass.
- Use `xlflow form snapshot <FormName> --out <path.json|path.yaml|path.yml> --session --json` when you need a persisted Designer spec suitable for review, diff, or later `form build`.
- Prefer `src/forms/specs/<FormName>.yaml` as that persisted artifact path in normal projects.
- Use `xlflow form build <spec> --session --json` to create a new Designer-backed UserForm from a persisted spec.
- Use `xlflow form build <spec> --session --overwrite --json` when the intended workflow is to replace an existing UserForm from spec.
- Use `xlflow form export-image <FormName> --out <path.png> --session --json` when visual verification depends on the runtime-rendered form.
- In `sidecar` mode, run `pull` before reviewing or editing `src/forms/code/<FormName>.bas`; `form snapshot` does not emit code-behind.
- `form apply` is hidden and should not be used for sidecar-aware UserForm workflows; prefer `form build --overwrite`.

## Persisted Spec Contract

`form snapshot` persists an `xlflow.userform` spec. `form build` consumes the same contract.

For ongoing maintenance, treat `src/forms/specs/*.yaml` as the canonical source-controlled artifact for Designer structure. Code-behind authority depends on `[userform].code_source`: new projects default to `sidecar`, where `src/forms/code/*.bas` is canonical, while imported projects default to `frm`, where `.frm` embedded code remains canonical until migration. Exported `.frm` / `.frx` files can stay in the repository, but they are generated Designer artifacts rather than the primary source of truth for Designer-backed behavior. Successful `form build` re-materializes them back into `src/forms/`, and `push` now fails preflight instead of importing stale/mismatched artifacts when spec filename, `form.name`, `.frm` basename, or `.frm` `Attribute VB_Name` disagree.

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
- Each authored control must have a stable `id`.
- `parentId` is optional and case-sensitive. When present, it must reference another control `id`.
- For known first-class controls, `parentId` must reference a container-capable control. Currently `Frame` is container-capable; custom controls with explicit `progId` remain unchecked for compatibility.
- `zIndex` is optional and is used to preserve sibling ordering when present.
- Explicit duplicate `id` values are validation errors. xlflow does not auto-correct them.
- Self-referencing `parentId` values and parent cycles are validation errors.
- Legacy nested `controls` input may still be accepted and normalized into the flat array, but new specs should use the flat structure directly.

## Supported Build Semantics

`form build` is designed to recreate UserForm structure and common Designer-backed properties, not to guarantee a lossless round-trip for every captured field.

Strongly supported:

- UserForm creation from `form.name`
- top-level vs nested control structure from `id` / `parentId`
- supported control types and ProgID mapping
- common geometry and visual properties such as `caption`, `text`, `value`, `visible`, and `enabled`

xlflow core defines the canonical UserForm contract for document fields, form fields, built-in control types, property value types, support levels, ProgID mappings, and container capability. `form build` validates raw YAML/JSON specs against that contract before Excel opens, including unknown fields, incorrect value types, fixed values, type-specific control properties, parent references, and known built-in `type` / `progId` consistency. Invalid specs fail with all detected issues in `spec.issues[]` so editors and agents can show multiple diagnostics from one pass.

Custom controls with an explicit custom `progId` remain accepted with reduced validation. xlflow validates common structural fields and the `properties` bag, then returns a `custom/unchecked` warning because type-specific behavior cannot be verified from the built-in contract.

Best-effort or observed-only:

- form-level `width` / `height`: best-effort only
- design-time `ComboBox` / `ListBox` `list` / `selectedIndex`: observed-only for round-trip expectations, even though xlflow still attempts to apply them
- persisted `warnings`, `observed`, `unsupported`, and legacy nested `controls`: snapshot-only metadata

Expect successful `form build` responses to return contract warnings when the spec depends on those weaker fields.

When `form build` fails before Excel opens, expect structured `spec_parse_failed`, `spec_validation_failed`, or `spec_schema_invalid` errors plus top-level `spec` metadata such as `path`, `format`, optional `line` / `column`, optional first `field`, optional remediation `suggestion`, and `issues[]` containing all validation issues.

## Overwrite and Safety Rules

- Without `--overwrite`, `form build` fails with `form_already_exists` when a UserForm with the same `form.name` already exists.
- `--overwrite` is only for replacing an existing UserForm. A same-name non-UserForm component must not be deleted.
- Overwrite is implemented as export-backup -> delete -> save -> rebuild.
- In `sidecar` mode, xlflow synchronizes tracked `.frm` embedded code from `src/forms/code/<FormName>.bas` and reapplies that sidecar to the new UserForm when present.
- In `sidecar` mode, if no sidecar exists yet, overwrite falls back to the deleted workbook form's code-behind so the rebuild does not silently drop VBA lines.
- In `frm` mode, overwrite preserves the deleted workbook form's code-behind without consulting `src/forms/code`.
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
- `form snapshot` uses the same non-executing Designer basis as `inspect form --designer`; it resolves concrete control types from `ProgId` or COM metadata when Excel exposes them.
- Disk-backed `inspect workbook|sheets|range|used-range|cell` commands do not reflect unsaved live session changes. Run `xlflow save --json` first if the live workbook may be newer than disk.

## Recommended Agent Workflow

When an agent needs to review or regenerate a UserForm safely:

1. `xlflow list forms --session --json`
2. `xlflow inspect form <FormName> --designer --session --json`
3. `xlflow pull --session --json`
4. `xlflow form snapshot <FormName> --out src/forms/specs/<FormName>.yaml --session --json`
5. in `sidecar` mode, review or edit `src/forms/code/<FormName>.bas` if code-behind changed
6. edit the persisted spec under `src/forms/specs/`
7. `xlflow form build src/forms/specs/<FormName>.yaml --session --overwrite --json`
8. inspect the result with `inspect form` and, when visuals matter, `form export-image`

Treat `form build` as a deterministic scaffold/rebuild command for supported structure and common properties, not as a promise of lossless Designer round-trip fidelity.
