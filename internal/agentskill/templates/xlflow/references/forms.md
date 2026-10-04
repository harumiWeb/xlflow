# UserForm Development Reference

Load this reference before creating or editing a UserForm, changing Designer or
code authority, extracting pictures, or inspecting form behavior. For building a
separate release workbook, also load [pack.md](pack.md).

## Choose the Workbook Boundary

| Goal                                                              | Workflow                                                     | Evidence provided                                                            |
| ----------------------------------------------------------------- | ------------------------------------------------------------ | ---------------------------------------------------------------------------- |
| Capture saved Designer, code and supported pictures without Excel | `xlflow pull --backend file --json` in sidecar mode          | Canonical spec, code sidecar and content-addressed assets from saved `.xlsm` |
| Apply canonical source to a closed development workbook           | `xlflow push --backend file --json`                          | Supported Designer/code changes in saved `.xlsm`; no VBE compile             |
| Generate a separate template or blank artifact                    | `xlflow pack ... --json`                                     | New `.xlsm`; no VBE compile or runtime execution                             |
| Inspect or rebuild a live Excel Designer                          | Session-backed `inspect form`, `form snapshot`, `form build` | Excel Designer state; use runtime inspection for execution evidence          |

File commands never read unsaved live state. Establish saved-file authority and
close the workbook/session before file push; use the existing session when live
Excel is authoritative. Unsupported pure-Go operations fail before publication;
do not silently switch to Excel or regenerate an unknown control.

## Canonical Sources and Code Authority

With the default source layout, keep these artifacts under version control:

```text
src/forms/specs/Login.yaml   # Designer intent; JSON/YML also accepted
src/forms/code/Login.bas     # Code-only VBA in sidecar mode
src/forms/assets/logo.bmp    # An image referenced by picture.path
```

Use the configured `[src].forms` root when it differs. Canonical specs and code
sidecars are flat in `specs/` and `code/`; the spec filename must match
`form.name`. Do not put exported `Attribute VB_*` headers in a sidecar.
`[userform].code_source = "sidecar"` makes the sidecar authoritative; `"frm"`
uses code embedded in the matching `.frm`. Confirm this setting before editing
code or migrating an imported project.

File pull in sidecar mode extracts specs, code and supported BMP/JPEG pictures
in one source transaction. Asset paths contain the full content hash; identical
files are reused, conflicting files fail, and obsolete assets are retained.
Excel may persist a JPEG input as BMP, so pull can produce a `.bmp` asset.

File pull does not refresh compatibility `.frm`/`.frx` exports and records
`compatibility_artifact_unsynchronized`. Canonical spec plus sidecar file push
and pack can ignore these stale exports. Excel push or a pack path that selects
the marked `.frm` for code still rejects it with `FRM201`; refresh those exports
with Excel pull before using that authority. Do not remove the warning marker to
force an import. `form snapshot` captures Designer state but does not export the
code sidecar or picture files; use pull when those artifacts are required.

## Minimal New Form

Write this complete spec as `src/forms/specs/Login.yaml` and provide a valid
BMP file at the image path. Geometry is in points relative to the owning parent.

```yaml
schemaVersion: 1
kind: xlflow.userform
basis: designer
coordinateSystem: points
form:
  name: Login
  caption: Sign in
controls:
  - id: details
    name: Details
    type: Frame
    left: 12
    top: 12
    width: 180
    height: 96
    caption: Account
  - id: username
    parentId: details
    name: Username
    type: TextBox
    left: 12
    top: 18
    width: 144
    height: 18
  - id: brand
    name: Brand
    type: Image
    left: 204
    top: 12
    width: 48
    height: 32
    picture:
      path: src/forms/assets/logo.bmp
```

For example, put this code-only event handler in `src/forms/code/Login.bas`:

```vba
Option Explicit

Private Sub UserForm_Initialize()
    Me.Username.Text = "Ready"
End Sub
```

Use file push to apply it to a configured, closed `.xlsm`, or pack to create a
separate workbook. No `.frx` is required for supported canonical generation.
For a new form that needs an explicit client size, use
`form.build.clientWidth` / `clientHeight` in points. Do not combine client and
outer dimensions. Pure-Go template edits reject root-size changes; Excel-backed
`form build` rejects client-size authoring. Root captions must fit the workbook
project code page; see [pack.md](pack.md) for Japanese blank projects.

## Editing Existing Forms

- Treat `controls` as the complete desired control collection, not a patch.
  Omitting an existing control removes it. Stable `id` values connect
  `parentId`; retained binary controls are matched by case-sensitive names.
  Rename or type replacement can replace the control and its persistence.
- Omitted properties on a retained control preserve existing values. Explicit
  empty strings, false and zero are authored values. `observed`, `warnings`
  and `unsupported` describe captured state rather than edit requests.
- Use `zIndex` for sibling ordering. Moving a control keeps omitted numeric
  geometry relative to its new parent; supply new coordinates when needed.
- File push keeps forms omitted from source regardless of pack topology. A
  missing sidecar preserves an existing form's entire code source, an explicitly
  empty sidecar clears its code, and a new form with no sidecar has empty code.
  Pack uses a different code fallback; read [pack.md](pack.md).
- Supported edits preserve unrelated/opaque Designer resources. If an edit
  would lose unsupported state, fix the request or report the capability limit;
  do not strip metadata to make generation succeed.

## Containers, Tabs and Supported Controls

Built-in controls include Label, TextBox, ComboBox, ListBox, CommandButton,
CheckBox, OptionButton, ToggleButton, SpinButton, ScrollBar, Image, Frame,
MultiPage, Page and TabStrip. Support is property- and backend-specific; use
contract diagnostics instead of assuming every captured field is authorable.
Custom ActiveX is outside pure-Go generation; Excel builds may accept explicit
custom ProgIDs with reduced validation and `custom/unchecked` warnings.

Frames can contain controls and nested Frames. A MultiPage contains only Pages;
a Page belongs to a MultiPage, and its controls belong under that Page. A
TabStrip is independent: tabs are records in `tabs`, not child controls.

```yaml
controls:
  - id: steps
    name: Steps
    type: MultiPage
    selectedIndex: 0
  - id: account
    parentId: steps
    name: Account
    type: Page
    caption: Account
  - id: email
    parentId: account
    name: Email
    type: TextBox
  - id: navigation
    name: Navigation
    type: TabStrip
    selectedIndex: 0
    tabs:
      - name: Details
        caption: Details
```

Use this control collection in a complete FormSpec. `selectedIndex` is zero-based;
`-1` represents no logical selection and an empty collection. Tab names must be
unique ignoring case. On edits, omitted `tabs` preserves the collection while
`tabs: []` removes it. Page geometry is derived from its MultiPage; do not copy
Page observed bounds into authored `left`/`top`/`width`/`height` fields. New
pure-Go MultiPage layout uses fixed 96-DPI geometry. ComboBox/ListBox list and
selection capture has weaker persistence guarantees; do not equate it with the
modeled MultiPage/TabStrip collection contract.

## Image Assets

An Image accepts `picture: { path: src/forms/assets/logo.bmp }` for validated
project-relative BMP/JPEG bytes or `picture: { remove: true }` to clear it.
Omission preserves an existing picture and leaves a new Image empty. Null,
empty/mixed objects, `remove: false`, unknown fields and non-Image picture
authoring are rejected. Assets must resolve inside the project, including
symlink targets; limits are 16 MiB and 16 million pixels. Asset bytes are embedded
in the workbook, so runtime does not require the original image file.

Picture authoring is supported by pure-Go pack and file push. Excel-backed
`form build` rejects it before mutation. Keep conventional image suffixes for
clarity, even though explicit references classify assets independently of suffix.
An image-only edit invalidates file-push changed-only state. Unreferenced files
under reserved `assets/` are ignored and never automatically garbage-collected.

## Live Excel Inspection and Rebuild

Use one session for commands against the same workbook; do not run them
concurrently. Start it for a managed closed workbook or attach to the user's
already-open configured workbook. `workbook_busy` supports bounded global
`--wait --wait-timeout <duration>` acquisition; waiting does not retry command
execution.

- Discover names with `xlflow list forms --session --json`.
- Capture without running VBA with `xlflow inspect form Login --designer --session --json`
  or `xlflow form snapshot Login --out src/forms/specs/Login.yaml --session --json`.
- Use `xlflow form build src/forms/specs/Login.yaml --session --json` for
  supported Excel creation, adding `--overwrite` only for intentional replacement.
  Overwrite backs up, deletes, saves and rebuilds; failure restores the old form.
  Sidecar mode reapplies its code or preserves old code when no sidecar exists;
  frm mode preserves old code without consulting sidecars. A same-name non-form
  must not be deleted. `--overwrite --no-save` is invalid.
- `form apply` is hidden; use sidecar-aware `form build` for this workflow.
- Inspect runtime-populated state with `xlflow inspect form Login --runtime --session --json`,
  optionally `--initializer <MethodName>`. Runtime inspection and
  `xlflow form export-image Login --out preview.png --session --json` use a
  temporary workbook copy and may execute UserForm_Initialize.

Check `spec.issues[]` and structured warnings before choosing a rebuild.
Common Excel form failures include `spec_parse_failed`, `spec_validation_failed`
and `spec_schema_invalid`. A snapshot is not a guarantee that every captured
property can be rebuilt through every backend.

## Proof and Completion

For file-backed work, validate source, apply the selected file command, and read
its structured result; report the output path and the absence of VBE validation.
For behavioral acceptance on Windows, use the session proof loop in SKILL.md
to run focused tests/macros. Inspect the rendered form when appearance matters;
use runtime inspection for initialization-populated controls and save/reopen
when persistence is required. Worksheet `export-image` does not prove a
UserForm's appearance. Do not overwrite a verified live workbook with a stale
file artifact or pull old disk state over newer unsaved edits.
