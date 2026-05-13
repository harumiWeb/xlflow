<!-- 設計メモ -->

# Add UserForm support: inspect, snapshot, image export, and future form spec workflow

## Background

xlflow currently works well for normal VBA source workflows such as:

- `pull`
- `push`
- `test`
- `run`
- `save`
- `session`

However, UserForm development is still difficult for AI agents and CLI-based workflows.

A UserForm is not represented by `.frm` text alone. In practice, its state may be spread across:

- `.frm` text
- `.frx` binary data
- VBIDE Designer state
- runtime UserForm state
- code inside `UserForm_Initialize`
- dynamically added controls via `Controls.Add`

This makes UserForm development error-prone for AI agents. In recent testing, the agent was able to complete a UserForm-based macro, but it struggled with layout differences and hidden assumptions around `.frm` / `.frx` / VBE Designer behavior.

## Problem

AI agents can edit `.frm` files as text, but they cannot reliably understand or verify the actual UserForm layout from `.frm` alone.

Common problems:

- `.frm` changes do not always reflect the actual runtime appearance.
- `.frx` may contain binary or designer-backed state that is hard to diff or review.
- VBE Designer state and runtime display may differ.
- `push --session --no-save` can leave disk state stale while live workbook state is newer.
- Existing manually created UserForms are difficult for agents to understand.
- UserForm layout is not currently available as structured JSON.
- Visual verification requires manual VBE / Excel interaction.

## Goals

Add first-class UserForm support to xlflow so that AI agents can inspect, verify, snapshot, and eventually generate UserForms without manually operating VBE.

Primary goals:

- Make existing UserForm state visible from the CLI.
- Export UserForm structure as JSON/YAML.
- Enable visual verification through image export.
- Support both Designer-based and runtime-based inspection.
- Provide a migration path from manually created UserForms to declarative specs.
- Avoid direct `.frx` parsing as the primary strategy.
- Keep `.frx` as a generated or opaque artifact where possible.

Non-goals for the first implementation:

- Complete `.frx` parsing.
- Perfect reconstruction of every MSForms property.
- Full support for external ActiveX controls.
- Full bidirectional Designer/spec synchronization in the initial release.

---

## Proposed feature set

### Phase 1: UserForm discovery and warning improvements

#### Commands affected

- `pull`
- `push`
- `save`
- `inspect`
- `session`
- future `form` commands

#### Behavior

When UserForms are detected, xlflow should emit explicit warnings or hints.

Example:

```text
UserForm detected: UserForm1

Note:
  UserForm state may not be fully represented by .frm text alone.
  Layout and binary properties may depend on .frx and VBIDE Designer state.

Recommended commands:
  xlflow form snapshot UserForm1 --out src/forms/UserForm1.form.json
  xlflow inspect form UserForm1 --runtime --json
  xlflow form export-image UserForm1 --out artifacts/UserForm1.png
```

#### Additional stale-state warning

When using `push --session --no-save`, and the workbook contains UserForms, xlflow should warn that disk state may not match live workbook state.

Example:

```text
Warning: workbook contains UserForms and current session changes are not saved.
Some inspect operations may use saved workbook state, while runtime/session operations may use live state.
Run `xlflow save --session` and `xlflow pull` before reviewing .frm/.frx differences.
```

---

## Phase 2: `xlflow list forms`

### Command

```bash
xlflow list forms --json
```

### Purpose

List UserForms in the current workbook/project.

### Output example

```json
{
  "ok": true,
  "forms": [
    {
      "name": "UserForm1",
      "component_type": "MSForm",
      "has_frx": true,
      "source_path": "src/UserForm1.frm",
      "frx_path": "src/UserForm1.frx"
    }
  ]
}
```

### Implementation notes

This can be implemented by inspecting `VBProject.VBComponents` and filtering components with type `vbext_ct_MSForm`.

Use late binding where possible to avoid requiring explicit VBA references in injected helper code.

```vb
Const vbext_ct_MSForm As Long = 3
```

---

## Phase 3: `xlflow inspect form <name>`

### Commands

```bash
xlflow inspect form UserForm1 --json
xlflow inspect form UserForm1 --designer --json
xlflow inspect form UserForm1 --runtime --json
xlflow inspect form UserForm1 --both --json
```

### Purpose

Return structured information about a UserForm and its controls.

There should be two inspection modes:

| Mode       | Source                 | Purpose                             |
| ---------- | ---------------------- | ----------------------------------- |
| `designer` | `VBComponent.Designer` | Inspect design-time layout          |
| `runtime`  | `UserForms.Add(name)`  | Inspect actual loaded runtime state |

Default mode should likely be `runtime`, because AI agents usually need to verify the actual displayed result.

However, `designer` is important for spec import/export and future form build workflows.

---

## Designer inspection

### Technical approach

Access the Designer object:

```vb
Set comp = ThisWorkbook.VBProject.VBComponents.Item("UserForm1")
Set designer = comp.Designer
```

Then read:

```vb
designer.Caption
designer.Width
designer.Height
designer.Controls
```

For each control:

```vb
ctrl.Name
TypeName(ctrl)
ctrl.Left
ctrl.Top
ctrl.Width
ctrl.Height
ctrl.Caption
ctrl.Text
ctrl.TabIndex
ctrl.Enabled
ctrl.Visible
```

Use safe property access because not all controls support all properties.

### Example output

```json
{
  "ok": true,
  "basis": "designer",
  "form": {
    "name": "UserForm1",
    "caption": "顧客登録",
    "width": 420,
    "height": 320,
    "controls": [
      {
        "name": "lblName",
        "type": "Label",
        "prog_id": "Forms.Label.1",
        "caption": "氏名",
        "left": 24,
        "top": 24,
        "width": 72,
        "height": 18,
        "visible": true,
        "enabled": true
      },
      {
        "name": "txtName",
        "type": "TextBox",
        "prog_id": "Forms.TextBox.1",
        "left": 120,
        "top": 20,
        "width": 180,
        "height": 24,
        "tab_index": 0,
        "visible": true,
        "enabled": true
      },
      {
        "name": "btnSubmit",
        "type": "CommandButton",
        "prog_id": "Forms.CommandButton.1",
        "caption": "登録",
        "left": 240,
        "top": 260,
        "width": 80,
        "height": 30,
        "tab_index": 1,
        "visible": true,
        "enabled": true
      }
    ]
  },
  "warnings": []
}
```

---

## Runtime inspection

### Technical approach

Load the form at runtime:

```vb
Set f = UserForms.Add("UserForm1")
```

Then enumerate:

```vb
f.Controls
```

Finally unload:

```vb
Unload f
Set f = Nothing
```

### Important warning

Runtime inspection executes `UserForm_Initialize`.

Therefore, xlflow must warn about possible side effects.

Example warning:

```json
{
  "warnings": ["Runtime inspection loads the form and executes UserForm_Initialize."]
}
```

### Runtime output example

```json
{
  "ok": true,
  "basis": "runtime",
  "form": {
    "name": "UserForm1",
    "caption": "顧客登録",
    "width": 420,
    "height": 320,
    "controls": [
      {
        "name": "txtName",
        "type": "TextBox",
        "left": 120,
        "top": 20,
        "width": 180,
        "height": 24,
        "tab_index": 0
      }
    ]
  },
  "warnings": ["Runtime inspection executed UserForm_Initialize."]
}
```

---

## Nested controls

UserForms may contain container controls such as:

- `Frame`
- `MultiPage`
- `Page`
- `TabStrip`

Controls inside containers should eventually be represented recursively.

Example:

```json
{
  "name": "fraCustomer",
  "type": "Frame",
  "caption": "顧客情報",
  "left": 16,
  "top": 16,
  "width": 360,
  "height": 120,
  "controls": [
    {
      "name": "txtCustomerName",
      "type": "TextBox",
      "left": 100,
      "top": 20,
      "width": 180,
      "height": 24
    }
  ]
}
```

Coordinates should be parent-relative.

Add this to output metadata:

```json
{
  "coordinate_system": "parent-relative"
}
```

MVP may start with top-level controls only, but Frame support should be prioritized because it is common in business UserForms.

---

## Phase 4: `xlflow form snapshot`

### Command

```bash
xlflow form snapshot UserForm1 --out src/forms/UserForm1.form.json
xlflow form snapshot UserForm1 --out src/forms/UserForm1.form.yaml
```

### Purpose

Convert an existing manually created UserForm into a structured spec.

This is useful for:

- Existing VBA assets.
- AI-agent understanding.
- Git diff review.
- Future migration to declarative form management.
- Future `form build` / `form apply`.

### Important distinction

`snapshot` should not imply complete round-trip support at first.

It means:

```text
Designer state -> structured snapshot/spec
```

Not necessarily:

```text
structured spec -> perfectly identical Designer state
```

### Output example

```yaml
schemaVersion: 1
kind: xlflow.userform
basis: designer
coordinateSystem: parent-relative

form:
  name: UserForm1
  caption: 顧客登録
  width: 420
  height: 320

controls:
  - type: Label
    progId: Forms.Label.1
    name: lblName
    caption: 氏名
    left: 24
    top: 24
    width: 72
    height: 18
    visible: true
    enabled: true

  - type: TextBox
    progId: Forms.TextBox.1
    name: txtName
    left: 120
    top: 20
    width: 180
    height: 24
    tabIndex: 0
    visible: true
    enabled: true

  - type: CommandButton
    progId: Forms.CommandButton.1
    name: btnSubmit
    caption: 登録
    left: 240
    top: 260
    width: 80
    height: 30
    tabIndex: 1
    visible: true
    enabled: true

warnings: []
```

### Unsupported properties

Some properties may be backed by `.frx` or external binary data.

Examples:

- `Picture`
- `MouseIcon`
- Image control binary content
- External ActiveX internal state

Initial behavior should be detection + warning, not full export.

Example:

```yaml
controls:
  - type: Image
    progId: Forms.Image.1
    name: imgLogo
    left: 24
    top: 24
    width: 120
    height: 60
    unsupported:
      - Picture

warnings:
  - control: imgLogo
    message: Picture property may be backed by .frx and is not exported yet.
```

---

## Phase 5: `xlflow form export-image`

### Command

```bash
xlflow form export-image UserForm1 --out artifacts/UserForm1.png
```

### Purpose

Show the UserForm at runtime and export a PNG screenshot.

This is especially useful for AI agents because it enables visual verification without manual VBE/Excel operation.

### Technical approach

1. Start or reuse an Excel session.
2. Inject temporary helper macro.
3. Load the form with `UserForms.Add(name)`.
4. Set a unique temporary caption token.
5. Show the form modeless.
6. Use Win32 API from Go to find the form window.
7. Capture with `PrintWindow` or `BitBlt`.
8. Save PNG.
9. Unload the form.
10. Clean up temporary helper modules.

### VBA-side sketch

```vb
Public Sub __xlflow_show_form_for_capture(ByVal formName As String, ByVal token As String)
    Dim f As Object
    Set f = UserForms.Add(formName)

    f.Caption = f.Caption & " [xlflow-capture-" & token & "]"
    f.Show vbModeless

    DoEvents
End Sub
```

### Go-side window search

Find the target window using a combination of:

- Excel process ID
- unique caption token
- UserForm window class
- visible window state

Possible Win32 APIs:

- `EnumWindows`
- `GetWindowText`
- `GetClassName`
- `GetWindowThreadProcessId`
- `GetWindowRect`
- `PrintWindow`
- `BitBlt`

Do not rely only on window class names such as `ThunderDFrame` / `ThunderXFrame`, because these may differ by Office version/environment.

### Warnings

`form export-image` loads the form at runtime and executes `UserForm_Initialize`.

Example CLI warning:

```text
Warning: form export-image loads the UserForm at runtime.
UserForm_Initialize will be executed.
```

### Experimental status

This command should initially be marked experimental because it depends on:

- Windows GUI behavior
- DPI scaling
- Excel window state
- UserForm initialization behavior
- Win32 capture reliability

Example:

```text
Note: form export-image is experimental and currently supports Windows + desktop Excel only.
```

---

## Phase 6: Designer abstraction layer

### Motivation

Direct use of:

```vb
wb.VBProject.VBComponents.Item("UserForm1").Designer
```

should be hidden behind an internal abstraction.

The rest of xlflow should not depend on raw COM/VBIDE objects.

### Proposed internal structure

```text
internal/excel/forms/
  backend.go
  designer_backend.go
  runtime_backend.go
  snapshot.go
  spec.go
  properties.go
  errors.go
```

### Suggested interfaces

```go
type FormBackend interface {
    ListForms(ctx context.Context) ([]FormInfo, error)
    InspectForm(ctx context.Context, name string) (*FormSnapshot, error)
}

type DesignerFormBackend interface {
    FormBackend
    SnapshotForm(ctx context.Context, name string) (*FormSpec, error)
}

type RuntimeFormBackend interface {
    FormBackend
    ExportImage(ctx context.Context, name string, outPath string) error
}
```

### Suggested data models

```go
type FormSnapshot struct {
    Name             string            `json:"name"`
    Caption          string            `json:"caption,omitempty"`
    Width            float64           `json:"width,omitempty"`
    Height           float64           `json:"height,omitempty"`
    Controls         []ControlSnapshot `json:"controls,omitempty"`
    Basis            string            `json:"basis"` // designer | runtime | source
    CoordinateSystem string            `json:"coordinate_system,omitempty"`
    Warnings         []FormWarning      `json:"warnings,omitempty"`
}
```

```go
type ControlSnapshot struct {
    Name        string             `json:"name"`
    Type        string             `json:"type"`
    ProgID      string             `json:"prog_id,omitempty"`
    Caption     *string            `json:"caption,omitempty"`
    Text        *string            `json:"text,omitempty"`
    Left        float64            `json:"left,omitempty"`
    Top         float64            `json:"top,omitempty"`
    Width       float64            `json:"width,omitempty"`
    Height      float64            `json:"height,omitempty"`
    TabIndex    *int               `json:"tab_index,omitempty"`
    Enabled     *bool              `json:"enabled,omitempty"`
    Visible     *bool              `json:"visible,omitempty"`
    Font        *FontSnapshot      `json:"font,omitempty"`
    Properties  map[string]any     `json:"properties,omitempty"`
    Unsupported []string           `json:"unsupported,omitempty"`
    Controls    []ControlSnapshot  `json:"controls,omitempty"`
}
```

```go
type FormWarning struct {
    Code    string `json:"code"`
    Message string `json:"message"`
    Control string `json:"control,omitempty"`
}
```

---

## Phase 7: Future `form build` / `form apply`

This should be considered future work after read-only inspection and snapshot are stable.

### Possible commands

```bash
xlflow form build src/forms/UserForm1.form.yaml
xlflow form apply src/forms/UserForm1.form.yaml
xlflow form diff UserForm1 --spec src/forms/UserForm1.form.yaml
```

### Design principle

Avoid directly editing `.frx`.

Instead:

```text
form.yaml
  ↓
VBIDE Designer API
  ↓
Excel/VBE saves .frm/.frx
```

### Designer write approach

Use `VBComponent.Designer` and `designer.Controls.Add`.

Example:

```vb
Set comp = ThisWorkbook.VBProject.VBComponents.Item("UserForm1")
Set designer = comp.Designer

Set ctrl = designer.Controls.Add("Forms.TextBox.1", "txtName", True)
ctrl.Left = 120
ctrl.Top = 20
ctrl.Width = 180
ctrl.Height = 24
```

### Important distinction

There are two different `Controls.Add` workflows:

```text
designer.Controls.Add
  - design-time
  - persisted into .frm/.frx
  - useful for form build/apply

runtimeForm.Controls.Add
  - runtime only
  - disappears when form is unloaded
  - useful for dynamic UI
```

xlflow should model these separately.

---

## Phase 8: Dynamic UserForm template

For new AI-generated UserForms, dynamic layout may be more reliable than editing Designer-backed `.frm/.frx`.

### Command idea

```bash
xlflow form init OrderEntryForm --dynamic
```

### Generated files

```text
src/forms/OrderEntryForm.frm
src/forms/OrderEntryFormBuilder.bas
src/forms/OrderEntryFormHandlers.cls
```

### Concept

- `OrderEntryForm.frm` is a minimal empty container.
- `OrderEntryFormBuilder.bas` creates controls with `Controls.Add`.
- `OrderEntryFormHandlers.cls` handles events with `WithEvents`.
- Layout is represented as normal VBA code or generated from future `form.yaml`.

### Example runtime builder

```vb
Public Sub BuildOrderEntryForm(ByVal form As Object)
    Dim lbl As Object
    Set lbl = form.Controls.Add("Forms.Label.1", "lblOrderDate", True)
    With lbl
        .Caption = "受注日"
        .Left = 24
        .Top = 24
        .Width = 80
        .Height = 18
    End With

    Dim txt As Object
    Set txt = form.Controls.Add("Forms.TextBox.1", "txtOrderDate", True)
    With txt
        .Left = 120
        .Top = 20
        .Width = 120
        .Height = 24
    End With
End Sub
```

### Event handler caveat

Dynamically added controls need `WithEvents` handler classes.

Example:

```vb
' Class Module: XlflowButtonHandler
Option Explicit

Public WithEvents Button As MSForms.CommandButton

Private Sub Button_Click()
    MsgBox "Clicked: " & Button.Name
End Sub
```

The UserForm must keep handlers alive:

```vb
Private handlers As Collection

Private Sub UserForm_Initialize()
    Set handlers = New Collection

    Dim btn As MSForms.CommandButton
    Set btn = Me.Controls.Add("Forms.CommandButton.1", "btnSubmit", True)
    btn.Caption = "登録"

    Dim h As XlflowButtonHandler
    Set h = New XlflowButtonHandler
    Set h.Button = btn
    handlers.Add h
End Sub
```

This should be template-generated because AI agents are likely to forget to keep handler instances alive.

---

## Error handling

Add specific error codes for UserForm workflows.

Suggested codes:

```text
form_not_found
vbproject_access_denied
designer_access_failed
runtime_form_load_failed
form_initialize_failed
control_enumeration_failed
unsupported_control
image_capture_failed
window_not_found
temporary_component_cleanup_failed
```

Example JSON error:

```json
{
  "ok": false,
  "error": {
    "code": "vbproject_access_denied",
    "message": "VBProject access is denied.",
    "hint": "Enable 'Trust access to the VBA project object model' in Excel Trust Center."
  }
}
```

---

## Security and environment requirements

UserForm features that use VBIDE require:

```text
Excel Trust Center:
  Trust access to the VBA project object model = enabled
```

These features are Windows + desktop Excel only.

Unsupported environments:

- Excel Online
- LibreOffice
- macOS Excel, at least initially
- headless environments without desktop Excel

---

## Implementation strategy

### Recommended order

1. Add UserForm detection and warnings.
2. Implement `xlflow list forms --json`.
3. Implement `xlflow inspect form <name> --designer --json`.
4. Implement `xlflow inspect form <name> --runtime --json`.
5. Implement `xlflow form snapshot <name> --out ...`.
6. Implement `xlflow form export-image <name> --out ...` as experimental.
7. Add `form diff` against snapshot/spec.
8. Add `form build/apply` via Designer API.
9. Add `form init --dynamic`.

### Why this order

Read-only inspection is much safer than write operations.

`Designer -> Snapshot` gives immediate value for existing manually created UserForms.

`Runtime inspect` and `form export-image` give AI agents a verification loop.

`Spec -> Designer` should come after the snapshot format is proven.

---

## MVP scope

The first deliverable should include:

```text
- xlflow list forms --json
- xlflow inspect form <name> --designer --json
- xlflow inspect form <name> --runtime --json
- xlflow form snapshot <name> --out <path>
- UserForm warnings in push/pull/save/session flows
```

Optional but high-value:

```text
- xlflow form export-image <name> --out <path>
```

Initial supported controls:

```text
- Label
- TextBox
- ComboBox
- ListBox
- CommandButton
- CheckBox
- OptionButton
- Frame
```

Initial properties:

```text
- name
- type
- prog_id
- caption
- text
- left
- top
- width
- height
- tab_index
- enabled
- visible
```

Unsupported properties should be reported, not treated as fatal.

---

## Acceptance criteria

### `list forms`

Given a workbook with `UserForm1`, running:

```bash
xlflow list forms --json
```

returns `UserForm1`.

### `inspect form --designer`

Given a manually created UserForm with a Label, TextBox, and Button, running:

```bash
xlflow inspect form UserForm1 --designer --json
```

returns form metrics and control metrics from the Designer.

### `inspect form --runtime`

Given the same form, running:

```bash
xlflow inspect form UserForm1 --runtime --json
```

loads the form, returns runtime control metrics, unloads the form, and warns that `UserForm_Initialize` was executed.

### `form snapshot`

Given a manually created UserForm, running:

```bash
xlflow form snapshot UserForm1 --out src/forms/UserForm1.form.json
```

writes a structured spec file that can be reviewed by humans and AI agents.

### `form export-image`

Given a UserForm, running:

```bash
xlflow form export-image UserForm1 --out artifacts/UserForm1.png
```

creates a PNG image of the runtime form, or returns a structured `image_capture_failed` / `window_not_found` error.

---

## Risks

### Runtime side effects

Runtime inspection and image export execute `UserForm_Initialize`.

Mitigation:

- Warn clearly.
- Add future `--no-initialize` only if technically possible.
- Encourage separating layout from side-effectful initialization.

### Cleanup failures

Temporary helper modules/forms may fail to delete.

Mitigation:

- Use `__xlflow_` prefix.
- Add retry with `DoEvents`.
- Add cleanup command.
- Report cleanup failure as warning when main operation succeeds.

### DPI and window capture instability

`form export-image` depends on desktop UI behavior.

Mitigation:

- Mark experimental initially.
- Use unique caption token.
- Use process ID + caption token for window discovery.
- Provide structured errors.

### Unsupported controls

External ActiveX controls may not serialize cleanly.

Mitigation:

- Include minimal metrics.
- Mark as `unsupported_control`.
- Avoid failing entire snapshot.

---

## Future ideas

- `xlflow form diff UserForm1 --spec src/forms/UserForm1.form.yaml`
- `xlflow form apply src/forms/UserForm1.form.yaml`
- `xlflow form build src/forms/UserForm1.form.yaml`
- `xlflow form normalize UserForm1`
- `xlflow form validate UserForm1`
- `xlflow form init UserForm1 --dynamic`
- `xlflow form export-assets UserForm1`
- `xlflow inspect form UserForm1 --compare designer,runtime`
- `xlflow cleanup` for stale temporary xlflow components

---

## Summary

This feature would make UserForms observable and eventually controllable from the CLI.

The key design choice is to avoid direct `.frx` parsing as the primary strategy.

Instead:

```text
Existing manual UserForm:
  VBComponent.Designer -> snapshot/spec

Runtime verification:
  UserForms.Add(name) -> inspect/export image

Future generation:
  form.yaml -> VBIDE Designer API -> .frm/.frx generated by Excel/VBE

New AI-generated forms:
  dynamic UserForm template using Controls.Add
```

This would make xlflow significantly more useful for AI-agent-driven Excel/VBA GUI development.
