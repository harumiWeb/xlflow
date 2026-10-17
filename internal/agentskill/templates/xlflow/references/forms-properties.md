# FormSpec Property Guidance

Read this before writing field values. Use the bundled, generated
[property contract](forms-contract.md) for exact keys, types, required members,
enum values and control applicability. Use [complete examples](forms-examples.md)
for starting layouts and [forms.md](forms.md) for source authority, applying
changes and verification. No xlflow repository checkout or editor Hover is
needed to discover the supported fields.

## Choose a Field

1. Identify the target: `form`, a control with a known `type`, or a record in
   `TabStrip.tabs`. They have different property sets.
2. Check common control fields and type-specific rows in the contract. Page
   excludes common geometry. A property supported by one MSForms class is not
   automatically supported by another.
3. Check support level and the additional constraints below. `supported` means
   modeled authoring; existing opaque binary state can still limit an edit.
   `observed-only` identifies captured state with limited persistence; the
   current pure-Go compiler rejects list/selection edits as described below.
4. Use canonical camelCase keys and native YAML/JSON values. VBA property names,
   `fm*` constants and snake_case inspection keys are not FormSpec syntax.

Do not invent `multiLine`, `wordWrap`, `default`, `cancel`, `groupName`, `min`
or `max` from general MSForms knowledge: they are not modeled top-level fields
in the current contract. `properties` is an unchecked compatibility bag, not
a guarantee of pure-Go support. For a missing field, explain the limitation and
assess supported runtime VBA initialization when appropriate. Runtime assignment
does not make a field a persisted Designer property.

## Required Structure and Values

A complete document has `schemaVersion: 1`, `kind: xlflow.userform`,
`basis: designer`, `form` with `name`, and a `controls` array. Each control
requires `id`, `name`, `type`. IDs must be unique and stable. Use built-in type
names from the contract; `progId` is normally unnecessary for built-ins.
An explicit built-in ProgID must match its type. Custom ActiveX requires its
own ProgID and is outside pure-Go generation.

| Intent                         | Write                                        | Avoid                                       |
| ------------------------------ | -------------------------------------------- | ------------------------------------------- |
| Boolean                        | `enabled: false`                             | `enabled: "false"`                          |
| Integer enum                   | `textAlign: 2`                               | `textAlign: Center` or `"2"`                |
| Literal text                   | `caption: "Yes"`                             | Relying on YAML scalar coercion for text    |
| RGB color                      | `backColor: { kind: rgb, value: "#FFFFFF" }` | Bare `#FFFFFF`, which starts a YAML comment |
| No edit to a retained property | Omit the key                                 | Adding `null` or a rendering default        |
| Explicit empty text            | `text: ""`                                   | Assuming omission clears text               |

JSON uses the same hierarchy and keys, double-quoted strings, lowercase
`true`/`false`, and numeric enums. Optional does not mean nullable: do not use
`null` for removal. Use the explicit picture/collection actions below.

## Geometry, Identity and Order

Geometry is in points (72 points per inch). Control `left` and `top` are
relative to their owning parent even though `controls` is flat. `parentId`
refers to an ID, not a VBA name; omit it for form-root controls.
`coordinateSystem` accepts `points` or `parent-relative`; neither enables
pixels or absolute screen coordinates. `zIndex` describes sibling order;
`tabIndex` describes keyboard tab order.

`form.width` and `form.height` author persisted MS-OFORMS `DisplayedSize`.
New forms default omitted axes to 240 by 180 points. Existing forms preserve
omitted axes. VBFrame client fields are compiler-maintained companion values,
not the source of these dimensions. Runtime `InsideWidth`/`InsideHeight` and
outer dimensions are separate observations; do not add fixed border offsets.
Form dimensions must be finite, non-negative and representable in HIMETRIC.

Page belongs to MultiPage and has no authored bounds. Put child controls under
the Page's ID and set their relative geometry. MultiPage contains only Pages;
Frame and Page contain permitted controls. TabStrip contains tab records,
not child controls.

## Text, Values and Selection

| Target                               | Typical authored state              | Notes                                                                                  |
| ------------------------------------ | ----------------------------------- | -------------------------------------------------------------------------------------- |
| Label, CommandButton, Frame          | `caption: "Continue"`               | `text` is not their caption field                                                      |
| TextBox                              | `text: ""`                          | Prefer one representation; do not author contradictory `text` and `value`              |
| CheckBox, OptionButton, ToggleButton | `value: false`                      | Useful initial state; `any` is not a promise to persist arbitrary objects              |
| SpinButton, ScrollBar                | `value: 0`                          | Integer; no modeled `min`/`max` keys                                                   |
| ComboBox                             | `text: "One"` or `value: "One"`     | String value is modeled; initialize list items/selection in runtime VBA                |
| ListBox                              | Omit list/value/selection authoring | List/selection edits and ListBox text/value edits are rejected by the pure-Go compiler |
| MultiPage, TabStrip                  | `selectedIndex: 0`                  | Modeled zero-based selection; `-1` for no logical selection                            |

Pure-Go text/value compilation for TextBox and ComboBox requires strings.
The compiler rejects ListBox text/value edits because they depend on list state
whose persistence is not established. ComboBox/ListBox `list` and `selectedIndex`
are capture fields, not reliable authoring inputs: list/selection edits are
rejected, and new-form generation rejects nonempty `list` or an explicit
`selectedIndex`. Leave those fields unspecified for a new control and populate
list items and selection in runtime VBA when needed. This restriction does not
apply to modeled MultiPage/TabStrip selection.
CheckBox/OptionButton checked state accepts booleans or 0/1/-1;
tri-state is not supported. Prefer booleans, also for new ToggleButtons.
New SpinButton/ScrollBar generation supports the default value range 0..100;
the broad `integer` metadata does not remove this compiler constraint.
`tabIndex` must be in 0..32767 for compilation.

For MultiPage/TabStrip, selection must be consistent with the Page/Tab collection
size. Empty collections normalize to `-1`. TabStrip `tabs` records require
`name` and allow `caption`, `controlTipText`, `tag`, `accelerator`, `enabled`,
`visible`. Names must be unique ignoring case within that TabStrip. Page has
its own caption, tooltip, tag and accelerator fields; they are not common fields
for every control. Model hierarchy with `parentId`, not a grouping property.

## Fonts

```yaml
font:
  name: MS UI Gothic
  size: 9
  bold: false
  italic: false
  underline: false
  strikeThrough: false
  charset: 128
```

Place this object under `form` or an applicable control. Members are partial
edits: specify only what should change. Size is in points, greater than 0 and
at most 3276.75; TextProps persistence requires 0.05-point increments.
`charset` is an integer from 0 to 255 (128 is useful for Japanese text), not
limited to a short preset list.

Form and most text/container controls accept Font. Page, Image, SpinButton,
ScrollBar and individual Tab records do not. Empty font names and non-ASCII
name updates to an existing StdFont are rejected. Omission preserves the
existing representation. A form font does not request a recursive source
rewrite of every control.

## Colors and Appearance

```yaml
foreColor: { kind: rgb, value: "#203040" }
backColor: { kind: system, value: buttonFace }
borderStyle: 1
```

| kind     | value                                      | Meaning                                                       |
| -------- | ------------------------------------------ | ------------------------------------------------------------- |
| `rgb`    | `"#RRGGBB"`                                | Fixed RGB                                                     |
| `system` | `buttonFace`, `window`, `windowText`, etc. | Named OLE system color; preserves system semantics            |
| `ole`    | `"0x8000000F"`                             | Raw eight-digit OLE value, including system or unknown values |

Supported system names are `scrollBar`, `desktop`, `activeCaption`,
`inactiveCaption`, `menu`, `window`, `windowFrame`, `menuText`, `windowText`,
`captionText`, `activeBorder`, `inactiveBorder`, `appWorkspace`, `highlight`,
`highlightText`, `buttonFace`, `buttonShadow`, `grayText`, `buttonText`,
`inactiveCaptionText`, `buttonHighlight`, `buttonDarkShadow`, `buttonLight`,
`infoText`, `infoBackground`, `hotLight`, `gradientActiveCaption`,
`gradientInactiveCaption`, `menuHighlight`, `menuBar`. Use exact spelling.

| Property           | Integer values                                               | Applies to                                                              |
| ------------------ | ------------------------------------------------------------ | ----------------------------------------------------------------------- |
| `borderStyle`      | 0 None, 1 Single                                             | Form, Frame, Label, TextBox, ComboBox, ListBox, Image                   |
| `textAlign`        | 1 Left, 2 Center, 3 Right                                    | Label, TextBox, ComboBox, ListBox, CheckBox, OptionButton, ToggleButton |
| `pictureAlignment` | 0 TopLeft, 1 TopRight, 2 Center, 3 BottomLeft, 4 BottomRight | Image only                                                              |
| `pictureSizeMode`  | 0 Clip, 1 Stretch, 3 Zoom                                    | Image only; 2 is invalid                                                |

Image accepts BackColor but not ForeColor. Page has none of these typed
appearance fields. CommandButton has no TextAlign. Consult generated property
rows before sharing a style across different control types.

## Pictures

```yaml
picture:
  path: src/forms/assets/logo.jpg
pictureAlignment: 2
pictureSizeMode: 3
```

These are Image fields. The asset must exist, be valid BMP/JPEG, and resolve
inside the project root, including through symlinks/junctions. Limits are 16 MiB
and 16 million pixels. Paths are relative to the project root, not the spec.
`picture: { remove: true }` clears a picture; omission preserves it. Empty
objects, mixed actions, `remove: false` and null are invalid.

## Existing Source and Captured Metadata

- `controls` is the full desired collection: dropping a control removes it.
  Preserve unrelated controls and stable IDs when adapting examples.
- Omitted properties preserve retained values. Empty strings, false and zero
  are edits. Partial font objects preserve omitted members. Omitted `tabs`
  preserves the collection; `tabs: []` clears it.
- `observed`, `unsupported` and persisted `warnings` describe captured state.
  Do not author them to request edits. New captures avoid duplicate authored
  values; existing observations remain accepted and should survive local edits.
- Preserve unrelated legacy `properties`. Use typed top-level keys for modeled
  fields; conflicting typed and legacy values are errors. Do not delete opaque
  metadata merely to bypass a compiler limitation.

Use structured `spec.issues[]`, warnings and compiler results to check the
request. Static validation does not prove rendering or runtime behavior;
follow the verification workflow in [forms.md](forms.md).
