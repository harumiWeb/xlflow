# UserForm Specification

`src/forms/specs/<FormName>.yaml` is the source-controlled Designer specification for a UserForm. `xlflow form build`, `push`, and `pack` use the shared pure-Go compiler for this contract. Excel-backed build imports the compiled form artifact into the target workbook and restores the previous component if import fails.

## Minimal Authoring Spec

`xlflow form new SampleForm` creates an authoring-focused starting point:

```yaml
schemaVersion: 1
kind: xlflow.userform
basis: designer
form:
  name: SampleForm
  caption: SampleForm
  width: 240
  height: 180
controls: []
```

Do not add snapshot-oriented fields such as `warnings` to a new spec unless they came from `form snapshot` or another capture workflow.

## Document Fields

| Field              | Type    | Required | Rules                                                          |
| ------------------ | ------- | -------- | -------------------------------------------------------------- |
| `schemaVersion`    | integer | Yes      | Must be `1`.                                                   |
| `kind`             | string  | Yes      | Must be `xlflow.userform`.                                     |
| `basis`            | string  | Yes      | Must be `designer`.                                            |
| `coordinateSystem` | string  | No       | `points` or `parent-relative`.                                 |
| `form`             | object  | Yes      | Form metadata and persisted root `DisplayedSize` dimensions.   |
| `controls`         | array   | Yes      | Flat array of control mappings.                                |
| `warnings`         | array   | No       | Snapshot-only capture metadata; omit from hand-authored specs. |

## Form Fields

| Field             | Type   | Required | Support                                                                                                                                     |
| ----------------- | ------ | -------- | ------------------------------------------------------------------------------------------------------------------------------------------- |
| `name`            | string | Yes      | Supported. VBA UserForm component name.                                                                                                     |
| `caption`         | string | No       | Supported.                                                                                                                                  |
| `width`, `height` | number | No       | Persisted MS-OFORMS `DisplayedSize` in points. New forms default to 240 × 180; omitted dimensions on existing template forms are preserved. |
| `observed`        | object | No       | Genuine snapshot observations, including measured outer and inside dimensions; not a source of authoring values.                            |

## Controls

`form.width` and `form.height` directly author persisted MS-OFORMS
`DisplayedSize`, in points. Snapshots and file pull seed these fields from the
persisted `DisplayedSize`; inspection fields named `client_width` and
`client_height` carry that value and are not read from VBFrame
`ClientWidth` / `ClientHeight`. VBFrame client properties are compiler-
maintained companion values and never seed root authoring dimensions. Runtime
MSForms `InsideWidth` / `InsideHeight` are separate observations and are not
converted to or from these values. `form.observed` can retain captured outer
`width`/`height` and `insideWidth`/`insideHeight` when available. It does not synthesize
`clientWidth`/`clientHeight` by copying the authored `form.width`/`height`.
Excel import/export may normalize
`DisplayedSize`, so snapshots report the actual persisted size and exact
round-trip equality with authored points is not guaranteed. The pure-Go
nearest-integer HIMETRIC rounding contract is separate. The legacy
`form.build` object and its
dimension aliases are invalid. The same pure-Go compiler handles form build,
both push backends, and pack. New-form construction defaults omitted axes to
240 × 180; edits to existing template forms preserve omitted axes. Excel build
compiles against a private copy of the live workbook, exports the compiled
component from a temporary Excel instance, then imports it into the target
through VBIDE with rollback from an exported backup. Canonical specs do not
require prebuilt `.frm` or `.frx` files.

Every authored control requires `id`, `name`, and `type`. Controls are a flat list; use `parentId` to place a child inside a container.

| Common field                     | Type                  | Support          | Notes                                                                                           |
| -------------------------------- | --------------------- | ---------------- | ----------------------------------------------------------------------------------------------- |
| `id`                             | string                | Supported        | Stable, unique identifier used by `parentId`.                                                   |
| `name`                           | string                | Supported        | VBA control name.                                                                               |
| `type`                           | string                | Supported        | One of the built-in types below, or a custom type with `progId`.                                |
| `progId`                         | string                | Supported        | Must match a known built-in `type` when both are known.                                         |
| `parentId`                       | string                | Supported        | Must reference another control ID and obey the built-in parent/child rules below.               |
| `zIndex`                         | integer               | Supported        | Sibling ordering hint.                                                                          |
| `left`, `top`, `width`, `height` | number                | Supported        | Designer coordinates in points; these are not authoring fields for `Page`.                      |
| `tabIndex`                       | integer               | Supported        | Designer tab order.                                                                             |
| `enabled`, `visible`             | boolean               | Supported        | Initial control state.                                                                          |
| `picture`                        | object                | Image only       | Project-relative BMP/JPEG asset path or explicit `{ remove: true }`.                            |
| `observed`, `unsupported`        | object / string array | Snapshot-only    | Genuine captured state; do not use as authoring fields or duplicate authored scalar/tab values. |
| `controls`                       | array                 | Snapshot-only    | Legacy nested form accepted for compatibility; prefer flat `controls` plus `parentId`.          |
| `properties`                     | object                | Custom/unchecked | Unchecked property bag; avoid it in hand-authored built-in controls.                            |

### Built-in Control Types

| Type            | Default `progId`        | Type-specific properties                                                         |
| --------------- | ----------------------- | -------------------------------------------------------------------------------- |
| `Label`         | `Forms.Label.1`         | `caption` (string)                                                               |
| `TextBox`       | `Forms.TextBox.1`       | `text` (string), `value` (any)                                                   |
| `ComboBox`      | `Forms.ComboBox.1`      | `text` (string), `value` (any), `list` (string array), `selectedIndex` (integer) |
| `ListBox`       | `Forms.ListBox.1`       | `text` (string), `value` (any), `list` (string array), `selectedIndex` (integer) |
| `CommandButton` | `Forms.CommandButton.1` | `caption` (string)                                                               |
| `CheckBox`      | `Forms.CheckBox.1`      | `caption` (string), `value` (any)                                                |
| `OptionButton`  | `Forms.OptionButton.1`  | `caption` (string), `value` (any)                                                |
| `ToggleButton`  | `Forms.ToggleButton.1`  | `caption` (string), `value` (any)                                                |
| `SpinButton`    | `Forms.SpinButton.1`    | `value` (integer)                                                                |
| `ScrollBar`     | `Forms.ScrollBar.1`     | `value` (integer)                                                                |
| `Image`         | `Forms.Image.1`         | Common fields; optional `picture` asset                                          |
| `Frame`         | `Forms.Frame.1`         | `caption` (string); the built-in container type                                  |
| `MultiPage`     | `Forms.MultiPage.1`     | `selectedIndex` (integer; zero-based Page index or `-1`)                         |
| `Page`          | `Forms.Page.1`          | `caption`, `controlTipText`, `tag`, `accelerator` (strings)                      |
| `TabStrip`      | `Forms.TabStrip.1`      | `selectedIndex` (integer; zero-based Tab index or `-1`), `tabs` (Tab array)      |

### Font, colors and appearance

FormSpec remains version 1. Applicable built-in controls and the form accept
partial `font` objects and typed colors:

```yaml
form:
  name: Main
  font: { name: Tahoma, size: 9, bold: false, charset: 128 }
  foreColor: { kind: rgb, value: "#203040" }
  backColor: { kind: system, value: buttonFace }
```

Font members are `name`, `size` (points), `bold`, `italic`, `underline`,
`strikeThrough` and numeric `charset`; omit members to retain existing values.
Page, Image, SpinButton and ScrollBar do not accept Font. Custom ActiveX and
individual Tab records do not gain these built-in typed fields.

`foreColor` and `backColor` use `{ kind, value }`. RGB is `#RRGGBB`, system
colors use names such as `buttonFace`, and raw OLE is `0xXXXXXXXX`, for example
`{ kind: ole, value: "0x8000000F" }`. System and unknown OLE values preserve
their meaning rather than being rewritten to current RGB. Image supports
BackColor only. Applicable controls expose `textAlign` (1 Left, 2 Center,
3 Right) and `borderStyle` through the canonical contract. Image additionally
exposes `pictureAlignment` (0 TopLeft, 1 TopRight, 2 Center, 3 BottomLeft,
4 BottomRight) and `pictureSizeMode` (0 Clip, 1 Stretch, 3 Zoom).

The built-in applicability matrix is:

| Target                               | Font | ForeColor | BackColor | BorderStyle | TextAlign |
| ------------------------------------ | ---- | --------- | --------- | ----------- | --------- |
| Form                                 | Yes  | Yes       | Yes       | Yes         | No        |
| Label, TextBox, ComboBox, ListBox    | Yes  | Yes       | Yes       | Yes         | Yes       |
| CheckBox, OptionButton, ToggleButton | Yes  | Yes       | Yes       | No          | Yes       |
| CommandButton                        | Yes  | Yes       | Yes       | No          | No        |
| Frame                                | Yes  | Yes       | Yes       | Yes         | No        |
| MultiPage, TabStrip                  | Yes  | Yes       | Yes       | No          | No        |
| Image                                | No   | No        | Yes       | Yes         | No        |
| SpinButton, ScrollBar                | No   | Yes       | Yes       | No          | No        |
| Page                                 | No   | No        | No        | No          | No        |

Page has none of these appearance properties, and CommandButton has no
TextAlign. Font size updates persisted as TextProps require 0.05 pt increments;
explicitly empty font names and non-ASCII updates to an existing StdFont are
rejected. Existing unsupported font records are preserved on no-op.

Legacy property-bag values remain readable. Confirming a typed property in
the Designer migrates only its matching legacy key; unrelated bag values are
retained. Equal typed/legacy values may coexist; conflicting values are errors.
Opening a dialog never fills omitted source values from rendering defaults.

### Image pictures

An Image may reference a BMP or JPEG file using a path relative to the xlflow
project root. Paths are checked after resolving filesystem aliases, including
symlinks and Windows junctions, and must remain inside the project. Assets are
limited to 16 MiB and 16 million pixels. The enclosing Designer stream retains
its 64 MiB limit.

```yaml
- id: brand-mark
  name: BrandMark
  type: Image
  picture:
    path: src/forms/assets/brand.jpg
```

Use `picture: { remove: true }` to explicitly clear an existing picture.
Omitting `picture` preserves an existing picture and leaves a new Image empty.
Null, empty values, `remove: false`, mixed `path` and `remove`, unknown keys,
and `picture` on non-Image controls are rejected. Picture source bytes are
referenced by path and are not embedded in the YAML or JSON spec.

The shared pure-Go compiler supports this field in `form build`, both push
backends, and pack. Excel-backed form build applies the compiled picture
through the bridge.
File pull emits supported embedded pictures under `src/forms/assets/` using
content-addressed BMP/JPEG names. Identical assets may be shared; pull does not
automatically delete old assets. File push loads, validates, and fingerprints
only referenced assets. Unreferenced files under `src/forms/assets/` are
retained and ignored.
Excel may normalize a JPEG input to a BMP picture resource when saving, so a
pulled asset can be BMP and need not be byte-identical to the original JPEG.

### MultiPage, Page, and TabStrip

The authored `controls` array stays flat. Use `parentId` to describe the
container hierarchy: a `MultiPage` may contain only `Page` controls, each
`Page` must belong to a `MultiPage`, and controls inside a MultiPage belong
under a `Page`. A `TabStrip` is independent: it contains no controls, and its
tabs are records in its own `tabs` array rather than entries in `controls`.

```yaml
controls:
  - id: steps
    name: Steps
    type: MultiPage
    selectedIndex: 0
  - id: identity
    parentId: steps
    name: Identity
    type: Page
    caption: Identity
  - id: email
    parentId: identity
    name: Email
    type: TextBox
  - id: navigation
    name: Navigation
    type: TabStrip
    selectedIndex: 0
    tabs:
      - name: Details
        caption: Details
        controlTipText: Show details
        tag: section-details
        accelerator: D
        enabled: true
        visible: true
```

`TabStrip.tabs` is its authored collection. Each tab requires `name`; names
must be unique within that TabStrip, ignoring case. Optional fields are
`caption`, `controlTipText`, `tag`, and `accelerator` (strings), plus `enabled`
and `visible` (booleans). These fields belong to the tab record, not to the
TabStrip control. When editing an existing form, omitting `tabs` leaves the
collection unspecified; author `tabs: []` to remove all tabs. An authored array
defines the desired collection and order.

`MultiPage.selectedIndex` addresses its Pages collection, while
`TabStrip.selectedIndex` addresses its independent Tabs collection. Both are
zero-based; `-1` means no logical selection. An empty collection is represented
with `selectedIndex: -1` in the normalized UserForm model, including an empty
MultiPage whose raw Excel selection value may be stale. Inspection snapshots
use `selected_index`; authored FormSpec uses `selectedIndex`. MultiPage
snapshots expose Pages as child controls, while TabStrip snapshots expose
independent tab metadata in `tabs` and no child controls.

Page bounds are derived from the owning MultiPage. Page `left`, `top`, `width`,
and `height` are not authoring properties; if inspection captures them, they
are snapshot-only values under `observed`. New pure-Go generation lays out
standard top tabs using fixed 96-DPI geometry and an explicit Tahoma 8.25pt
font, independent of the build host's display and font settings.

File projections and Excel snapshots retain captured Page geometry and form
outer/inside measurements. They do not add observed copies of authored
captions, control scalar properties, TabStrip tab data, or client dimensions.
Existing `observed` values in YAML, YML, and JSON remain accepted and are kept
by localized source edits; they are never promoted into authoring properties.

Modeled UserForm edits are all-or-nothing. The Go compiler validates the
requested state and applies it to an independent model; a stale, invalid, or
unsupported edit returns an error without changing the input model. Callers
publish only a fully compiled result.

`ComboBox` and `ListBox` `list` and `selectedIndex` are observed-only: xlflow attempts to apply them, but round-trip fidelity is not guaranteed. `MultiPage` and `TabStrip` `selectedIndex` are modeled authoring fields.

### Parent and Custom-Control Rules

- IDs must be unique; `parentId` must resolve to an existing ID.
- A control cannot parent itself, and parent references cannot form a cycle.
- A built-in child can only use a permitted built-in container (`Frame`,
  `Page`, or `MultiPage`) as its parent. `MultiPage` accepts only `Page`
  children, and `Page` accepts controls but can only be parented by
  `MultiPage`.
- `TabStrip` does not contain controls; model tabs with its `tabs` property.
- A custom `type` requires an explicit custom `progId`. xlflow validates common fields but emits a `custom/unchecked` warning because type-specific properties cannot be verified.

## Support Levels

| Level            | Meaning                                                                                    |
| ---------------- | ------------------------------------------------------------------------------------------ |
| Supported        | xlflow validates and applies the field as part of the normal Designer build.               |
| Best-effort      | xlflow attempts to apply it; inspect the rebuilt form to confirm the result.               |
| Observed-only    | Captured list state that may be applied best-effort but is not guaranteed to round-trip.   |
| Snapshot-only    | Capture metadata, not normal authored build intent.                                        |
| Custom/unchecked | Accepted for compatibility, but xlflow cannot validate detailed control-specific behavior. |

## Validation and Editor Diagnostics

`form build` reports all detected contract issues before Excel opens. YAML files directly under the configured `src/forms/specs` directory receive the same live diagnostics in the xlflow LSP.

The LSP also provides context-aware completion and Hover for known UserForm YAML fields, built-in control types, and built-in ProgIDs. Hover shows the expected value type, required status, applicable controls, and support level. In particular, root `width` and `height` author persisted MS-OFORMS `DisplayedSize`; `ComboBox` and `ListBox` list selection state is observed-only; and `warnings`, `observed`, `unsupported`, and `properties` are snapshot-oriented or custom/unchecked metadata rather than normal build inputs. Snapshot-only support metadata does not produce a generic `UFV013` warning by presence alone; real custom, observed-only, and best-effort warnings remain. A persisted warning entry with a message is reported with its original code and message at `warnings[i].message`; the mere presence of `warnings`, `observed`, or `unsupported` containers is not a warning.

| Code              | Meaning                                                                                                                            |
| ----------------- | ---------------------------------------------------------------------------------------------------------------------------------- |
| `UFY001`          | YAML parse error.                                                                                                                  |
| `UFV001`–`UFV005` | Unknown field, invalid value type/fixed value, missing required field, or unsupported property.                                    |
| `UFV006`–`UFV012` | Unsupported control type, duplicate ID, invalid parent reference, parent cycle, invalid parent type, or type/ProgID mismatch.      |
| `UFV013`–`UFV015` | Support-level warning (generic snapshot-only presence is suppressed), custom-control warning, or unresolved snapshot-only control. |
| `UFV016`          | Invalid form dimension (finite, non-negative, representable HIMETRIC required).                                                    |

`UFV021` reports invalid typed appearance values or conflicting typed/legacy
representations. `UFV019` remains reserved for Tab validation.

The removed `form.build` aliases are reported as unknown fields (`UFV001`).

For a Designer capture, use `xlflow form snapshot <FormName> --out src/forms/specs/<FormName>.yaml`. Captured `warnings`, `observed`, and other snapshot fields are preserved for review, but new authoring should begin with the minimal form above. New projections do not duplicate authored values under `observed`; existing observed source values are preserved by local edits and never become authoring values. Keep authored specs directly in `src/forms/specs/` so the LSP recognizes them; files outside that configured location are intentionally ignored.

## Related

- [form command](../commands/form)
- [Project structure](./project-structure)
- [LSP](../commands/lsp)
