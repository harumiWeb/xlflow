# UserForm Specification

`src/forms/specs/<FormName>.yaml` is the source-controlled Designer specification for a UserForm. `xlflow form build` validates this contract before opening Excel.

## Minimal Authoring Spec

`xlflow form new SampleForm` creates an authoring-focused starting point:

```yaml
schemaVersion: 1
kind: xlflow.userform
basis: designer
form:
  name: SampleForm
  caption: SampleForm
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
| `form`             | object  | Yes      | Form metadata and build intent.                                |
| `controls`         | array   | Yes      | Flat array of control mappings.                                |
| `warnings`         | array   | No       | Snapshot-only capture metadata; omit from hand-authored specs. |

## Form Fields

| Field             | Type   | Required | Support                                                                                    |
| ----------------- | ------ | -------- | ------------------------------------------------------------------------------------------ |
| `name`            | string | Yes      | Supported. VBA UserForm component name.                                                    |
| `caption`         | string | No       | Supported.                                                                                 |
| `width`, `height` | number | No       | Best-effort. Verify the rebuilt form in Excel.                                             |
| `build`           | object | No       | Supported build intent: `caption`, `width`, `height`.                                      |
| `observed`        | object | No       | Snapshot-only captured state: `caption`, `width`, `height`, `insideWidth`, `insideHeight`. |

## Controls

Every authored control requires `id`, `name`, and `type`. Controls are a flat list; use `parentId` to place a child inside a container.

| Common field                     | Type                  | Support          | Notes                                                                                  |
| -------------------------------- | --------------------- | ---------------- | -------------------------------------------------------------------------------------- |
| `id`                             | string                | Supported        | Stable, unique identifier used by `parentId`.                                          |
| `name`                           | string                | Supported        | VBA control name.                                                                      |
| `type`                           | string                | Supported        | One of the built-in types below, or a custom type with `progId`.                       |
| `progId`                         | string                | Supported        | Must match a known built-in `type` when both are known.                                |
| `parentId`                       | string                | Supported        | Must reference another control ID; only `Frame` is a built-in container.               |
| `zIndex`                         | integer               | Supported        | Sibling ordering hint.                                                                 |
| `left`, `top`, `width`, `height` | number                | Supported        | Designer coordinates in points.                                                        |
| `tabIndex`                       | integer               | Supported        | Designer tab order.                                                                    |
| `enabled`, `visible`             | boolean               | Supported        | Initial control state.                                                                 |
| `observed`, `unsupported`        | object / string array | Snapshot-only    | Captured state; do not use as normal authoring fields.                                 |
| `controls`                       | array                 | Snapshot-only    | Legacy nested form accepted for compatibility; prefer flat `controls` plus `parentId`. |
| `properties`                     | object                | Custom/unchecked | Unchecked property bag; avoid it in hand-authored built-in controls.                   |

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
| `Frame`         | `Forms.Frame.1`         | `caption` (string); the built-in container type                                  |

`ComboBox` and `ListBox` `list` and `selectedIndex` are observed-only: xlflow attempts to apply them, but round-trip fidelity is not guaranteed.

### Parent and Custom-Control Rules

- IDs must be unique; `parentId` must resolve to an existing ID.
- A control cannot parent itself, and parent references cannot form a cycle.
- A built-in child can only use a built-in container (`Frame`) as its parent.
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

The LSP also provides context-aware completion and Hover for known UserForm YAML fields, built-in control types, and built-in ProgIDs. Hover shows the expected value type, required status, applicable controls, support level, and build limitations. In particular, `width` and `height` are best-effort; `list` and `selectedIndex` are observed-only state that may be applied best-effort; and `warnings`, `observed`, `unsupported`, and `properties` are snapshot-oriented or custom/unchecked metadata rather than guaranteed normal build inputs.

| Code              | Meaning                                                                                                                       |
| ----------------- | ----------------------------------------------------------------------------------------------------------------------------- |
| `UFY001`          | YAML parse error.                                                                                                             |
| `UFV001`–`UFV005` | Unknown field, invalid value type/fixed value, missing required field, or unsupported property.                               |
| `UFV006`–`UFV012` | Unsupported control type, duplicate ID, invalid parent reference, parent cycle, invalid parent type, or type/ProgID mismatch. |
| `UFV013`–`UFV014` | Support-level warning or custom-control validation warning.                                                                   |

For a Designer capture, use `xlflow form snapshot <FormName> --out src/forms/specs/<FormName>.yaml`. Captured `warnings`, `observed`, and other snapshot fields are preserved for review, but new authoring should begin with the minimal form above. Keep authored specs directly in `src/forms/specs/` so the LSP recognizes them; files outside that configured location are intentionally ignored.

## Related

- [form command](../commands/form)
- [Project structure](./project-structure)
- [LSP](../commands/lsp)
