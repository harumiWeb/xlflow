# UserForm Typed Property Editors

This supplements the scalar Property Grid, source-preserving semantic edit,
MS-OFORMS and picture-asset specifications. ADR-0074 records the decisions.
The running Go contract is authoritative for applicable fields and validation;
a Webview registry maps editor kinds to UI components only.

## Negotiation and authored values

The server advertises `capabilities.experimental.userFormPropertyEditors: 1`.
A new client opts into extended `xlflow/userFormPreview` metadata with
`propertyEditorVersion: 1`. Descriptors add `editor` and child descriptors to
existing field/type/presence metadata. Without opt-in, return scalar metadata;
do not expose an object to a scalar-only client. Missing capability or missing
metadata disables the additional editor rather than guessing a schema.

Keep FormSpec `schemaVersion: 1`. Values come from source syntax rather than
normalized rendering defaults. Preserve absent fields, explicit null, empty
strings, false and zero. Optional null is not deletion, and opening an editor
must not materialize default values. Unknown or ambiguous source syntax must
not produce guessed editable state. Custom ActiveX and individual Tab records
do not acquire typed built-in editors.

## Font and Color source forms

```yaml
form:
  name: Main
  font:
    name: Tahoma
    size: 9
    bold: false
    charset: 128
  foreColor: { kind: rgb, value: "#203040" }
  backColor: { kind: system, value: buttonFace }
controls:
  - id: heading
    name: Heading
    type: Label
    font: { italic: true, underline: false, strikeThrough: false }
    foreColor: { kind: ole, value: "0x8000000F" }
    textAlign: 2
```

`font` is partial: `name`, `size` in points, `bold`, `italic`, `underline`,
`strikeThrough`, and integer `charset`. Omitted members preserve persisted
values. Numeric charset is authoritative; named choices are conveniences and
unlisted valid numbers remain representable. Font selection offers the current
value, representative candidates and free input; it does not enumerate OS
fonts. Page, Image, SpinButton and ScrollBar have no Font editor.

Colors are discriminated objects: RGB uses `#RRGGBB`, system uses its supported
symbolic name, and OLE uses an eight-digit `0xXXXXXXXX` string. The raw mode
retains unknown OLE values. Resolving a system color for rendering must never
write its current RGB approximation back. Image exposes BackColor only.
Form and Frame ForeColor do not provide inherited values for existing child
controls. Each control's preview uses its own authored color or its independent
display default; CSS inheritance must not propagate a container's text color.
These display defaults must never be written into FormSpec source.
The same boundary applies to Font: a container's family, size, bold, italic,
underline or strike-through must not change existing child controls. A partial
child Font uses independent display defaults for unspecified members, rather
than inheriting the parent's font through CSS. Its authored members still win.
Palette/System tabs offer 48 RGB swatches and system-color choices from the
server metadata. Switching tabs alone leaves the authored value unchanged;
selecting a swatch changes only the modal draft until OK. Arrow keys move focus
within the RGB palette; direct RGB and raw OLE input remain available.
`textAlign` uses 1/2/3 for Left/Center/Right where the contract permits it;
`borderStyle` uses the contract's numeric choices. Both use inline ComboBoxes.
Boolean properties use checkboxes. State/type choices remain separate from
editor launchers, and simple inputs do not display a dedicated editor ellipsis.

Legacy property-bag representations remain readable. On confirmation, insert
or update the typed field and remove only that property's matching legacy key
in the same transaction. Leave every unrelated bag entry unchanged. Equivalent
typed and legacy values may coexist; conflicting values are validation errors.
Do not perform bulk migration by opening or saving a dialog.

### Applicable built-in properties

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

The Go contract enforces this matrix. Page exposes no Font, ForeColor,
BackColor or BorderStyle: focused runtime VBA `CallByName` observations
reported error 438 for those properties. CommandButton has no TextAlign.
Individual Tab records and custom ActiveX controls are excluded. MultiPage's
font is persisted through its hidden TabStrip, not through its Page records.

## Editor lifecycle

A dedicated modal binds to the source URI, selected target ID and version.
It provides initial focus, a focus trap, focus restoration, Escape Cancel and
Enter OK, excluding IME composition and keys handled by child controls. Labels
for property names remain English under every UI language; categories, help,
errors and actions follow localization and VS Code theme tokens.

Font and Color render a sample and temporary Canvas overlay as input changes.
This overlay is separate from the source-text Designer draft, save state and
native Undo history. Opening or Cancel causes zero semantic edits. OK sends
one validated semantic transaction and produces one native Designer Undo step.
A failed commit keeps the draft and diagnostics visible. A pending edit blocks
opening another property editor. Explicit Designer Save continues to publish
source according to ADR-0073; modal OK does not publish YAML/JSON.

An external source update, target change or LSP restart invalidates the bound
version. Retain the modal input, remove the temporary Canvas overlay, and
disable OK. Cancel and reopen to obtain fresh state; never automatically rebase
an unconfirmed modal. The browser's fallback fonts and system-color rendering
are approximate and must be presented as previews, not exact Excel output.

## Position transaction

`setGeometry` carries optional `left`, `top`, `width`, and `height`, with at
least one changed axis. Optional `controlId` targets a single control; omission
targets the form and permits only `width` and `height`. Page and multiple
selection are excluded. All supplied axes are validated and committed together;
unchanged axes remain absent from the operation and source. One confirmation
is one native undoable action, including when several axes change.

## Image editor

The extension uses a native BMP/JPEG file picker. Selection is staged for modal
and Canvas preview; the Webview cannot read arbitrary filesystem paths. A
project-contained asset uses its existing project-relative path. An external
asset is copied only on OK, after canonical image validation, into the configured
`<forms-root>/assets/<full-lowercase-SHA256>.bmp|jpg`. Hash the image bytes,
reuse identical content and reject a different-content collision. Resolve
symlinks and junctions before root-containment checks. Cancel creates no asset.

The source edit follows successful staging/import. If confirmation fails,
remove only a newly created asset that is still unreferenced. Never remove a
pre-existing or shared asset. Retain successful assets after Undo so Redo can
resolve them. Explicit Remove emits `picture: { remove: true }`; absence, null
and Cancel never imply removal. PNG/GIF, custom ActiveX pictures and pictures
on buttons/labels are excluded. `pictureAlignment` offers TopLeft (0), TopRight
(1), Center (2), BottomLeft (3), BottomRight (4); `pictureSizeMode` offers Clip
(0), Stretch (1), Zoom (3). PicturePosition is not part of Image editing.

Rollback reference checks include unsaved text buffers, open Designer history
and closed Designer checkpoints. Pending checkpoint writes, changes to the
checkpoint revision during a scan, malformed recovery data and read failures
retain the asset. Recheck references after asynchronous file-identity checks
before deletion.

## Persistence and verification

Font StdFont/TextProps, colors, alignment and image changes flow through the
shared pure-Go compiler used by pack, both push backends and form build/apply.
Preserve unrelated record fields, unedited font data and unknown OLE values;
fail unsafe unsupported mutation rather than replacing opaque binary state.
TextProps FontHeight uses whole twips (0.05 pt), independent of machine DPI.
Reject a requested size that cannot be represented in those units rather than
silently rounding. StdFont uses 1/10000 pt storage; its face-name updates
require ASCII. Reject explicitly empty font names and non-ASCII StdFont name
updates before publication. Existing opaque or unsupported font records remain
intact on no-op; these update restrictions do not authorize rebuilding them.

TextProps Bold confirmation must update FontWeight (700 for bold, 400 for
regular) together with the FontEffects bold bit. Excel was observed retaining
Weight 400 and reporting regular when only the effect bit was set. When an
explicit weight is present, projection derives Bold from that weight instead
of treating the bit alone as Excel compatibility evidence.

Appearance validation uses `UFV021`; `UFV019` remains reserved for Tab
validation. No-op serialization retains the existing lossless contract.
There is no per-property imperative COM writer fallback.

Verification must cover YAML/YML/JSON comments and formatting, partial typed
members, legacy migration/conflicts, explicit source states, malformed or
ambiguous edits, precision, stale versions and old clients/servers. UI checks
cover zero edits on open/Cancel, one transaction on OK, Undo/Redo, stale draft
retention, preview clearing, focus/IME, Japanese/English and narrow/theme
layouts. Asset checks cover import/reuse/collision, root escape, corrupt input,
failure cleanup and Undo/Redo retention. Focused Windows/Excel checks must
separately establish actual Font/Color/Image persistence and existing-form
preservation through file push and pack; pure source tests make no such claim.

## Related

- [Designer](userform-designer.md)
- [Semantic edits](userform-semantic-edit.md)
- [Picture assets](userform-picture-assets.md)
- [MS-OFORMS](ms-oforms.md)
