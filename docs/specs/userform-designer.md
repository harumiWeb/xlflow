# VS Code UserForm Designer

This specification defines the read-only UserForm Designer preview in the
xlflow Visual Studio Code extension. The VS Code `TextDocument` remains the
canonical source. The Designer renders the current document through the xlflow
Language Server Protocol (LSP); it does not create or persist another source
representation.

## Editor selection and document eligibility

The extension registers `xlflow.userFormDesigner` as an optional custom text
editor for `.yaml`, `.yml`, and `.json` files. Its `option` priority keeps the
normal text editor available. Users can choose the visual editor with **Reopen
Editor With...**.

The preview is available only for a FormSpec directly inside the `specs`
directory under the configured forms root (`[src].forms` in `xlflow.toml`,
defaulting to `src/forms`). Arbitrary YAML, YAML nested in another directory,
and other JSON files are not UserForm documents. The file must belong to an
xlflow project with `xlflow.toml` and be handled by the currently selected
project's single xlflow LSP client. Opening a document associated with another
project, or without a compatible LSP connection, produces an editor error.

## Source and preview protocol

The extension host reads source from the open VS Code `TextDocument`, including
unsaved changes. It sends the document URI, its current version, and the full
text to the LSP. The Webview never reads the document or workspace files
directly.

The server advertises preview support through
`capabilities.experimental.userFormPreview: true`. The request is:

```json
{
  "uri": "file:///project/src/forms/specs/Main.yaml",
  "version": 12,
  "text": "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: Main\ncontrols: []\n"
}
```

The `xlflow/userFormPreview` response always includes the requested `version`
and may include a parsed canonical `document`, validation `warnings`, or an
`error`:

```json
{
  "version": 12,
  "document": { "form": { "name": "Main" }, "controls": [] },
  "warnings": []
}
```

An error has a stable `code` and a human-readable `message`; `line` and `column`
are included when the parser can identify a source location. The language
server validates the URI and FormSpec kind and parses the supplied text with
xlflow's canonical FormSpec parser. It does not replace the supplied text with
saved file contents or invoke Excel.

The extension refreshes the preview when the document changes or the LSP
connection becomes available again. Results whose document version or
connection generation is no longer current are discarded. Warnings are shown
alongside a valid preview. If parsing, validation, project selection, or the
LSP request fails, the Webview displays the error and retains the last valid
rendered document. A later valid document automatically replaces that state.

The host-to-Webview protocol carries a rendered `document`, an
`invalidDocument` status, or a `themeChanged` notification. The Webview can
signal that it is ready and can ask VS Code to open the backing document in the
normal text editor. It does not send document edits.

## Read-only rendering

The Webview uses Preact and TypeScript to render the FormSpec form and its
controls. It displays the built-in controls supported by the FormSpec model,
including Label, TextBox, CommandButton, CheckBox, OptionButton, ComboBox,
ListBox, ToggleButton, SpinButton, ScrollBar, Image, Frame, MultiPage, Page,
and TabStrip. Unknown control types receive a textual placeholder. Control
positions and sizes are an approximate visual representation, not a pixel
accurate reproduction of the VBA designer.

Dimensions are represented in points and displayed at 96/72 CSS pixels per
point. Missing control positions default to zero. Missing control sizes use
these display defaults, in points:

| Control                                     | Width × height |
| ------------------------------------------- | -------------: |
| Label, CheckBox, OptionButton, ToggleButton |        72 × 18 |
| TextBox, ComboBox                           |       120 × 18 |
| CommandButton                               |        72 × 24 |
| ListBox                                     |       120 × 72 |
| SpinButton                                  |        18 × 36 |
| ScrollBar                                   |       120 × 18 |
| Image                                       |        72 × 72 |
| Frame                                       |      144 × 108 |
| MultiPage                                   |      240 × 180 |
| TabStrip                                    |       240 × 48 |
| Unknown control                             |        72 × 24 |

If the form has no usable dimensions, its display size defaults to 240 × 180
points. The preview marks dimensions supplied by these defaults as approximate.
Page content is laid out inside its parent MultiPage; the Page is not rendered
as an independently positioned top-level control. The selected Page from the
FormSpec is displayed, defaulting to the first Page when no selection is
recorded.

Explicitly invisible controls and their descendants are not rendered. Invisible
Page tabs are omitted, but selection indexes retain the full ordered Page
collection. A selected invisible Page does not fall through to a different
Page. Omitted visibility defaults to visible.

Page collection order follows stable sibling `zIndex` order, matching the
pure-Go generation and topology compiler. Selection is applied before hiding
invisible Pages. Explicit build client dimensions take precedence over observed
client/inside dimensions, which take precedence over outer dimensions. Either
axis relying on outer dimensions remains approximate. ComboBox display prefers
explicit text, then value (including empty strings), then the first list item.

The extension host supplies localized UI strings through a `localization`
message when the Webview reports readiness. These are rendered as text; parser
diagnostics and workspace-authored captions retain their original wording.

The Designer does not edit the FormSpec, provide zoom controls, allow Page
selection, or load image assets. Image controls use a placeholder even when a
picture path is present in the FormSpec.

## Webview security

Each Webview document uses a fresh random nonce for its bundled script. Its
Content Security Policy sets `default-src 'none'`, permits scripts only with
that nonce, and permits styles only from the VS Code Webview resource origin.
The Webview's local resource root is limited to the packaged Designer bundle.
Workspace-provided strings are passed as data to the renderer and are not
interpolated into the HTML document.

## Related

- Issue #915
- `docs/adr/ADR-0064-userform-designer-text-document-boundary.md`
- `docs/specs/ms-oforms.md`
- `docs/specs/userform-picture-assets.md`
