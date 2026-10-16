# Changelog

## Unreleased

- Fixed the invisible Designer Undo icon; Undo and Redo now both use bundled
  icon glyphs.

- Fixed native Designer double-clicks being ignored after pointer capture
  retargeted them to the viewport.

- Added default-event navigation by double-click and **Open event…** selection
  in Designer, with existing-handler navigation and native undoable code
  generation in the configured sidecar or frm source.

- Added **Open in Designer** to the FormSpec editor context menu, native
  Undo/Redo toolbar buttons, and Backspace deletion in Canvas and Structure.
- Added individual and all-form sidebar builds, with targeted source saving,
  workbook overwrite confirmation and sequential session-backed execution.

- Simplified UserForm Designer property inputs by removing type and null
  selectors. Existing absent/null fields remain unchanged until a concrete edit;
  false, zero and empty strings remain valid, and FormSpec/API null support is
  unchanged.

- Fixed child Label/control colors and fonts changing with Form or Frame
  ForeColor/Font in the Designer preview when Excel keeps independent settings.

- Added Palette/System tabs to UserForm Color editors, with 48 RGB swatches,
  system-color previews and keyboard navigation.

- Added opt-in typed UserForm property editors for partial Font, RGB/system/OLE
  Color, Image BMP/JPEG assets and single-target Position. Modal and Canvas
  previews stay outside draft history until OK; stale input remains visible
  with confirmation disabled. Image imports occur only on confirmation and
  retain content-addressed assets for Undo/Redo. Go/LSP metadata remains the
  authority, with scalar-client compatibility and edited-property-only legacy
  migration.

- UserForm Designer property names stay English regardless of the UI language.

- Refined the VS Code UserForm Designer with a compact contextual toolbar,
  grouped Toolbox, two-column Property Grid, Properties/Structure/Tab order
  tabs, independently scrolling panes, and compact layouts below 900 px.
  Structure selects nested and overlapping controls without changing source
  hierarchy.

- New Toolbox MultiPages start with two Pages and `selectedIndex: 0`; their
  control and Pages are created atomically as one native Undo step. New
  TabStrips start with `Tab1`, `Tab2`, and `selectedIndex: 0`. MultiPage and
  TabStrip insertion also require their respective Page or Tab capability;
  existing authored empty collections remain unchanged.

- Fixed false source conflicts during Designer Save by matching native source
  line endings and adopting successfully saved formatter output.

- Added URI-keyed deferred-save UserForm Designer drafts shared across panels.
  Native Save, Save All, menu Save, and AutoSave checkpoint drafts in extension
  storage without writing YAML/YML/JSON; the Designer Save button and scoped
  Ctrl+S/Cmd+S command explicitly publish the validated draft. The native tab
  dirty dot tracks its checkpoint independently from the source-not-applied
  badge. Non-overlapping text-editor changes merge, while conflicts retain both
  versions and require explicit resolution. Global AutoSave settings remain
  unchanged.

- Added selected-form right/bottom edge and corner dragging to resize
  `form.width` / `form.height` in client points, with snapping, child bounds,
  cancellation and atomic Undo/Redo.

- Unified UserForm authoring and preview dimensions as `form.width` and
  `form.height`, persisted client content size in points. Removed the unreleased
  `form.build` aliases; validation recommends the canonical fields, and observed
  snapshot dimensions no longer supply preview defaults.

- Fixed control-interior clicks and subsequent drags losing selection in the
  UserForm Designer because leaf controls intercepted clicks as blank containers.

- Added sibling multi-selection, marquee selection, group movement,
  alignment/distribution, same-size and z-order commands, shared Cut/Copy/Paste
  and Duplicate, and parent-local tab-order editing. Actions negotiate
  `userFormProductivityEdit` and preserve native Undo/Redo transactions.

- Added Page-list add/delete/rename/caption/reorder actions, transient active
  Page navigation, control placement and movement between Pages, and TabStrip
  tab collection/property editing. YAML/JSON edits preserve source, explicit
  selection identity and native Undo/Redo. The running LSP must advertise
  `userFormPageEdit` or `userFormTabEdit` for the corresponding actions;
  Page placement/reparenting also retains existing capability requirements.

- Added Frame placement, drag reparenting and clickable hierarchy breadcrumbs.
  Drop previews preserve apparent positions with boundary clamping, reject
  invalid destinations, and commit YAML/JSON parent and coordinate updates as
  one undoable edit. Older servers retain root placement and existing movement.
- Clipped Frame children to their content bounds so overflow cannot become an
  unintended root drag target.
- Reject custom-ProgID Frame destinations before placement or reparenting is
  dispatched, matching the server's built-in Frame contract.

- Added a UserForm Designer Toolbox for click-to-place root controls and
  structural deletion, with immediate leaf removal and confirmation before
  removing descendants. YAML/JSON edits use a separate LSP capability, stale
  document guards, and native undo/redo.
  Canvas Delete ignores focused Page-tab buttons to protect the previous
  control selection when keyboard focus moves to a tab.

- Added a UserForm Property Grid driven by the running LSP's field metadata.
  Select the form, a control or a Page tab to edit supported authored scalars
  in YAML/YML/JSON with one native undoable edit per confirmation. Unset fields,
  explicit null, empty strings, false and zero remain distinct; invalid input
  retains its field error without changing source. Property editing requires
  the new capability; existing geometry support remains available on older
  geometry-capable servers.
  Edit results survive hiding and restoring the same Designer panel, so an
  interrupted property edit cannot leave it stuck waiting for a response.
  Property-only server capabilities work independently of geometry support;
  numbers that would lose precision fall back to source editing.
  Form drafts remain separate from a control with ID `form`, and custom ProgID
  controls retain supported common scalar fields.

- Added UserForm selection, mouse movement, eight-direction resize, grouped
  arrow-key movement, 8-point grid/snapping toggles, and 50–200% zoom. Completed
  YAML/YML and JSON interactions create one localized, native undoable edit;
  parent bounds and container children constrain layout operations.
- UserForm Designer prefers measured client dimensions, displays explicit ComboBox
  values, and localizes its preview UI. Page selection retains compiler zIndex ordering.

- Added an optional UserForm Designer for canonical YAML, YML, and
  JSON FormSpecs. It previews the current TextDocument through xlflow LSP and
  retains the last valid rendering while the source is invalid.
  Button captions are centered horizontally and vertically, including multiline text.
  Buttons use classic raised edges, and ComboBox previews have a recessed field
  with a full-height raised dropdown button.

## v0.12.0

- Added UserForm Image `picture` field intelligence for project-relative BMP
  and JPEG assets, including explicit picture removal.

## v0.11.0

- Accept xlflow capability schemas v1 through v3, including backend-specific
  Excel requirements and dynamic `pull --backend auto` selection metadata.

## v0.10.2

- Fixed inline-suppression Quick Fix discovery with the current v2 rule
  metadata while retaining fail-closed behavior for unsupported schemas.

## v0.10.1

- Keep completion, signature help, hover, and definition responsive while Full
  diagnostics and workspace semantic indexing run in the language server.

- Publish Full diagnostics for giant VBA modules by reusing revision-scoped
  analysis indexes and bounded procedure parallelism in the language server.

- Keep language features responsive after editing a large VBA file when
  diagnostics are still reading its previous revision.

- Keep automatic server code actions from blocking language features on large
  files, with cancellation and stale-document protection in the matching LSP.

- Start the language client without a blocking duplicate `xlflow version`
  preflight and keep update, capability, project, and sidebar refreshes off the
  activation critical path. Opt-in startup logs correlate LSP lifecycle stages
  and the first usable hover, definition, and completion results.

## v0.10.0

- Fixed VBA hover documentation for procedures in currently open modules.

## v0.9.0

- Load static-analysis rule metadata from `xlflow rules --json` and offer inline-suppression Quick Fixes only for rules declared suppressible by the active CLI. Metadata discovery is cached per resolved CLI executable and fails closed when unavailable or incompatible.
- Link xlflow diagnostics in the Problems view to their generated rule documentation through LSP code descriptions.

## v0.8.1

- Minor fixes accompanying updates to the xlflow CLI

## v0.8.0

- Added the `xlflow.lsp.performanceLogging` setting to opt into structured language server performance measurements.
- Restart the language server when `xlflow.toml` is created, changed, or deleted so project configuration changes apply immediately.

## v0.7.0

- Added discovery of `xlflow capabilities --json` so the extension can use the
  CLI's central command-safety metadata instead of maintaining a local policy
  table. Workbook-busy operations now show their active command details and
  offer Retry or a 30-second Wait action when the CLI reports it is retryable.

## v0.6.1

- Fixed `xlflow module install --push` from the VS Code extension so helper source files appear in the sidebar even when push is blocked by preflight lint, and clarified the failure notification.

## v0.6.0

- Added VBA annotation completion for xlflow `@ExpectedError(...)`, `@Skip(...)`, and `@Todo(...)` test metadata when typing `@` in apostrophe comments.
- Added VS Code editor support for xlflow VBA documentation comments, including Quick Fix snippet generation from `'''`, Rubberduck `@Description` annotation completions in comments, doc-comment continuation on Enter, and highlighting for `'''` comments and Rubberduck description annotations.
- Updated Test Explorer to use xlflow's stable qualified VBA test IDs and run focused tests with `xlflow test --filter Module.TestName`.

## v0.5.0

- Added editor diagnostics for high-signal analyze warnings, including `VBA201`, `VBA204`, `VBA208`, `VBA209`, and `VBA212`.
- Suppressed Format Document failure notifications for incomplete or syntactically invalid VBA buffers; formatting now leaves the document unchanged instead.
- Added a VS Code Formulas sidebar view for discovering generated formula snapshots, opening `names.jsonl` and per-sheet region JSONL files, and running `xlflow formulas pull`.

## v0.4.0

- Added setup warnings for disabled Excel VBA object model access and for `.bas`, `.cls`, or `.frm` files that are not opened with xlflow's `vba` language mode.
- Added a Project view session action to connect xlflow to an already-open workbook with `xlflow session attach`.
- Open the configured workbook from the Project view in the associated desktop app instead of opening the binary workbook inside VS Code.

## v0.3.0

- Added Quick Fix actions for xlflow diagnostics to insert `xlflow:disable-next-line` or `xlflow:disable-line` suppression comments.
- Offer a Push & Run choice before VS Code macro/procedure runs when source files are newer than the workbook.
- Run macros from VS Code with `xlflow run --interactive` so user-triggered macro execution uses interactive mode explicitly.
- Run user-triggered macro and procedure commands in the VS Code terminal instead of the Output panel.
- Keep the xlflow Output panel in the background for routine commands; diagnostics still open it explicitly, and failures offer an Open Output action.

## v0.2.0

- Added a visible pulse icon for the `Run Doctor` setup action in the VS Code sidebar.
- Renamed the Japanese VS Code setup action for existing workbook initialization to clarify that it converts a workbook into an xlflow project instead of wiping the workbook.
- Added xlflow CLI update notifications, a Project view update indicator, and a command palette action to check for updates manually.
- Added sidebar save actions so the Project view toolbar and `Save required` indicator can run `xlflow save` when a session workbook needs saving.

## 0.1.0

- Added CodeLens commands for running no-argument VBA procedures and tests through `xlflow.runProcedure` and `xlflow.runTestProcedure`.
- Added CodeLens and save-before-run settings forwarded to `xlflow lsp --stdio` at startup.
- Added command palette actions for `xlflow skill install` and `xlflow module install`.
- Added Modules TreeView context menu actions for opening, renaming, deleting, revealing, copying, and running module/procedure items.
- Added UserForms TreeView context menu actions for opening, renaming, deleting, revealing, and copying UserForm source artifacts.
- Added the initial production VS Code extension for xlflow.
- Registered `.bas`, `.cls`, and `.frm` as VBA files.
- Added a thin `vscode-languageclient` client for `xlflow lsp --stdio`.
- Added command palette actions for restarting the language server, checking the environment, pulling workbook sources, pushing sources, and running the configured macro.
- Added command palette actions for creating projects, initializing projects from existing workbooks, and starting, checking, or stopping workbook sessions.
- Added a command palette action for saving the managed workbook with `xlflow save`.
- Added a command palette action for running workbook tests with `xlflow test` and showing results in the output channel.
- Added VS Code Test Explorer integration backed by `xlflow test list --json` and `xlflow test --json`.
- Added command palette actions for linting the workspace and formatting the active document or full project.
- Added a Status Bar session indicator with QuickPick actions for starting, stopping, restarting, inspecting, and diagnosing xlflow sessions.
- Tightened v0.1 review feedback by wiring LSP trace settings, gating startup test discovery, documenting Node 22 development requirements, and hardening xlflow process handling.
