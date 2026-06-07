# Dialog Watcher

The .NET Excel bridge owns reusable Excel/VBE dialog detection and safe
suppression behavior.

## Detection

- Enumerate top-level and VBE thread windows through Win32.
- Correlate candidates by Excel PID, VBE thread, HWND, owner chain, title,
  class name, process image, and detection time.
- Do not require `IsWindowVisible`. Excel can create or defer painting a modal
  runtime error dialog until the Excel window receives focus.
- Use UI Automation metadata and InvokePattern when available, but do not assume
  a UIA element and HWND have identical identity.

Supported fingerprints are `runtime`, `compile`, `msgbox`, `inputbox`, and
`filedialog`.

## Snapshot Schema

JSON diagnostics may include:

- `kind`, `detected_at_ms`, `sources`
- `hwnd`, `pid`, `thread_id`, `owner_hwnd`, `root_owner_hwnd`
- `title`, `class_name`, `visible`, `process_image`
- optional `automation_id`, `name`, `control_type`
- `text`, `buttons`, `children`
- `action`, `action_method`, `action_target`, `action_succeeded`

Button and child snapshots may include language-neutral Win32 metadata:
`access_key`, `control_id`, and `enabled`. `access_key` is extracted from
native accelerator markers such as `&D` or localized suffixes such as `(D)`.

## VBE Selection Diagnostics

For `.NET` `run` compile/runtime dialog suppression and `push` compile dialog
suppression, the bridge attempts to capture VBE source selection before
dismissing the dialog and retries once after dismissal if no meaningful location
was found. This capture is best-effort and timeout-bounded; failure to read VBE
state must not prevent dialog suppression or change the command failure
classification.

When available, `run_diagnostic.location` or `push_diagnostic.location` may
include `confidence`, `method`, `source_path`, `component`, `component_type`,
`procedure`, `line`, `column`, `end_line`, `end_column`, and selected line
`text`. Verified line values are source-file line numbers, adjusted for VBE
hidden export metadata such as `Attribute VB_*`; unreliable column values are
omitted. Capture failure metadata is reported under `*.location_capture.attempts`
with timing labels such as `before_dialog_action` and `after_dialog_action`.

## Action Policy

- Runtime error: prefer accelerator-key selection for Debug (`D`) or End (`E`)
  according to the active policy, then localized text fallback, then explicit
  OK/Close. End is preferred for unattended suppression because Debug can leave
  VBE in break mode.
- Compile error: prefer standard OK/Close control IDs and localized OK/Close
  text. For VBE compile dialogs only, a single enabled non-Help primary button
  may be used as a language-neutral fallback. Do not use an arbitrary
  first-button or window-close fallback.
- Native MsgBox/InputBox/FileDialog: cancel only when an explicit Cancel/Close
  action is identifiable.
- Do not assign scripted values to arbitrary native dialogs. Deterministic
  values remain limited to stable `XlflowUI` marker IDs.

Action order is UIA InvokePattern, `BM_CLICK`, then an explicitly supported
window-close fallback.
