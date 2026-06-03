# ADR-0010: Hybrid Excel Dialog Watcher and Isolated Modal COM Calls

## Status

Accepted

## Context

The .NET Excel bridge must prevent Excel and VBE modal dialogs from blocking
agent execution. Excel runtime errors, VBE compile errors, MsgBox, InputBox, and
FileDialog windows do not expose one stable identity model:

- UI Automation detects some task dialogs quickly, but VBE dialogs can live in a
  separate UIA scope.
- Win32 HWND enumeration reliably sees modal windows that do not raise useful UIA
  events.
- UIA elements and HWNDs do not always map one-to-one across dialog types.
- A synchronous `Application.Run` or VBE compile call can remain blocked while a
  modal dialog is open, preventing the bridge process from returning diagnostics.

ADR-0008 assigns Win32, UI Automation, dialog watching, and macro execution to
the .NET bridge. This ADR defines the durable implementation strategy.

## Decision

Implement a reusable hybrid dialog watcher in the .NET bridge.

- Subscribe to desktop-root UI Automation window events as a fast path.
- Poll Win32 top-level and thread windows in parallel as the reliable fallback.
- Correlate candidates using HWND, PID, thread ID, owner chain, title, class
  name, process image, detection timing, and optional UIA properties.
- Classify supported runtime, compile, MsgBox, InputBox, and FileDialog
  fingerprints and return structured dialog snapshots.
- Prefer UIA InvokePattern for reliable buttons, then use Win32 button messages.
  Use close/keyboard fallbacks only for explicitly supported fingerprints.
- Do not require `IsWindowVisible` for an owned dialog candidate. Excel can defer
  painting a runtime error dialog until it receives focus.
- For unattended runtime error suppression, prefer the explicit End action over
  Debug. Debug can leave VBE in break mode and make later COM attachment fail.
- Suppress VBA runtime and compile errors by default. For native MsgBox,
  InputBox, and FileDialog windows, cancel only when a safe Cancel/Close action
  is identifiable; deterministic value responses remain the responsibility of
  the `XlflowUI` wrapper contract.

Run modal COM calls in disposable child bridge processes that reconnect to the
target Excel instance by process ID. The parent bridge owns the dialog watcher,
timeout, child-process cleanup, and structured response. This applies to macro
invocation and VBE compile operations that can block on modal UI.

## Consequences

- Positive: Modal Excel/VBE dialogs no longer need to block the parent bridge.
- Positive: Dialog snapshots provide useful diagnostics even when COM calls do
  not return normally.
- Positive: The watcher can be reused by run, test, form, and future workbook
  automation commands.
- Positive: Native UI is not operated with ambiguous scripted values.
- Negative: The bridge gains Win32, UI Automation, correlation, and child
  process lifecycle complexity.
- Negative: Window fingerprints require Windows Excel COM E2E coverage because
  unit tests cannot reproduce every Office build and locale.
- Negative: A child worker can be terminated, but Excel itself may still need
  explicit cleanup after an unrecoverable modal or break-mode state.
- Negative: A timeout can return before Excel finishes user VBA. Cleanup that
  requires COM access cannot be performed synchronously while Excel is busy and
  the workbook must not be saved on that path.

## Alternatives Considered

1. **UI Automation only** — Rejected because VBE and modal-pump dialogs can be
   outside the expected UIA tree or event stream.
2. **Win32 polling only** — Rejected because UIA provides faster detection and
   more reliable semantic button invocation for some task dialogs.
3. **Synchronous COM calls in the parent bridge** — Rejected because a modal
   dialog can block the process that must produce diagnostics.
4. **Apply CLI response values to arbitrary native dialogs** — Rejected because
   native dialogs lack stable xlflow dialog IDs and can be miscorrelated.

## Related

- `docs/adr/ADR-0008-dotnet-excel-bridge.md`
- `docs/adr/ADR-0009-bridge-provider-contract.md`
- `docs/specs/runtime-debugging.md`
- `docs/bridge/dotnet-bridge.md`
- `docs/design.md`
