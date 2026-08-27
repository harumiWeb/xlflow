# VBA Resource Lifecycle Analysis

`VBA219` detects procedure-local Workbook and VBA file-handle acquisitions
that can reach an exit without a recognized release. ADR-0024 records the
design rationale.

## Acquisitions and ownership

The rule recognizes only these v1 acquisition forms:

- `Set <local> = Workbooks.Open(...)` or
  `Set <local> = Application.Workbooks.Open(...)` for a procedure-local
  variable declared as `Workbook` (including a qualified type such as
  `Excel.Workbook`); and
- `Open ... For Input|Output|Append|Binary|Random As #<local-or-literal>`.

The acquisition takes effect only along the CFG statement's normal successor.
An uncaptured `Workbooks.Open` remains the responsibility of `VBA205`.
Parameters are borrowed references and do not create an obligation. A local
Workbook is transferred only when `Set <FunctionName> = <local>` remains intact
through an object-returning Function's normal exit. An exceptional, unknown, or
termination exit after that assignment still requires a Close, as does a later
overwrite of the return slot. A direct
`Set <FunctionName> = Workbooks.Open(...)` assignment to a `Workbook` return
slot establishes the same transfer when it reaches a normal exit. Non-Workbook
locals and return slots, ByRef parameters, module variables, helper calls, and
external APIs do not establish acquisition, transfer, or release in v1.

## Releases and paths

Workbook aliases formed by direct `Set alias = owner` assignments and file
handle aliases formed by direct scalar assignments are recognized locally.
File-handle aliases established before the matching `Open ... As #handle` are
preserved when they are definitely equal on every CFG path to that Open.
`alias.Close`, `Close #alias`, and bare `Close` (for all tracked file handles)
release the corresponding acquisition.

The analyzer traverses normal, exceptional, termination, and unknown CFG exits.
Cleanup labels and error handlers require no name-based exception: a release is
safe whenever the CFG proves it is reached. A reached release statement is the
cleanup boundary even if the release call itself could fail; modeling failures
inside Close is outside this rule. When a matching Close is recognized on some
path but not proven on every exit, the diagnostic explains both facts. It must
not imply that no cleanup source exists, and it must not attribute the remaining
uncertainty specifically to `On Error Resume Next`.

Unreachable acquisitions and recovered statements are never used as ownership,
release, transfer, or alias evidence. Alias assignments take effect only on
their normal CFG edges, so exception handlers retain the pre-assignment alias
state.

Each finding is located at the acquisition and its reason identifies the
uncovered exit witness and, when available, the recognized cleanup statement.
The rule is default-enabled, warning-only,
non-blocking, realtime, and inline suppressible. Disable it project-wide with
`[analyze].disabled_rules = ["VBA219"]`; the legacy
`detect_resource_leaks` boolean remains accepted with a deprecation warning.

## Boundaries

The rule does not yet model FileSystemObject/TextStream, ADODB, Office
automation applications, temporary-file removal, dynamic aliases, or
interprocedural cleanup. Those resources must add their own explicit
acquisition/release pairs before participating in this contract.
