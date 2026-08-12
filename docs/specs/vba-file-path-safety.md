# VBA file and path safety

<!-- xlflow-rule-contract: {"id":"VBA245","family":"analyze","category":"security","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_unsafe_file_path","inline_suppressible":true,"preflight_blocking":false} -->

`VBA245` reports destructive or state-dependent VBA file operations whose path
cannot be shown safe by the procedure-local analysis. It is a default-enabled,
warning-level, non-blocking rule available to batch analysis and realtime/LSP
diagnostics. Disable it with `[analyze].disabled_rules = ["VBA245"]` or use an
inline suppression on an intentional operation.

The rule recognizes `Kill`, `RmDir`, `Name ... As`, `FileCopy`, file opens for
`Output`/`Append` and binary opens that reach `Put`/`Write`, FileSystemObject
delete/copy/move/create methods, and workbook `SaveAs`/`SaveCopyAs`. It tracks simple assignments, aliases, concatenation,
`BuildPath`, `ThisWorkbook.Path`, configured project anchors, and the explicit
external-input sources already used by the VBA source-to-sink analysis.

Findings distinguish definite path hazards (`empty_path`, `root_path`,
`wildcard_delete`, `relative_path`, `directory_traversal`,
`current_directory_dependency`, `same_source_destination`, and
`unchecked_overwrite`) from input-dependent and lifecycle risks. Existence
checks such as `Dir`, `FileExists`, `FolderExists`, and `GetAttr` are not sinks.
Input-dependent kinds include `untrusted_filename` and `unknown_path`; the
lifecycle kind is `temporary_cleanup_missing`.
The analyzer is lexical and symbolic: it does not resolve runtime symlinks,
junctions, or the actual current directory.

The additive `file_operation` context contains the operation, path role, risk
class, risk kind, origin state, trusted anchor when known, and an explicit
overwrite fact when the API provides one. Suggestions require a trusted anchor,
validated leaf name, rejected separators and traversal segments, explicit
overwrite policy, and cleanup on every local exit for temporary files.

When `VBA245` is disabled, the existing `VBA224` source-to-sink fallback remains
available for its legacy destructive-file and SaveAs sinks; non-destructive
workbook-open flows remain under `VBA224` regardless of this setting.
