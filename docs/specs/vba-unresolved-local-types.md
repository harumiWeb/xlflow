# VBA229 unresolved local As type names

`VBA229` is a default-enabled, compile-equivalent analyzer diagnostic for
procedure-local `Dim` and `Static` declarations whose `As <Type>` identifier
cannot be resolved.

The rule runs in batch analysis and the LSP. It reports an error on the type
identifier itself, is unsuppressible, and blocks source preflight before Excel
opens. The diagnostic is emitted only when resolution completed successfully;
an unavailable workspace or type database is not treated as proof that a name
is missing.

Type-database load integrity and type-name absence are separate facts. A
missing, malformed, empty, or partially materialized generated TypeLib
manifest leaves reference resolution incomplete. The embedded database remains
available for positive resolution in that state, but a lookup miss does not
produce `VBA229`. This prevents the availability of machine-local generated
TypeLib files from changing an unsupported absence claim into a blocking error.

Resolution uses the same production configuration as the analyzer and LSP:
VBA intrinsic types, built-in host metadata, generated/reference TypeLib
metadata, embedded enum groups, project `Type`/`Enum` declarations, class
modules, UserForms, and document modules are accepted. Module-qualified names
use the workspace index. A class qualified by the configured project name is
resolved against the project's object modules; TypeLib-qualified names use the
database lookup.

Version 1 deliberately covers procedure-local declarations only. Parameters,
function/property return types, module-level declarations, reference
management, and new oracle fixture schema surfaces remain separate work.
