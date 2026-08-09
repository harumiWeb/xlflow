# VBA222 public API type safety

`VBA222` is a default-enabled, warning-level, non-blocking, batch-only
`analyze` rule. It checks the resolved types in public `Function` and
`Property Get` returns, all parameters of public procedures and property
accessors, and parameters of public custom `Event` declarations. It is not
part of LSP real-time diagnostics.

## Public API surfaces

Standard-module procedures and events are public surfaces unless their
declaration is `Private` or `Friend`. A class or interface module is a public
surface only when its exported module attributes contain `VB_Exposed = True`.
Missing or false `VB_Exposed` therefore makes the class implementation
internal, including its otherwise public members. Host-required procedures
recognized by `ProcedureSymbol.IsEventHandler` (`Worksheet_*`, `Workbook_*`,
UserForm/control events, and `Auto_Open`/`Auto_Close`) are excluded because
their signatures are dictated by the host contract.

Property getter/setter pairs are also checked for effective visibility
mismatch. The mismatch is reported under `VBA222` because it makes the public
property contract unstable even when each individual type is known.

## Type resolution

The project index includes classes, `Private`/`Public Type` declarations, and
Enums. A type in the declaring module wins over another project candidate;
qualified names are matched to their module. A unique inaccessible project
type is reported as inaccessible, while multiple remaining candidates are
reported as ambiguous. Intrinsic VBA types, `Object`, `Variant`, and types
resolved from the embedded built-in or generated TypeLib database are allowed.

The embedded baseline includes the VBA `VbAppWinStyle` enum and the standard
OLE Automation interfaces `IUnknown` and `IEnumVARIANT`, including their
`stdole`-qualified spellings. A qualified external type may use a short-name
database entry only when that entry belongs to the same library; an unrelated
type such as `ADODB.DataTypeEnum` must not satisfy `DAO.DataTypeEnum`.

An external type absent from both the project index and the available TypeLib
database is reported conservatively as unresolved. This rule does not attempt
to prove the complete VBE reference configuration or whether an optional
library is installed. Every finding includes the problematic type in its
message, reason, and suggestion and points at the parameter or declaration
line when the parser provides that range.

The implementation uses `procedureir` only for syntax facts and module
attributes; the project index and diagnostic ownership remain in `analyze`.
This keeps the IR reusable by batch and LSP consumers without adding the
batch-only rule to the real-time path.
