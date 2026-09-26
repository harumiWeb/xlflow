# VBA Class/Interface Public-API Hazard Diagnostics

<!-- xlflow-rule-contract: {"id":"VBA273","family":"analyze","category":"maintainability","default_severity":"warning","scope":"file-local","realtime":true,"configuration_key":"detect_public_member_underscore_names","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VBA274","family":"analyze","category":"correctness","default_severity":"warning","scope":"file-local","realtime":true,"configuration_key":"detect_document_module_public_enum","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VBA275","family":"analyze","category":"maintainability","default_severity":"warning","scope":"file-local","realtime":true,"configuration_key":"detect_write_only_property","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VBA276","family":"analyze","category":"correctness","default_severity":"warning","scope":"file-local","realtime":true,"configuration_key":"detect_public_interface_event_members","inline_suppressible":true,"preflight_blocking":false} -->
<!-- xlflow-rule-contract: {"id":"VBA277","family":"analyze","category":"runtime-safety","default_severity":"warning","scope":"procedure-local","realtime":true,"configuration_key":"detect_predeclared_instance_access","inline_suppressible":true,"preflight_blocking":false} -->

Issue #826 adds five opt-in diagnostics for VBA class and interface
hazards. All rules consume module-kind and declaration metadata from the
procedure IR rather than text matching, are disabled by default, and run
on both the batch analyze and realtime/LSP surfaces with parity. Enable
each rule through the `detect_*` keys listed in the contracts above, or
suppress a site with `xlflow:disable-line` / `xlflow:disable-next-line`.

## Rule contracts

### VBA273 - Public member name contains underscore

Public Sub/Function/Property members in class, form, or document modules
whose names contain `_` collide with two reserved VBA naming contracts:
`<Interface>_<Member>` implementation bindings and `<Object>_<Event>`
handler names. The rule reports the member's declaration. Recognized
event handlers (`IsEventHandler`) and members whose name already matches
an `Implements`-declared interface prefix (covered by VBA276) are skipped
so each member produces one finding. The mandated class lifecycle names
`Class_Initialize` and `Class_Terminate` are also skipped in class
modules: VBA reserves those exact identifiers, so they cannot be renamed
and cannot collide with a user-chosen interface or control binding.
Private and Friend members are not part of the public default interface
and are not flagged.

### VBA274 - Public Enum in document module

A non-Private `Enum` declaration in a worksheet or workbook code-behind
module is duplicated into every runtime copy of that document, producing
an ambiguous-name compile error. Only module-kind `document` is
affected; standard and class modules are not flagged. The rule reports
the Enum declaration itself.

### VBA275 - Write-only property

A `Property Let` or `Property Set` member with no `Property Get` on
the same name is a write-only API. The rule flags each property name once,
anchored on the first public-facing writer (implicit Public, explicit
Public, or Friend). Private writers are module-internal conveniences and
stay silent. A `Get` on the same name satisfies the contract regardless
of accessor count.

### VBA276 - Interface implementation or event handler exposed as Public

Two public-interface exposure hazards share this rule:

- A member whose name matches an `Implements` interface's
  `<Interface>_<Member>` binding. Interface member names are matched
  against the complete declared interface name, because interface names
  may themselves contain underscores (`Implements I_Foo` binds
  `I_Foo_Bar`, not `I_*`). Both implicit-public and explicit
  `Public` members are flagged.
- A recognized event handler declared explicitly `Public` (implicit
  visibility is the VBA default and is not flagged). Event procedures are
  raised by the host; declaring them `Public` exposes them on the
  module's default interface.

Both shapes are reported on the member declaration. Private `I_M`
implementation members - the required VBA form - are not flagged.

### VBA277 - Predeclared-instance self-name access

Inside a module that owns a predeclared default instance - document
modules (worksheet/workbook classes), UserForms, and classes carrying
`Attribute VB_PredeclaredId = True` - a reference to the containing
module's own name binds to the shared default instance rather than to
`Me`. The rule reports each such access. Type-position identifiers
(`As` clauses, `Implements` targets, `New <Name>` and
`TypeOf ... Is <Name>` operands) are declaration or type operands and
never produce an access, so they are not flagged.

## Fail-open boundaries

All five rules skip symbols marked recovered and procedures declared under
conditional-compilation branches, because their IR projection may be
incomplete. A file whose `ModuleKind` is missing or unrecognized
produces no VBA273/VBA276/VBA277 findings (object-module kinds only);
VBA274 requires `document` and VBA275 runs on any module kind with
properties.

## Precision and performance

VBA273, VBA275, and VBA276 iterate the module's already-materialized
procedure symbols; VBA274 iterates module declarations. VBA277 iterates
the procedure's access list and walks the expression-parent chain only for
same-name hits. None of the rules add planner requirements, fixed-point
work, or interprocedural evaluation. Diagnostic ordering follows the
analyzer's deterministic source-order sort.

## Boundaries

VBA273, VBA275, and VBA276 derive from declaration metadata and do not
depend on VBE compile behavior. VBA274's ambiguous-name-on-sheet-copy
hazard and VBA276's interface-binding collision are documented VBA
semantics; focused VBE oracle cases bind the compile-error shape where
the local oracle confirms it. The rules cover the inspection intent of
Rubberduck's underscore member-name, document-module Enum, write-only
property, member-not-on-interface, and predeclared-instance inspections,
implemented against xlflow's module-kind and declaration model rather
than Rubberduck's declaration resolver.
