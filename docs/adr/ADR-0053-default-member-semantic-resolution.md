# ADR-0053: Conservative Semantic Resolution for VBA Default Members and Bang Notation

## Status

Accepted

## Context

VBA silently invokes default members in several contexts. An assignment such as
`value = rangeValue`, an indexed expression such as `items(key)`, and a
procedure call can all resolve through a property that is not written in the
source. Default members can themselves return values with another default
member, and `object!name` adds a stringly typed access path that is easy to
mistake for an ordinary member reference.

Issue #817 asks xlflow to identify known, indexed, recursive, and unbound
default-member access, distinguish it from explicit member syntax, and expose
deterministic failures without treating late-bound VBA as statically known.
The existing revision-scoped document resolver provides the project-symbol and
expression-type view, while `procedureir` preserves dot versus bang member
operators and `internal/vbadb` carries default-member and member-signature
metadata for generated and curated TypeLib data. Batch analysis, realtime
analysis, and the LSP workspace must use the same semantic result; otherwise a
source edit can receive different findings depending on the caller.

This feature also crosses the existing diagnostic ownership boundary. A known
implicit access is a high-confidence correctness/reliability signal, an
unresolved access is only an advisory risk, and a complete type model can prove
that an access cannot satisfy its value context. Treating all of
these as one warning would either hide deterministic failures or turn
`Object`, `Variant`, and incomplete TypeLib metadata into false positives.

## Decision

Add one shared, revision-scoped default-member semantic resolver. It consumes
logical source statements, the document expression-type resolver, project
declarations, retained member-operator facts, and the available TypeDB view.
It classifies the access shape and expected use context once; batch, realtime,
and LSP projections consume that result rather than implementing independent
classifiers.

The initial resolver recognizes default-member metadata from the TypeDB member
set and exported project declarations where an exact
`Attribute <member>.VB_UserMemId = 0` names a value-producing Function or
Property Get in the same class-like module. A comment, naming convention, or
missing metadata is not sufficient to prove a default member or its absence.
Dot and explicit named member calls remain explicit syntax and are not reported
as implicit access. Bang notation is retained as a separate syntax
classification.

The public rule contracts are:

| Rule     | Contract                                                                                                                                                               | Default                                                                      |
| -------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------- |
| `VBA253` | A known implicit or indexed default-member invocation, including a known recursive chain, is reported as a high-precision correctness warning.                         | Disabled initially; corpus evaluation found 406 existing idioms.             |
| `VBA254` | An implicit/default-member access whose binding is unbound, late-bound, or otherwise incomplete is reported as a lower-confidence advisory.                            | Disabled; opt-in                                                             |
| `VBA255` | Bang notation is reported as a stringly typed or recursively/unbound default-member access concern.                                                                    | Disabled; opt-in                                                             |
| `VBA249` | With default-member analysis enabled, a complete semantic model proves that the required value-producing default member is absent or a default-member chain is cyclic. | Existing runtime-error rule ID and severity; additive opt-in analysis branch |
| `VBA202` | A receiver proven uninitialized or `Nothing` remains the owner of deterministic error 91 findings.                                                                     | Existing contract                                                            |

`VBA253`, `VBA254`, and `VBA255` are non-blocking, inline-suppressible,
procedure-local analyzer rules available to batch and realtime/LSP surfaces.
Their canonical compatibility keys are respectively
`detect_implicit_default_member_access`,
`detect_unbound_default_member_access`, and `detect_bang_notation`. The shared
`[analyze].disabled_rules` policy remains authoritative when both a legacy
boolean and a rule-level disable are present. All three rules initially remain
opt-in: the corpus run found 406 VBA253 candidates, so explicit project policy
is required until that signal is reviewed.

The finding envelope adds a `default_member` context with the access kind,
known/unbound binding, the initial expected context (`value`), resolved member
when known, and default-chain depth. A
default-member-owned `VBA249` finding may use `binding: "invalid"` to mark a
complete negative proof. This additive context is stable across batch JSON and
internal realtime projections; LSP preserves the diagnostic code, message,
severity, and range. `VBA249` continues to use its existing `runtime_error`
context; its new kinds identify only deterministically invalid default-member
resolution. A possible or unknown failure is never promoted to `VBA249`.

Resolution is deliberately fail-open. Explicit generated metadata may prove a
member exists even when the wider database is incomplete, but missing, stale,
malformed, partial, or curated-only TypeDB data cannot prove absence. Recovered
source, ambiguous project shadowing, unknown `Object`/`Variant` values, late
binding, dynamic dispatch, and unresolved external calls likewise do not
establish negative evidence. A complete metadata view may establish absence.
A receiver that may
be `Nothing` or uninitialized invalidates a default-member runtime proof, and
`VBA202` owns the error-91 finding when the receiver is proven invalid. The
resolver must not report a duplicate default-member finding for the same
receiver and expression when another rule already owns the proven failure.

## Consequences

- Known implicit access is visible with stable source ranges and member-chain
  evidence, while explicit `.Value`, `.Item(...)`, and ordinary member calls
  remain quiet.
- Batch, realtime, and LSP diagnostics share one classification and preserve
  parity for the same source revision and workspace snapshot.
- Deterministic 438-style default-member failures can be reported without
  claiming that valid source is a VBE compile rejection. Error 91 remains under
  the established `VBA202` ownership contract.
- Unknown COM and late-bound code remains less covered, and incomplete TypeDB
  data can suppress a finding even when one Excel installation would fail. This
  is intentional and prevents negative inference from missing metadata.
- `VBA253` through `VBA255` remain opt-in after corpus evaluation found 406
  existing VBA253 candidates, many in deliberate collection/dictionary idioms.
  Their default policy must not be changed solely because a small synthetic
  fixture is clean.
- The shared resolver, registry metadata, config mapping, suppression,
  generated references, and public specification must evolve together. A
  resolver change must not silently change one consumer's result shape.

## Alternatives Considered

1. **Report every omitted member as a default-member warning.** Rejected
   because `Object`, `Variant`, late binding, and incomplete TypeDB data do not
   prove that VBA will invoke a particular member.
2. **Treat all default-member findings as compile-equivalent errors.** Rejected
   because many implicit calls are valid VBA and deterministic runtime failure
   is not the same as VBE rejection.
3. **Extend explicit member diagnostics or `VBA202` instead of adding a
   semantic resolver.** Rejected because implicit/default binding, receiver
   initialization, and explicit member availability have different facts and
   ownership boundaries.
4. **Resolve defaults independently in batch, realtime, and LSP.** Rejected
   because duplicated source scans would diverge on recursive chains,
   shadowing, and partial workspaces.
5. **Use Excel/VBE during ordinary analysis.** Rejected because source
   analysis must remain deterministic, Excel-free, parallel-safe, and usable
   on non-Windows environments. VBE oracle cases remain development evidence
   for compile-equivalent questions only.
6. **Enable unbound and bang diagnostics immediately.** Rejected because
   their medium-confidence evidence requires corpus review and can flag valid
   data-access idioms such as DAO/ADO recordset bang access.

## Evidence

- Issue #817 defines the implicit, indexed, recursive, unbound, incompatible,
  and bang-notation cases and requires corpus evaluation before enabling
  medium-confidence rules.
- `internal/vba/procedureir` preserves the immutable member-operator facts used
  by semantic consumers, while `internal/vba/intel` provides the shared
  revision-scoped document resolver used by batch, realtime, and LSP analysis.
- `internal/vbadb/db.go` defines `DefaultMember`, `DefaultMemberType`, member
  signatures, generated TypeLib provenance, and the complete-versus-curated
  merge behavior used for negative evidence.
- `docs/adr/ADR-0021-procedure-analysis-ir.md`,
  `docs/adr/ADR-0024-shared-static-analysis-rule-registry.md`, and
  `docs/adr/ADR-0048-revision-scoped-semantic-query-dag.md` establish the
  immutable IR, shared rule metadata, and revision-scoped query boundaries.
- `docs/adr/ADR-0042-deterministic-runtime-error-diagnostics.md` defines the
  distinction between runtime-error, compile-equivalent, and runtime-safety
  evidence and the ownership rules extended by `VBA249` here.
- `docs/specs/vba-runtime-error-diagnostics.md`,
  `docs/specs/vba-analysis-ir.md`, and the existing analyzer/language-server
  regression suites provide the contract and parity patterns for the new
  diagnostics.

## Related

- Issue #816 (parent)
- Issue #817
- ADR-0021, ADR-0024, ADR-0042, ADR-0048
- `docs/specs/vba-default-member-diagnostics.md`
- `docs/specs/vba-analysis-ir.md`
- `docs/specs/vba-runtime-error-diagnostics.md`
