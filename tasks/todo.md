# Issue #826: VBA class/interface public-API hazard diagnostics (VBA273-VBA277)

Parent issue #816. Implemented five opt-in warning rules, all default-disabled,
warning-level, inline-suppressible, non-blocking, on batch + realtime + LSP:

- VBA273 public member underscore names (class/form/document; skips
  event handlers, Implements bindings, Class_Initialize/Class_Terminate).
- VBA274 non-Private Enum in document modules.
- VBA275 write-only Property (Let/Set without Get), one finding per name.
- VBA276 public members matching `<Interface>_<Member>` binding (full interface
  name prefix match) and explicitly Public event handlers.
- VBA277 self-name access inside predeclared modules (document/form/
  VB_PredeclaredId class); type operands excluded via expression parents.

Artifacts: internal/analyze/class_interface_hazards.go + tests, config keys
detect_public_member_underscore_names / detect_document_module_public_enum /
detect_write_only_property / detect_public_interface_event_members /
detect_predeclared_instance_access, registry.json entries,
docs/specs/vba-class-interface-diagnostics.md, cli-contract/vitepress docs,
CHANGELOG Unreleased entry, corpus materialization flags + snapshots +
273 true-positive ledger rows, six VBE oracle fixtures (promoted
2026-09-25T18:12:27Z, asserted expectations).

Verified: go test ./internal/analyze ./internal/config
./internal/staticanalysis/rules ./internal/oracle; task corpus:test
(snapshots match, unreviewed=0); task corpus:metrics (reviewed=10531
tp=7648 fp=2883 allowed=90); pnpm docs:check / format:check; gofmt.

Remaining: commit/push/PR on user request; record oracle case IDs
(vba273-public-underscore-member, vba274-document-public-enum,
vba275-write-only-property, vba276-public-event-handler,
vba276-public-interface-member-collision, vba277-predeclared-self-access)
in the PR body.

# PR #837 review follow-up (Issue #822, VBA257/VBA258)

Review comments on internal/analyze/discarded_return.go. Verified against IR
dump and resolver source; all six are valid. Plan:

1. [Devin] Parse errors permit false all-discard claims: VBA258 claims every
   caller discards, but a parse error can hide a call or a dynamic-dispatch
   reference (Application.Run reaches Private procedures too). Fix: bail out
   of the whole pass when any file.IR.Parse has HasError/HasMissing.
2. [CodeRabbit] PathFilter removes modules before parsedFiles; hidden modules
   can contain callers or dynamic references. Note: resolutionComplete is too
   broad (AnalyzeProject always sets typeDBResolutionIncomplete), so gate on
   PathFilter==nil instead. Combined with #1 this also covers the Friend case.
3. [Devin] Same-name arguments hidden by isCalleePart: 'Foo Foo' argument is
   skipped because it shares the callee name inside the call range. Fix:
   locate the exact callee token range (bytes.Index of Callee.Text inside
   call.Range) and exempt only expressions inside that range.
4. [CodeRabbit] Implicit Application dispatch: receiverless 'Run "x"' and
   With-block '.Run' produce no dynamic references. Fix: in calls package,
   map callee member run/ontime/onkey with nil/empty/"." receiver to
   "application.<member>" when resolution ran and did not bind a project
   procedure (keeps inspect path unchanged via ResolutionNotAttempted).
5. [Devin] Underscore exclusion too broad: Implements modules exclude every
   underscored proc. Fix: collect implemented interface names from
   TypeReferences (kind=="implements") and exclude only <Interface>\_<Member>.
6. [CodeRabbit nitpick] Implements exclusion test cannot fail (unknownDynamic
   suppresses everything and IWorker_DoWork is never called). Fix: add an
   independent test with a standalone IWorker_DoWork call and no dynamic
   dispatch.

Non-actionable: CodeRabbit docstring-coverage warning (bot threshold, repo
convention does not doc-comment every function).

# PR #837 review follow-up 2 (post-merge)

Second-round comments verified against merged HEAD (bfa7e23a). Devin's three
comments and CodeRabbit's two actionable comments are stale: they target the
pre-fix head and are already covered by 92cfacb3 (parse-error bail-out,
calleeTokenRange, implementedInterfaceNames, projectViewComplete gate,
implicitApplicationAPIName) - CodeRabbit marks both as addressed.

One valid finding remains:

1. [CodeRabbit] Interface names containing underscores: `Implements I_Foo`
   makes the implementation `I_Foo_Bar`, but strings.Cut splits at the first
   underscore so qualifier `i` never matches `i_foo`. Fix: prefix-match the
   lowered symbol name against each complete interface name + `_`. Add a
   regression test with an underscored interface name.

# PR #838 review follow-up (Issue #827, VBA260/VBA261/VBA262)

Verified against commit `2ebcf3f6`; actionable plan:

1. [Devin] Exclude bracket-escaped identifiers that bind to procedure, module,
   or project declarations from VBA262. Add declared enum/local/module and
   undeclared host-expression regressions.
2. [Devin] Require VBA260's `ThisWorkbook` token to be an unqualified,
   unshadowed project-global receiver. Reject member-qualified and locally
   shadowed forms with focused tests.
3. [Devin/CodeRabbit] Preserve logical-to-physical source positions for VBA260
   continuation statements so ranges and line suppressions bind to the actual
   physical line. Cover indentation and a match beginning on a continued line.
4. [CodeRabbit] Distinguish partial worksheet catalogs from unavailable
   catalogs in structured warning text while retaining valid mappings.
5. [CodeRabbit] Skip `sheetData` subtrees through `xml.Decoder.Skip` to reduce
   token handling while still validating XML through EOF and rejecting later
   conflicting `sheetPr` metadata.
6. [CodeRabbit acceptance check] Extend VBA261 diagnostics and tests to explain
   the error-behavior difference as well as binding and return-type semantics.
7. Run focused tests, full affected packages, corpus snapshots, lint/docs, and
   push the review-fix commit. Re-read PR comments and remote checks.

Non-actionable/follow-up: repository style does not require doc comments on all
private helpers. The security job reports GO-2026-6452 in the pre-existing
excelize v2.11.0 dependency with no fixed version, so it is tracked separately
from this review fix. VBE-oracle need will be reassessed after checking existing
bracket fixtures; VBA262 is not compile-equivalent and will not be promoted as
compile evidence.
