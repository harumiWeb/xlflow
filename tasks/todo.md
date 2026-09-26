# Issue #826: VBA class/interface public-API hazard diagnostics (VBA278-VBA282)

Parent issue #816. Five opt-in warning rules, all default-disabled,
warning-level, inline-suppressible, non-blocking, on batch + realtime + LSP.
Renumbered from the originally implemented VBA273-VBA277 because origin/main
PR #845 (issue #823) already owns those IDs for parameter-passing rules.

- VBA278 public member underscore names (class/form/document; skips
  control-verified/intrinsic event handlers, verified Implements bindings,
  unresolved-Implements prefixes, and Class_Initialize/Class_Terminate).
- VBA279 non-Private Enum in document modules.
- VBA280 write-only Property (Let/Set without Get), one finding per name;
  Friend counts as public-facing; conditional/recovered accessors mark the
  name uncertain and fail open.
- VBA281 public members whose <Interface>\_<Member> binding resolves to a
  declared public member of the implemented interface class, and explicitly
  Public recognized event handlers.
- VBA282 self-name access inside predeclared modules (document/form/
  VB_PredeclaredId class); local/parameter/module-scope shadows and TypeOf
  type operands excluded; recovered/conditional procedures skipped.

PR #847 review fixes applied (Devin 1-6 + CodeRabbit 1-4; all verified valid):

- Form helper misclassified as event: recognizedEventMember validates the
  control prefix against designer controls when known and falls back to
  intrinsic UserForm\_\* names when unknown. Lazily resolves controls because
  VBA220's gate only covers \_change/\_click names.
- IFace_Utility FP: interfaceBindingMatch requires the suffix to name a
  public member of the resolved, unambiguous interface class; resolved but
  absent suffixes fall back to VBA278, unresolved interfaces fail open.
- VBA280 uncertain accessors: Recovered/ConditionalBranches mark the
  property name uncertain; Friend writers now count as public-facing.
- VBA282 guards: proc.IR.Symbol.Recovered/ConditionalBranches fail open;
  ScopeLocal/ScopeParameter/ScopeModule accesses excluded; TypeOf exclusion
  narrowed to the last (type-operand) child of type_of_expression.
- Spec fail-open text updated to match implementation.

# Conflict resolution (merge origin/main @ aabb6575)

- Rule ID collision: our VBA273-277 -> VBA278-282; config keys unchanged.
- registry.json: main entries kept, ours appended renumbered.
- diagnostics.md regenerated from registry; oracle fixture dirs + ids and
  fixture_contract_test.go renumbered; workspace_test.go keys renumbered.
- Corpus snapshots to be regenerated (corpus:update-snapshots) and ledger
  rows renumbered to VBA278-282; metrics numbers rechecked after regen.

Remaining: corpus regen + metrics, format:check/docs:check, full test pass,
commit merge, push, update PR #847 body (new rule IDs + oracle case IDs
vba278-public-underscore-member, vba279-document-public-enum,
vba280-write-only-property, vba281-public-event-handler,
vba281-public-interface-member-collision, vba282-predeclared-self-access).

# PR #845 review follow-up (Issue #823, VBA273-VBA277)

Review comments on internal/analyze/parameter_passing.go. Verified against
code; both bug reports are valid and in scope. Plan:

1. [Devin + CodeRabbit] Property accessor collision: mutation records are
   keyed by qualified name only, so Property Get Item and Property Let Item
   share one summary (first registered wins). VBA275 can falsely claim the
   Let index parameter is never written; VBA274 can miss a written Let value
   parameter. Fix: kind-qualify record keys (kind|name) for summary storage
   and finding lookup; call-edge resolution uses candidate kind when present
   and falls back to unique-name match or possiblyWritten when ambiguous.
   Regression test: Property Get before Property Let, Let writes index and
   assigns value parameter.
2. [Devin] applyArgumentMutation checks parameterTypeIsObject before
   parameterTypeIsUDT, so a member argument on a UDT that shadows a builtin
   object name (Type Collection) drops the propagated write and VBA275
   falsely fires. Same latent issue family as the recordParameterMemberWrite
   fix. Fix: UDT check first, matching VBA name-resolution precedence.
   Regression test: module UDT named Collection, member forwarded to a
   writing ByRef call.

Advisory (Devin): no executed VBE oracle case IDs. Oracle v1 supports
compile probes only; the modeled semantics (property value ByVal, forced
ByVal parens) are covered by byref-parenthesized-variable (accepted) and
property-signature-valid/invalid at compile level; runtime cases are out of
scope for the current probe contract. Rules remain opt-in and fail open.

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

# Issue #823 final-review pass 1 follow-up

Review run_7c453f1dc6a5 on commit 202af7ff. All four confirmed findings
verified against IR probes and real Excel runtime checks; all valid. Plan:

1. F1: With-block implicit member writes invisible to the mutation summary.
   Fix: build a shared withReceiver map (statement ID to innermost With
   receiver expression), resolve implicit member targets and implicit member
   call arguments through it, and fail open on the unresolved spaced-callee
   form (Replace .Left parses into the callee expression).
2. F2: Erase/Input #/Line Input #/Get # operands surface as unknown
   statements with read accesses. Fix: record write targets for those
   syntax kinds (file_number_literal excluded; Get uses TargetID only).
3. F3: VBA273 fires on ParamArray with an impossible suggestion. Fix:
   exclude ParamArray from VBA273 and VBA277.
4. F4: cfg/clone.go cloneParameters missed PassingRange. Fix: clone it and
   extend the snapshot-isolation test.
5. Follow-ups: VBA276 gains the constrained exclusion (VBE rejects
   Implements passing-modifier mismatches in both directions; verified on
   real Excel). Indexed element writes (v(0) = 1) become possiblyWritten:
   not a VBA274 reassignment, still caller-visible so VBA275 stays
   suppressed. Call Mutate((x)) is forced ByVal and no longer propagates.
   Missing mutation summary now returns nil explicitly (fail-open).

# Issue #823 final-review pass 2 follow-up

Review run_7c453f1dc6a5 pass 2 on commit 378b1333. Three confirmed findings
verified against the report's IR dumps, probe outputs, and the implicated
source; all valid. Plan:

1. F-P2-1: named-argument ExpressionIDs bind the widest expression in the
   whole `a:=x` range, so a name at least as long as its value hides the
   value from propagation (VBA275 false positive) and from every consumer
   that assumes ExpressionIDs[i] == Named[].ExpressionID
   (object_use_before_set, array_safety_source_order, object_container_flow,
   variable_assignment). Fix in procedureir/visitor.go: use
   argument.valueRange when the argument is named; add a width-inverted
   regression test for parameter passing and named-argument coverage for the
   sibling consumers.
2. F-P2-2: indexed call expressions (arr(0)) emit no receiver access, so
   element-valued arguments, With receivers, and file-statement operands are
   invisible to propagation (VBA275 false positives). Fix: resolve
   ExpressionCall roots through the callee child in parameterExprRootName,
   generalize the implicit-member argument fallback to a root-resolution
   fallback for any unhandled argument, and mark ExpressionCall write
   targets possiblyWritten in recordParameterWriteTarget.
3. F-P2-3: parameterMemberTargetRoot stops at spaces so [My Field] loses
   member writes, and parameterTypeIsObject runs before the UDT check so a
   UDT shadowing a builtin object name (e.g. Collection) drops member
   writes. Fix: extract bracketed names whole, fall back to expression
   resolution whenever the text root matches no parameter, and check UDT
   membership before the builtin object list.
4. Regression tests for each, plus docs/spec updates (named mapping is now
   width-independent; element-valued arguments and receivers covered).
