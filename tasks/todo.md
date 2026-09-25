# PR #845 review follow-up (Issue #823, VBA270-VBA274)

Review comments on internal/analyze/parameter_passing.go. Verified against
code; both bug reports are valid and in scope. Plan:

1. [Devin + CodeRabbit] Property accessor collision: mutation records are
   keyed by qualified name only, so Property Get Item and Property Let Item
   share one summary (first registered wins). VBA272 can falsely claim the
   Let index parameter is never written; VBA271 can miss a written Let value
   parameter. Fix: kind-qualify record keys (kind|name) for summary storage
   and finding lookup; call-edge resolution uses candidate kind when present
   and falls back to unique-name match or possiblyWritten when ambiguous.
   Regression test: Property Get before Property Let, Let writes index and
   assigns value parameter.
2. [Devin] applyArgumentMutation checks parameterTypeIsObject before
   parameterTypeIsUDT, so a member argument on a UDT that shadows a builtin
   object name (Type Collection) drops the propagated write and VBA272
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
3. F3: VBA270 fires on ParamArray with an impossible suggestion. Fix:
   exclude ParamArray from VBA270 and VBA274.
4. F4: cfg/clone.go cloneParameters missed PassingRange. Fix: clone it and
   extend the snapshot-isolation test.
5. Follow-ups: VBA273 gains the constrained exclusion (VBE rejects
   Implements passing-modifier mismatches in both directions; verified on
   real Excel). Indexed element writes (v(0) = 1) become possiblyWritten:
   not a VBA271 reassignment, still caller-visible so VBA272 stays
   suppressed. Call Mutate((x)) is forced ByVal and no longer propagates.
   Missing mutation summary now returns nil explicitly (fail-open).

# Issue #823 final-review pass 2 follow-up

Review run_7c453f1dc6a5 pass 2 on commit 378b1333. Three confirmed findings
verified against the report's IR dumps, probe outputs, and the implicated
source; all valid. Plan:

1. F-P2-1: named-argument ExpressionIDs bind the widest expression in the
   whole `a:=x` range, so a name at least as long as its value hides the
   value from propagation (VBA272 false positive) and from every consumer
   that assumes ExpressionIDs[i] == Named[].ExpressionID
   (object_use_before_set, array_safety_source_order, object_container_flow,
   variable_assignment). Fix in procedureir/visitor.go: use
   argument.valueRange when the argument is named; add a width-inverted
   regression test for parameter passing and named-argument coverage for the
   sibling consumers.
2. F-P2-2: indexed call expressions (arr(0)) emit no receiver access, so
   element-valued arguments, With receivers, and file-statement operands are
   invisible to propagation (VBA272 false positives). Fix: resolve
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
