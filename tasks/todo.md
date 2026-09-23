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
