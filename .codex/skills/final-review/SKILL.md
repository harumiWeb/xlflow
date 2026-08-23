---
name: final-review
description: >
  Run the repository's final independent Codex review after completing an
  implementation and its required verification. Use this before declaring
  implementation work complete. Apply a bounded review-and-fix loop rather
  than repeatedly reviewing until no possible finding remains.
---

# Final Codex Review

Use this skill after the implementation and its required verification are
complete.

Run:

```powershell
rtk proxy task review:codex
```

The runner uses automatic scope selection. A dirty worktree is rejected when
it also contains committed changes relative to the review base, because
`--uncommitted` would otherwise omit those committed changes. In that case,
commit or stash one scope first, or explicitly choose `-ReviewMode base` or
`-ReviewMode uncommitted` when reviewing only one scope is intentional.

The command intentionally suppresses the nested reviewer's intermediate
reasoning and tool output.

While it runs, heartbeat messages such as:

```text
⠙ Codex review running... 00:01:15
```

mean the reviewer is still alive.

Do not cancel, restart, or duplicate the review while heartbeat messages
continue.

## Review objective

Treat this review as an independent quality gate, not as a proof that the
implementation is free of every possible defect.

The objective is to identify and resolve high-confidence defects introduced by
the current change without entering an unbounded review-driven patch loop.

Prefer correctness and structural fixes over repeatedly extending local
fallbacks or special cases.

## Handling a review

When a review completes:

1. Read the entire review before editing code.
2. Verify every actionable finding against the implementation.
3. Classify each finding as:
   - valid and blocking;
   - valid and in scope;
   - valid but better handled as follow-up work;
   - unsupported or false positive.
4. Fix valid in-scope findings as a batch where practical.
5. Do not blindly apply unsupported or false-positive findings.
6. Add focused regression coverage for defects that are fixed.
7. Run targeted verification for the affected area after the batch of fixes.

Do not run the complete repository verification suite after every individual
finding unless repository policy or the nature of the change specifically
requires it.

## Bounded review loop

Use the following review budget.

### Pass 1 — full independent review

Run the normal final review after the implementation and initial verification.

Evaluate all findings before making review-driven changes.

If there are valid in-scope findings:

- fix them as a batch;
- run focused tests and relevant verification;
- then run one additional full review.

### Pass 2 — convergence review

Use the second review to detect regressions introduced by the first batch of
review fixes and important issues missed by the first pass.

If Pass 2 has no blocking or clearly in-scope defects, stop the review loop.

P2/P3 findings that are valid but peripheral, pre-existing, speculative, or
better suited to separate work do not require another review cycle. Record or
report them as follow-up work when appropriate.

If Pass 2 exposes additional defects in the same mechanism that was repeatedly
patched during Pass 1, do not immediately continue adding special cases.
Reassess the design first.

Examples of a repeated-pattern signal include successive findings involving:

- additional branches or control-flow forms;
- increasingly specific syntax cases;
- repeated fallback-parser extensions;
- repeated state-machine exceptions;
- repeated compatibility exceptions around the same abstraction.

When this occurs, ask whether the implementation should instead:

- reuse an existing parser, IR, CFG, or semantic abstraction;
- narrow the fallback's responsibility;
- replace multiple special cases with one structural invariant;
- fail conservatively rather than emulate increasingly complex behavior.

### Pass 3 — exceptional final review

A third full review is allowed only when Pass 2 caused a meaningful code change
because of:

- a P0 or P1 finding;
- a clear regression introduced by review-driven changes;
- a structural correction needed to resolve a repeated class of defects;
- another correctness issue that would reasonably block merge.

After Pass 3, do not start another full review merely because a new non-blocking
P2/P3 observation is available.

If a blocking defect still remains after Pass 3, stop the automatic loop and
report the unresolved issue rather than continuing indefinitely.

## Verification strategy

Use verification proportionally.

During review-driven iteration:

```text
finding batch
    ↓
focused regression tests
    ↓
affected package / subsystem verification
```

Before declaring the task complete:

```text
review fixes settled
    ↓
required full verification
    ↓
final review state assessed
```

Avoid repeatedly running expensive repository-wide validation when the code has
only received a small local fix and focused verification is sufficient for the
current iteration.

Repository policy may still require specific full checks before completion;
this skill does not override those requirements.

## Severity and completion

Always resolve before completion:

- P0 findings;
- P1 findings;
- clear regressions introduced by the current implementation;
- high-confidence correctness defects directly within the requested scope.

Usually resolve P2 findings when they are directly caused by the current change
and can be fixed without substantially expanding the implementation scope.

A valid P2/P3 finding may be deferred when fixing it would:

- substantially broaden the task;
- introduce a new subsystem or abstraction;
- trigger a continuing chain of peripheral edge-case work;
- address behavior that is not required for the current change.

When deferring a valid finding, preserve enough detail for follow-up work and
mention it in the implementation summary.

Completion does not require forcing the reviewer to return zero findings.
Completion requires that no verified blocking finding remains and that the
change has received the required bounded independent review and verification.

## Failure

If the review cannot run because Codex is unavailable, authentication is
missing, or the local environment cannot launch it:

- do not repeatedly retry it;
- report the review failure in the final implementation summary;
- do not treat infrastructure failure as a code-review finding.
