---
name: final-review
description: >
  Run the repository's final independent Codex review after completing an
  implementation and its required verification. Use this before
  declaring implementation work complete.
---

# Final Codex Review

Use this skill after implementation and required verification are complete.

Run:

```powershell
rtk proxy task review:codex
```

The command intentionally suppresses the nested reviewer's intermediate
reasoning and tool output.

While it runs, short heartbeat messages such as:

```text
⠙ Codex review running... 00:01:15
```

mean the reviewer is still alive.

Do not cancel, restart, or duplicate the review while heartbeat messages
continue.

## Handling the result

When the review completes:

1. Read the final review.
2. Verify each actionable finding against the implementation.
3. Fix valid findings that are within scope.
4. Do not blindly apply unsupported or false-positive findings.
5. Rerun relevant verification after fixes.
6. If review-driven changes were made, run:

   ```powershell
   rtk proxy task review:codex
   ```

   again.

Do not declare the implementation complete while a verified actionable finding
remains.

## Failure

If the review cannot run because Codex is unavailable, authentication is
missing, or the local environment cannot launch it:

- do not repeatedly retry it;
- report the review failure in the final implementation summary;
- do not treat infrastructure failure as a code-review finding.
