---
name: xlflow
description: Use when an AI agent needs to edit, test, debug, or validate an Excel VBA workbook with xlflow. Provides the safe source-to-workbook proof loop and routes specialized work to focused references.
---

# xlflow Skill

## Purpose

Use xlflow as the proof loop for Excel VBA work. Source changes are not complete
until xlflow has imported the relevant source and produced evidence appropriate to
the behavior: diagnostics, focused tests, macro execution, or workbook inspection.

This Skill is an orchestration contract. Use xlflow help, structured output, and
the command documentation for command syntax, flags, result schemas, and detailed
diagnostics. Load a specialized reference before taking on the matching workflow.

## Quick Orientation

When xlflow is unfamiliar, a needed capability is not in this map, or installed
behavior may differ from prior knowledge, run `xlflow --help` before choosing a
command. Use command-specific help when the map identifies a command but its
required inputs or safe options are unclear.

Use this map to choose the first command for a goal. It is a discovery aid, not a
complete CLI reference; read structured output and load the named reference for
the next decision.

| Goal                                      | Start with                                                                              | Use it when                                                                                           |
| ----------------------------------------- | --------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------- |
| Discover available capabilities           | `xlflow --help`                                                                         | xlflow is unfamiliar or the needed feature is unknown                                                 |
| Check source, workbook, and session state | `xlflow status --json`                                                                  | before workbook work or after a failure                                                               |
| Diagnose Excel, COM, or VBIDE setup       | `xlflow doctor --json`                                                                  | automation cannot be trusted                                                                          |
| Export workbook VBA into source           | `xlflow pull --session --json`                                                          | the workbook is authoritative or freshness is unclear                                                 |
| Find a runnable entrypoint                | `xlflow macros --session --json`                                                        | the macro name is not already proven                                                                  |
| Check source before Excel import          | `xlflow lint --json` and `xlflow analyze --json`                                        | after source edits                                                                                    |
| Estimate refactor blast radius            | `xlflow impact Module.Procedure --json`                                                 | changing existing behavior; load [code analysis](references/code-analysis.md)                         |
| Import edited source into Excel           | `xlflow push --fast --session --no-save --json`                                         | source validation passes                                                                              |
| Run repeatable behavior checks            | `xlflow test --session --no-save --json`                                                | tests cover the behavior; load [testing](references/testing.md)                                       |
| Run a macro with diagnostics              | `xlflow run Macro.Name --diagnostic --headless --session --json`                        | tests do not cover the intended behavior                                                              |
| Inspect values, formulas, or styles       | `xlflow inspect range --sheet Result --address A1:F20 --include-style --session --json` | an observable cell result is required                                                                 |
| Inspect rendered workbook output          | `xlflow export-image --sheet Sheet1 --range A1:M21 --out preview.png --session --json`  | layout, styles, charts, or appearance are requirements                                                |
| Find headless UI boundaries               | `xlflow inspect-gui --json`                                                             | headless execution may cross dialogs, forms, or UI hooks; load [UI guidance](references/xlflow-ui.md) |
| Persist or close verified work            | `xlflow save --session --json`; `xlflow session stop --json`                            | only after the proof loop is complete                                                                 |

Treat `inspect workbook` as workbook metadata and `inspect range` as value or
style evidence. Treat `export-image` as the required proof when the acceptance
criteria include rendered appearance; inspect the generated image before saving.

## Core Invariants

- Determine the authoritative source before editing. When source and workbook
  state disagree or their freshness is unclear, establish the source of truth
  through xlflow before changing either copy.
- For normal iterative work, use one xlflow session from orientation through
  verification. Start a managed session for a closed workbook; attach to the
  configured workbook when the user already has it open in Excel. Do not create a
  separate Excel instance for a workbook the user expects to remain open.
- Edit source-controlled artifacts, not VBA embedded directly in the binary
  workbook. Use workbook-state edits only when the task is explicitly workbook
  state or visual work and does not change production VBA.
- Treat unsaved session state as newer than the saved workbook. Inspect the live
  session or save it before relying on a disk-backed observation.
- Use structured JSON output for agent-facing commands. Read returned status,
  diagnostics, warnings, and recovery state rather than inferring success from
  terminal text or the absence of errors.
- Write code that passes xlflow lint and analysis without unnecessary
  suppression. Use a local suppression only for an intentional, understood
  exception; never suppress a finding xlflow marks as preflight-blocking or
  non-suppressible.
- If `recovery.required=true` or a recovery-required error is returned, stop
  normal workbook operations. Do not save or blindly retry. Load
  [references/recovery.md](references/recovery.md) and complete its workflow
  before resuming.
- Treat unresolved or ambiguous structural analysis as uncertainty. An empty
  confirmed caller list is not proof that no dependency exists when unresolved
  edges remain.
- Unattended agent workflows must not depend on raw interactive Excel UI. Keep
  GUI entrypoints thin and separate their core business logic so it can be tested
  or run headlessly when needed.

## Orientation Decisions

Before editing, answer these questions from project configuration, source, and
structured xlflow state:

1. Is source or the workbook authoritative, or must a pull reconcile the source
   tree first?
2. Is there an existing user-owned workbook that must be attached, or should
   xlflow start the managed development session?
3. Is the requested change source behavior, workbook state, form design,
   formulas, UI behavior, or a combination that requires a specialized
   reference?
4. Does the change alter existing behavior enough to require structural impact
   analysis before source inspection or test selection?
5. What observable result will prove the task: a focused test, macro result,
   diagnostic, worksheet value, form state, or rendered workbook output?

Do not begin a normal source edit while these decisions are unknown. When the
workbook is authoritative, pull into the configured source artifacts first; when
source is authoritative, push it through xlflow rather than manually importing
VBA in the VBE.

## Quick Start: Normal Source Change

For a typical source-backed workbook change, start with this session-first loop.
Replace the focused test, macro, and inspection target with evidence returned by
xlflow or the project contract.

```bash
xlflow status --json
# If its structured result reports recovery required, stop here and load
# references/recovery.md. Do not start, attach, pull, push, run, test, inspect,
# or save until recovery is complete.
# Use this for a closed workbook that xlflow should manage:
xlflow session start --json
# Instead, if the user already has the configured workbook open:
# xlflow session attach --json

# Run this before editing only when the workbook is authoritative or freshness is unclear.
xlflow pull --session --json

# Edit source-controlled artifacts, then validate them before importing to Excel.
xlflow lint --json
xlflow analyze --json
xlflow push --fast --session --no-save --json

# Prefer this focused existing test while iterating.
xlflow test --filter Module.TestName --session --no-save --json
# Instead, when no relevant test exists, run the intended entrypoint with diagnostics.
xlflow run --diagnostic --headless --session --json

# Inspect the live result when workbook state or appearance is part of the proof.
xlflow inspect workbook --session --json

# When tests exist, run broader affected verification, then persist deliberately.
xlflow test --session --no-save --json
xlflow save --session --json
xlflow session stop --json
```

Run `xlflow doctor --json` when Excel, COM, VBIDE access, or macro execution
cannot be trusted. Do not run both session start and attach: choose attach for
the user-open configured workbook and start for a new managed session. Omit pull
when source is already established as authoritative; do not omit it merely to
avoid resolving an unknown source-of-truth conflict.

This is the executable default, not a complete command reference. Load the
specialized reference when its task changes command selection, inputs, or proof.

## Validation Decisions

Use the smallest proof that exercises the requested behavior, then expand it:

- Select a focused existing test when one covers the changed behavior.
- Run the intended macro when tests do not exercise the behavior.
- Inspect live workbook state when the proof has not been saved; inspect the
  saved workbook only after it is the state intended for review.
- Include visual inspection when layout, styles, charts, or form rendering are
  requirements rather than incidental implementation details. Export and inspect
  an image before saving when rendered appearance is an acceptance criterion.
- Run the wider affected test set before finalizing a behavior or dependency
  change, especially after a public or shared-code refactor.

If diagnostics identify the cause, fix the source and repeat the proof loop. If
they do not, load the debugging reference before adding instrumentation. Never
trade structured diagnostics for a raw GUI dialog unless a human explicitly owns
that interaction.

## Canonical Development Loop

1. **Orient.** Read project configuration, inspect the relevant source and
   current xlflow state, determine source authority, and use the Quick Start
   commands to start or attach the appropriate session.
2. **Analyze structural impact when relevant.** Before changing existing behavior,
   load [references/code-analysis.md](references/code-analysis.md) when the
   change may affect callers, callees, modules, types, or tests.
3. **Plan.** Choose the smallest source change, relevant proof, and any required
   workbook observation before editing.
4. **Edit.** Change the authoritative source artifacts. Pull first when the
   workbook is authoritative.
5. **Validate source.** Run lint and analysis before importing source into Excel,
   then resolve relevant findings.
6. **Push/import.** Import validated source into the active development workbook
   and keep the proof loop reversible until evidence supports saving.
7. **Prove focused behavior.** Prefer the focused relevant test; otherwise run
   the intended macro or behavior with an appropriate unattended or human-led
   workflow.
8. **Inspect the workbook result when relevant.** Observe the live workbook or a
   saved snapshot according to the state being proved, including visual output
   when appearance is part of the requirement.
9. **Broaden verification.** Run the wider affected test set and re-check
   structural impact after a dependency-changing refactor.
10. **Persist deliberately.** Save only when the verified workbook changes should
    persist, then stop the session. Recovery-required state is the exception:
    follow the recovery reference instead of saving.

Use isolated non-session commands only for one-shot CI-style verification,
release checks, suspicious session state, or when the user explicitly requests
that Excel not stay open.

## Scope Boundaries

- Use xlflow to discover current capabilities instead of preserving command
  history in this Skill. Deprecated behavior belongs in migration material and
  CLI diagnostics, not in agent instructions.
- Do not treat a command's progress display as evidence. Wait for its structured
  result, then use that result to choose the next safe action.
- Do not substitute a static review of VBA for import and behavioral proof.
- Do not use broad structural analysis mechanically for comments, formatting, or
  unrelated workbook styling.

## Dispatch to Specialized References

| Task or signal                                                             | Load this reference                             |
| -------------------------------------------------------------------------- | ----------------------------------------------- |
| Writing tests, selecting tests, test metadata, hooks, or test failures     | [testing.md](references/testing.md)             |
| Formula-driven behavior, defined names, or sheet-layout formulas           | [formulas.md](references/formulas.md)           |
| UserForm design, code-behind authority, inspection, snapshots, or rebuilds | [forms.md](references/forms.md)                 |
| Dialogs, file pickers, headless UI, or interactive-vs-unattended behavior  | [xlflow-ui.md](references/xlflow-ui.md)         |
| Runtime or compile diagnostics that do not identify the cause              | [debugging.md](references/debugging.md)         |
| Recovery-required state or uncertain workbook termination                  | [recovery.md](references/recovery.md)           |
| Callers, callees, dependencies, refactors, blast radius, or affected tests | [code-analysis.md](references/code-analysis.md) |
| Lookups over unsorted data, criteria strings, structural inserts, or sorting | [object-model-traps.md](references/object-model-traps.md) |

Use the reference before editing when its subject changes source authority,
safety, test selection, or the evidence needed to prove the task.

## Evidence of Completion

Report and preserve evidence that answers all of these questions:

- Which source and workbook state was authoritative, and which artifacts changed?
- Which xlflow validation and focused behavior checks passed?
- What observable workbook result proves the requested behavior when workbook
  state or appearance matters?
- Which broader verification covered the change and its structural impact?
- Was the verified workbook saved and the session stopped, or is a recovery or
  other user-owned follow-up still required?

Do not declare VBA work complete from source review alone. If the available
evidence cannot prove the requested behavior safely, state the remaining gap.
