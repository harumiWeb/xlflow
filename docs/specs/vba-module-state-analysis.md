# VBA240 module-level mutable-state analysis

`VBA240` is an opt-in, warning-level, non-blocking, batch-only project analysis
for mutable module-level state. It reports structural lifecycle coupling while
leaving subjective fan-in thresholds to a later evidence-based decision.

## Indexed state

The analyzer indexes module declarations from the resolved procedure IR for
standard, class, document, and UserForm modules. `Const` declarations are
classified as constants. Non-constant scalar/reference declarations with no
observed write are reported as read-only configuration; Collection/Dictionary
state remains mutable even when no mutator call is observed.

Each field records distinct procedure readers and writers, recognized
Collection/Dictionary mutators, reachable entry roots, event access, call-cycle
membership, and cached Excel-object classification. Ambiguous, unresolved,
external, and dynamic calls do not establish a project-local dependency.

The index recognizes ordinary and `Set` assignments, `ReDim`, and resolved
Collection/Dictionary mutation methods. A field is keyed by its source file,
module, and case-insensitive declaration name, so class instance state is not
merged with project-global standard-module state.

## Root and cycle model

Confirmed roots are the configured `[project].entry`, public parameterless
standard-module `Sub` procedures, and parser-classified host event handlers.
Root reachability follows uniquely resolved project-local calls. The analyzer
computes call-cycle membership over the same resolved graph. Possible or
uncertain roots are not used to prove a warning or inflate the confirmed-root
counts.

## Diagnostics

One `VBA240` finding is emitted at a mutable field declaration when the field
has one or more of these structural hazards:

- a write or mutation is reachable from at least one confirmed entry root and
  the field is accessed from multiple confirmed entry roots;
- state written by one event handler and read or retained by another event
  invocation;
- both reads and writes inside a call-graph cycle;
- a cached Excel object or mutable Collection/Dictionary crossing one of those
  lifecycle boundaries.

Reader/writer counts alone never trigger a finding. Existing `VBA202` owns
object use-before-`Set` diagnostics, including module-level objects; `VBA240`
may expose that fact in metrics but must not duplicate the diagnostic.

## Metrics output

When `detect_risky_module_state = true`, `analyze` and `check` add an
`analysis_metrics.module_state` object. Its `fields` array contains the indexed
classification and counts, sorted by file, declaration line, and field name.
The metrics are informational and do not affect the command exit code.
The `fields` array includes confirmed root names and is accompanied by a
`procedures` array containing each procedure's resolved module-field `reads`,
`writes`, and Collection/Dictionary `mutators` set.

## Configuration and suppression

Enable with:

```toml
[analyze]
detect_risky_module_state = true
```

Disable with `[analyze].disabled_rules = ["VBA240"]` or suppress a specific
declaration using `xlflow:disable-line VBA240` /
`xlflow:disable-next-line VBA240`.

## Boundaries

This rule does not infer runtime ordering across arbitrary Excel callbacks,
does not treat VBE compilation as evidence, and does not enforce a fan-in
threshold. A future threshold policy must be based on reviewed corpus and
project metrics and must update this specification separately.
