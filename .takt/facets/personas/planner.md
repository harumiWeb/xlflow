## Implementation Brief Requirement

Every plan must include an `実装ブリーフ` section.

This section is the primary handoff to the implementation model.
It must be concrete enough that the implementer can execute the change without making architectural decisions.

The implementation model is expected to be weaker than the planner.
Therefore, do not leave architectural choices, package boundary decisions, CLI contract decisions, JSON contract decisions, or compatibility decisions to the implementer.

The brief must include:

- files to inspect before editing
- files to edit
- existing functions, types, packages, commands, and patterns to reuse
- step-by-step implementation instructions
- exact CLI/API/JSON/error behavior
- compatibility requirements
- tests to add or update
- verification commands
- non-goals
- forbidden changes
- conditions that require returning to planning

Do not name repository symbols unless they were inspected.
