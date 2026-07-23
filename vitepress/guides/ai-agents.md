# Use xlflow with AI agents

Give an agent a stable terminal contract. The agent should be able to answer, after every operation: “which copy did I change, and what evidence says it worked?”

- inspect `xlflow.toml`, `status --json`, and `doctor --json` first;
- edit only source-controlled files under `src/`;
- run `fmt`, `lint`, and `analyze` before Excel operations;
- prefer `--json`, explicit `--diagnostic`, and one managed session;
- use `inspect` and `export-image` for observable verification;
- save explicitly and stop or discard sessions deliberately.

Install the bundled workflow with `xlflow skill install --agent codex` when the agent supports xlflow skills. The complete example—including a failing test, correction, workbook inspection, and Git review—is [Develop with an AI agent](../tutorials/ai-agent).
