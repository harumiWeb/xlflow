# xlflow

Agent-ready VBA development framework.

xlflow turns Excel VBA projects into a CLI-first development workflow:

- export VBA modules from `.xlsm`
- edit VBA as normal source files
- import modules back into Excel
- run macros from the CLI
- lint VBA for safer automation
- return deterministic JSON for AI agents

## MVP Commands

```bash
xlflow new
xlflow new Sales
xlflow new Sales.xlsm
xlflow init Book.xlsm
xlflow doctor --json
xlflow pull --json
xlflow push --json
xlflow run Main.Run --json
xlflow lint --json
```

The MVP uses `xlflow.toml` as its project configuration file. Excel automation is Windows-first and uses PowerShell plus Excel COM.
`xlflow new` only accepts `.xlsm` workbook names because it always creates macro-enabled workbook content.
