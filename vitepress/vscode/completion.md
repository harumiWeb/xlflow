# Completion and type inference

The language server provides VBA keywords, built-in VBA/Excel/Office/MSForms metadata, project symbols, UserForm controls, and ProgID completion inside `CreateObject`/`GetObject` strings. Late-bound variables can acquire a curated type from known ProgIDs, enabling member completion and hover.

For example, type `CreateObject("Scripting.Dictionary")` and assign it to a variable. xlflow can recognise that well-known ProgID and suggest dictionary members even though VBA would otherwise treat the variable as late-bound. Suggestions are guidance, not a guarantee that a specific Office installation exposes every member; still run the macro to verify behavior.

When completion is missing, confirm the file is a supported source type, the workspace has `xlflow.toml`, and the language server is running. `xlflow type db status` reports generated TypeLib data; restart the server after refreshing type data. See [VS Code troubleshooting](./troubleshooting) for the exact checks.
