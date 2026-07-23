# FAQ

If you have not used xlflow before, start with [Getting Started](../getting-started). These short answers point to the deeper page when the decision can affect workbook data.

## Does xlflow edit my original workbook?

`init` copies an existing workbook into `build/`. Commands target the configured copy unless you explicitly configure another path. Your original workbook remains a fallback, but make a normal backup before any important migration.

## Which side is authoritative?

The newest intentional state is authoritative. Use `status` before choosing `pull` or `push`; use [Source, workbook, and session state](../concepts/workbook-session-source) for the decision table.

## Do I need Excel for linting?

No. `fmt`, `lint`, `analyze`, and many inspection operations are source/file based. Push, run, test, sessions, and COM inspection require Windows Excel.

## Can an agent work from WSL?

Yes. Keep the project on a Windows-mounted path and install both the WSL frontend and Windows bridge. See [Work from WSL](../tutorials/wsl).

## How do I report a useful bug?

Include `xlflow version --verbose --json`, `doctor --json`, the failing command, exit code, sanitized JSON, and whether a managed or external session was active.
