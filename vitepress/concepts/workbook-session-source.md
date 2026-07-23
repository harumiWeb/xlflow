# Workbook, Session, Source

xlflow has three places where the same VBA project can exist. They are deliberately separate so you can tell exactly what will be overwritten.

- **Source** is the editable text under `src/`. This is what Git reviews.
- **Saved workbook** is the `.xlsm` (or `.xlam`/`.xlsb`) file under `build/`. This is what survives after Excel closes.
- **Live session** is the workbook currently open in an Excel process. It exists only while a session is running and can be newer than the file on disk.

xlflow makes those transitions explicit:

```text
source files -- push --> saved workbook on disk
saved workbook on disk -- session start / attach --> live Excel workbook
live Excel workbook -- save --session --> saved workbook on disk

Reverse direction:

saved workbook on disk -- pull --> source files
```

`pull` copies **from workbook to source**. Use it after editing in VBE, or whenever Excel is newer. `push` copies **from source to workbook**. Use it after editing in VS Code. A `session` lets you repeat `push`, `run`, and `test` without reopening Excel; adding `--session` means “use the live copy, not the file on disk.”

When a session-backed command leaves the live workbook newer than disk, xlflow reports save-required metadata. Persist that state with:

```bash
xlflow save --session --json
```

To check the current state of all three layers at once, use `status`:

```bash
xlflow status --json
```

`status` shows whether source files are newer than the saved workbook, whether a live session is active and dirty, and which path xlflow believes is the current source of truth. It does not change anything, so it is the safest first command when you are unsure.

| Situation                                | Authoritative state | Recommended action                                  |
| ---------------------------------------- | ------------------- | --------------------------------------------------- |
| Workbook edited in VBE                   | Workbook            | `xlflow pull`                                       |
| Source edited in VS Code                 | Source              | `xlflow push`                                       |
| Source pushed with `--session --no-save` | Live workbook       | `xlflow save --session`                             |
| Unsure which side is newer               | Unknown             | `status`, create a backup, then choose deliberately |

### A safe default when you are unsure

1. Run `xlflow status --json`.
2. If you need a safety net, run `xlflow backup create --json` (or copy the workbook outside the project).
3. Decide which edit you want to keep, then use exactly one direction: `pull` to preserve Excel/VBE edits, or `push` to preserve source edits.

Do not run `pull` and `push` back-to-back just to “sync” without reading the status: the second command can overwrite the state the first one just made authoritative.
