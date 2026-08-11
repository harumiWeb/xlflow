# xlflow impact

Show the confirmed caller/callee blast radius of one project-local VBA procedure.

## Usage

```bash
xlflow impact Module.Procedure
xlflow impact Module.Procedure --depth 3
xlflow impact Module.Procedure --direction callers --json
xlflow impact Module.Procedure --path src/modules --json
```

## What it reads

`impact` reads exported VBA source under the configured source roots (or `--path`) and never opens Excel. It resolves only project-local calls with exactly one confident target.

The default `--direction both --depth 1` shows the immediate callers and callees. Use `--depth 0` to identify the target only, or a larger depth to inspect a wider blast radius.

| Option                                   | Description                                         | Default                 |
| ---------------------------------------- | --------------------------------------------------- | ----------------------- |
| `<Module.Procedure>`                     | Exact project-local procedure target.               | required                |
| `--direction callers \| callees \| both` | Traverse upstream, downstream, or both.             | `both`                  |
| `--depth <n>`                            | Maximum number of confirmed call edges to traverse. | `1`                     |
| `--path <dir-or-file>`                   | Restrict source analysis to a directory or file.    | configured source roots |

## AI and scripting workflow

Use JSON when deciding what code to inspect or test after a change:

```bash
xlflow impact InvoiceService.Recalculate --depth 3 --json
```

The `impact` payload separates confirmed `nodes` and `edges` from `uncertainty`. Do not treat `ambiguous`, `unresolved`, `external`, `builtin_like`, or `member_call` counts as confirmed dependencies. Each returned node and edge has a source location, while `cycles` makes recursion explicit. Cycle entries are additive: `nodes` are in canonical directed path order, `edges` align each caller with its next callee (wrapping at the end), and each unique rotation is emitted once. Consumers should preserve the ordered path rather than treating it as an unordered set.

If the selector matches no procedure, xlflow returns `impact_target_not_found`. If it matches multiple stable declarations, it returns `impact_target_ambiguous` with the candidate identities instead of merging their blast radii.

## Next actions

- Use `xlflow inspect calls --from Module.Procedure --json` to inspect individual call syntax.
- Run `xlflow lint` and `xlflow analyze` before `push`.
- Use `xlflow test` to verify the affected behavior after selecting the relevant tests.
