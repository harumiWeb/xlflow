# Command Design

The command surface follows a source-to-workbook workflow:

```text
new/init -> doctor -> pull -> edit source -> push -> lint/analyze -> test/run -> inspect/diff/export-image
                                   \-> backup list -> rollback -> pull
```

Machine consumers should use `--json`. Human output is allowed to be richer and less stable.

See the stable command contract in [`docs/specs/cli-contract.md`](https://github.com/harumiWeb/xlflow/blob/main/docs/specs/cli-contract.md).
