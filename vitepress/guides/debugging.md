# Run and debug VBA

Debugging begins by separating “xlflow could not start the macro” from “the macro ran but produced the wrong workbook state.” Discover the exact target first:

```bash
xlflow macros --json
xlflow run Module.Procedure --diagnostic --json
```

Use `--diagnostic` for agents and CI. It captures compile/runtime information and prevents a VBE dialog from becoming the only failure signal. Use `--interactive` only when a person is watching Excel and must complete a UserForm, file picker, or message box. For unattended work, replace raw `MsgBox`/`InputBox` calls with the [XlflowUI wrappers](../reference/xlflow-ui).

For a compile or source-preflight failure, fix the reported source location and rerun `lint`/`analyze` before pushing again. For `macro_not_found`, use the `qualified_name` from `macros`. For `macro_timeout`, inspect dialogs and session state before retrying. When execution succeeds but the result is wrong, run `xlflow inspect range --sheet ... --address ... --json` or `export-image` to compare an observable workbook result instead of relying on Excel's last visible screen.
