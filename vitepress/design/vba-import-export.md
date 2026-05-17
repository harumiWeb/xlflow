# VBA Import and Export

`pull` exports workbook VBA into source files. `push` imports edited source back into the workbook.

Important design points:

- Standard modules, class modules, document modules, and UserForms are handled separately.
- UserForm `.frx` companions remain binary artifacts.
- Folder-aware source layout can synchronize Rubberduck-compatible folder annotations.
- Push uses source preflight to block known modal VBE compile-dialog risks before Excel opens.

Related ADR: [`ADR-0001`](https://github.com/harumiWeb/xlflow/blob/main/docs/adr/ADR-0001-agent-ready-vba-cli-architecture.md)
