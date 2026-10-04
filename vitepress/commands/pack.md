# xlflow pack

Build a release `.xlsm` artifact from source, with either a workbook template or the fixed blank-workbook profile.

```bash
xlflow pack --out build/Release.xlsm

# No template: ThisWorkbook + one Sheet1
xlflow pack --blank --out build/Release.xlsm
```

`--blank` is mutually exclusive with `--template`. It requires `ThisWorkbook.bas` and `Sheet1.bas`, rejects additional document modules, and uses `[pack.blank].code_page` for project text encoding (default `1252`; use `932` for Japanese). VBA project LCID records always use the MS-OVBA-required value `0x00000409`.

Blank mode can create empty UserForms and common controls from canonical
`src/forms/specs/<Name>.yaml` (or JSON/YML) specs. In the default sidecar mode,
code comes from `src/forms/code/<Name>.bas`, falling back to matching `.frm`
code or empty code. In `frm` mode, matching `.frm` code is required. Designer
state always comes from the spec; `.frx` is not required. Forms references are
added automatically. Frames can contain common controls and nested Frames.
MultiPage/Page/TabStrip and BMP/JPEG Image pictures are supported. Custom ActiveX and
unsupported persisted list state fail before publication.
Image assets use `picture.path` or explicit `picture.remove: true`; see the
[UserForm specification](../reference/userform-spec).

`pack` is the stable, pure-Go release path. It is cross-platform and
Excel-independent: standard/class topology comes from source, workbook and
document topology comes from the template. Template mode also adds supported
UserForms from canonical specs and edits supported properties of existing
Designers, including Frame/common-control additions, removals, replacements,
parent changes, sibling ordering, and Image picture replacement/removal. The
spec lists every control; omitted properties retain template values. Existing
forms without a spec retain legacy code-only updates.
`pack` does not compile or execute VBA, so successful JSON keeps
`pack.backend = "pure-go"` and `pack.vbe_validation = "not_performed"`.

Omitted template forms are preserved by default. To make the source form set
authoritative, including deleting the final form, configure:

```toml
[pack]
userform_topology = "source"
```

This requires canonical specs for all forms remaining in source. The default
is `"template"`; blank mode always uses the source form set. Root-dimension edits and
unsupported property edits fail before publication with specific
`pack_userform_generation_*` or `pack_userform_edit_*` errors.

Use `parentId` or nested `controls` for hierarchy and `zIndex` for sibling order.
Coordinates are relative to the owning parent. When moving a control, supply
new coordinates if needed; omitted coordinates keep their numeric values.

For example, this control list creates a Frame with a child TextBox:

```yaml
controls:
  - id: frame_main
    name: FrameMain
    type: Frame
    controls:
      - id: input
        name: Input
        type: TextBox
        left: 12
        top: 18
        text: Hello
```

It validates every managed `.bas`, `.cls`, and `.frm` file as UTF-8 without BOM before source planning or binary generation. Invalid input returns `source_encoding_invalid`; run `xlflow encoding check`, then use `xlflow encoding convert --from cp932` only for eligible CP932 source. Open the resulting artifact in real Excel to compile/run a sentinel macro before publishing. See the repository's [pack specification](https://github.com/harumiWeb/xlflow/blob/main/docs/specs/pack-command.md) for release-gate details.

Use `build` when Excel/VBIDE-backed reconstruction and compile validation are
required. Use `push` to synchronize the complete source tree into the configured
development workbook or live session.

## Common failures

Unsupported extensions, missing templates, source-encoding or source-plan failures, an aliased `--out` path, or a locked destination return structured errors. UserForm `.frm` files and sidecar code are both included in encoding validation, while binary `.frx` files are excluded. The artifact is published atomically through a temporary sibling file, so a failed `pack` never corrupts a previously valid output. Keep the source and template under version control and never treat a generated artifact as the source of truth.

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow pack` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

## Prerequisites

Check the project configuration and run `xlflow doctor --json` before workbook-backed operations. Source-only commands can run without Excel; commands that read or mutate a workbook require Windows Excel and VBIDE access.

## What this command reads and changes

The command reads the inputs and configuration described in its syntax and examples. Treat source files, the saved workbook, and a live session as separate states; add `--session` when the live workbook is authoritative. Any mutation is reversible only when a backup or explicit session save boundary exists.

## Effect on source-of-truth state

Use `xlflow status --json` before and after the command. A source edit normally requires `push`; a workbook edit normally requires `pull`; a dirty live session requires `save --session` or an intentional discard.

## Common workflows

Combine this command with the relevant [source/workbook/session workflow](../concepts/workbook-session-source), and use `--json` in scripts and agent loops.

## Common failures

Read the structured `error.code`, exit code, and recovery metadata instead of scraping terminal text. The [symptom-oriented troubleshooting guide](../help/troubleshooting) maps installation, execution, session, VS Code, and WSL failures to recovery steps.
