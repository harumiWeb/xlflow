# encoding

Check the encoding of managed VBA source files, or explicitly convert CP932
files to the repository format. Managed `.bas`, `.cls`, and `.frm` files are
UTF-8 without a BOM. `.frx` files are binary UserForm companions and are not
scanned or converted.

## Check

```bash
xlflow encoding check [path...]
xlflow encoding check --json
```

Without paths, the command recursively scans the configured `[src]` roots and
the top-level `tests/` directory. An explicit path must be inside one of those
managed roots. Results are sorted by project-relative path, and UserForm
`.frm` files remain in scope even when `[userform].code_source = "sidecar"`.

The command returns exit code `0` when every selected file is valid. A UTF-8
BOM or invalid UTF-8 returns `source_encoding_invalid` with exit code `1` and
the first invalid byte's offset, line, and byte column. Multiple violations
are listed in `error.details.files`.

## Convert

```bash
xlflow encoding convert --from cp932 [path...]
xlflow encoding convert --from cp932 --json
```

`--from` is required; `cp932` is the only supported conversion source. Valid
UTF-8 without a BOM is left byte-identical. Other selected files must decode
strictly as CP932 and are written as UTF-8 without a BOM. Existing line
endings and file permissions are preserved.

The command reads and validates every selected file and stages every changed
file before replacing any original. Replacements use same-directory temporary
files; a failure during replacement restores originals. Temporary backups are
removed after a successful transaction and no permanent `.bak` files are
created. Decode failures, UTF-16 input, and UTF-8 BOMs are reported without
modifying any file.

Conversion is explicit. `analyze`, `lint`, `fmt`, `check`, `push`, and other
commands never guess or convert a source encoding automatically. For a
machine-readable result, the `source` envelope contains `expected`, `from`,
`to`, `files[]` (`path`, `status`), and `summary` counts. Conversion statuses
include `converted` and `unchanged`; check statuses include `valid`,
`utf8_bom`, and `invalid_utf8`.
