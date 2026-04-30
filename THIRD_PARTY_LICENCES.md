# Third Party Licences

This document lists third-party Go modules used by xlflow.

The list is based on `go.mod`. Direct dependencies are modules explicitly required by xlflow. Indirect dependencies are transitive dependencies resolved by the Go module graph.

This file is provided for attribution and licence review. It is not a substitute for the original licence files distributed by each upstream project.

## Direct dependencies

| Module | Version | Licence |
|---|---:|---|
| `github.com/BurntSushi/toml` | `v1.6.0` | MIT |
| `github.com/charmbracelet/bubbletea` | `v1.3.10` | MIT |
| `github.com/charmbracelet/lipgloss` | `v1.1.0` | MIT |
| `github.com/spf13/cobra` | `v1.10.2` | Apache-2.0 |
| `github.com/xuri/excelize/v2` | `v2.10.1` | BSD-3-Clause |

## Indirect dependencies

| Module | Version | Licence |
|---|---:|---|
| `github.com/aymanbagabas/go-osc52/v2` | `v2.0.1` | MIT |
| `github.com/charmbracelet/colorprofile` | `v0.2.3-0.20250311203215-f60798e515dc` | MIT |
| `github.com/charmbracelet/x/ansi` | `v0.10.1` | MIT |
| `github.com/charmbracelet/x/cellbuf` | `v0.0.13-0.20250311204145-2c3ea96c31dd` | MIT |
| `github.com/charmbracelet/x/term` | `v0.2.1` | MIT |
| `github.com/erikgeiser/coninput` | `v0.0.0-20211004153227-1c3628e74d0f` | MIT |
| `github.com/inconshreveable/mousetrap` | `v1.1.0` | Apache-2.0 |
| `github.com/lucasb-eyer/go-colorful` | `v1.2.0` | MIT |
| `github.com/mattn/go-isatty` | `v0.0.20` | MIT |
| `github.com/mattn/go-localereader` | `v0.0.1` | MIT |
| `github.com/mattn/go-runewidth` | `v0.0.16` | MIT |
| `github.com/muesli/ansi` | `v0.0.0-20230316100256-276c6243b2f6` | MIT |
| `github.com/muesli/cancelreader` | `v0.2.2` | MIT |
| `github.com/muesli/termenv` | `v0.16.0` | MIT |
| `github.com/richardlehane/mscfb` | `v1.0.6` | MIT |
| `github.com/richardlehane/msoleps` | `v1.0.6` | MIT |
| `github.com/rivo/uniseg` | `v0.4.7` | MIT |
| `github.com/spf13/pflag` | `v1.0.9` | BSD-3-Clause |
| `github.com/tiendc/go-deepcopy` | `v1.7.2` | MIT |
| `github.com/xo/terminfo` | `v0.0.0-20220910002029-abceb7e1c41e` | MIT |
| `github.com/xuri/efp` | `v0.0.1` | BSD-3-Clause |
| `github.com/xuri/nfp` | `v0.0.2-0.20250530014748-2ddeb826f9a9` | BSD-3-Clause |
| `golang.org/x/crypto` | `v0.48.0` | BSD-3-Clause |
| `golang.org/x/net` | `v0.50.0` | BSD-3-Clause |
| `golang.org/x/sys` | `v0.41.0` | BSD-3-Clause |
| `golang.org/x/text` | `v0.34.0` | BSD-3-Clause |

## Licence notes

### MIT

The MIT Licence is a permissive licence. It generally allows use, copying, modification, distribution, and sublicensing, provided that the copyright notice and permission notice are included in copies or substantial portions of the software.

### BSD-3-Clause

The BSD 3-Clause Licence is a permissive licence. It generally allows redistribution and use in source and binary forms, with or without modification, provided that the licence conditions are met. It also includes a non-endorsement clause.

### Apache-2.0

The Apache License 2.0 is a permissive licence. It generally allows use, reproduction, modification, distribution, and sublicensing, subject to its conditions. It also includes an express patent licence and NOTICE-related requirements when applicable.

## Review checklist before release

Before publishing a binary release, verify this file against the module cache or an automated licence scanner, for example:

```bash
go list -m all
```

For a stricter release check, use a licence scanning tool such as `go-licenses` or an equivalent SBOM/licence scanner and compare the result with this file.

```bash
go install github.com/google/go-licenses@latest
go-licenses report ./cmd/xlflow
```

If dependency versions change, update this file together with `go.mod` and `go.sum`.
