# Third Party Licences

This document lists third-party Go modules in the runtime and build dependency closure of `./cmd/xlflow`.

The list is cross-checked against `go list -deps ./cmd/xlflow` and `go-licenses report ./cmd/xlflow`. Direct dependencies are modules explicitly required by xlflow. Indirect dependencies are transitive dependencies reached by the `./cmd/xlflow` package dependency graph.

This file is provided for attribution and licence review. It is not a substitute for the original licence files distributed by each upstream project.

## Direct dependencies

| Module                               |   Version | Licence      | Licence file                                                      |
| ------------------------------------ | --------: | ------------ | ----------------------------------------------------------------- |
| `github.com/BurntSushi/toml`         |  `v1.6.0` | MIT          | `https://github.com/BurntSushi/toml/blob/v1.6.0/COPYING`          |
| `github.com/charmbracelet/bubbletea` | `v1.3.10` | MIT          | `https://github.com/charmbracelet/bubbletea/blob/v1.3.10/LICENSE` |
| `github.com/charmbracelet/lipgloss`  |  `v1.1.0` | MIT          | `https://github.com/charmbracelet/lipgloss/blob/v1.1.0/LICENSE`   |
| `github.com/spf13/cobra`             | `v1.10.2` | Apache-2.0   | `https://github.com/spf13/cobra/blob/v1.10.2/LICENSE.txt`         |
| `github.com/xuri/excelize/v2`        | `v2.10.1` | BSD-3-Clause | `https://github.com/xuri/excelize/blob/v2.10.1/LICENSE`           |

## Indirect dependencies

| Module                                  |                                 Version | Licence      | Licence file                                                              |
| --------------------------------------- | --------------------------------------: | ------------ | ------------------------------------------------------------------------- |
| `github.com/aymanbagabas/go-osc52/v2`   |                                `v2.0.1` | MIT          | `https://github.com/aymanbagabas/go-osc52/blob/v2.0.1/LICENSE`            |
| `github.com/charmbracelet/colorprofile` |  `v0.2.3-0.20250311203215-f60798e515dc` | MIT          | `https://github.com/charmbracelet/colorprofile/blob/f60798e515dc/LICENSE` |
| `github.com/charmbracelet/x/ansi`       |                               `v0.10.1` | MIT          | `https://github.com/charmbracelet/x/blob/ansi/v0.10.1/ansi/LICENSE`       |
| `github.com/charmbracelet/x/cellbuf`    | `v0.0.13-0.20250311204145-2c3ea96c31dd` | MIT          | `https://github.com/charmbracelet/x/blob/2c3ea96c31dd/cellbuf/LICENSE`    |
| `github.com/charmbracelet/x/term`       |                                `v0.2.1` | MIT          | `https://github.com/charmbracelet/x/blob/term/v0.2.1/term/LICENSE`        |
| `github.com/erikgeiser/coninput`        |    `v0.0.0-20211004153227-1c3628e74d0f` | MIT          | `https://github.com/erikgeiser/coninput/blob/1c3628e74d0f/LICENSE`        |
| `github.com/inconshreveable/mousetrap`  |                                `v1.1.0` | Apache-2.0   | `https://github.com/inconshreveable/mousetrap/blob/v1.1.0/LICENSE`        |
| `github.com/lucasb-eyer/go-colorful`    |                                `v1.2.0` | MIT          | `https://github.com/lucasb-eyer/go-colorful/blob/v1.2.0/LICENSE`          |
| `github.com/mattn/go-isatty`            |                               `v0.0.20` | MIT          | `https://github.com/mattn/go-isatty/blob/v0.0.20/LICENSE`                 |
| `github.com/mattn/go-localereader`      |                                `v0.0.1` | MIT          | `https://github.com/mattn/go-localereader/blob/v0.0.1/LICENSE`            |
| `github.com/mattn/go-runewidth`         |                               `v0.0.16` | MIT          | `https://github.com/mattn/go-runewidth/blob/v0.0.16`                      |
| `github.com/muesli/ansi`                |    `v0.0.0-20230316100256-276c6243b2f6` | MIT          | `https://github.com/muesli/ansi/blob/276c6243b2f6/LICENSE`                |
| `github.com/muesli/cancelreader`        |                                `v0.2.2` | MIT          | `https://github.com/muesli/cancelreader/blob/v0.2.2/LICENSE`              |
| `github.com/muesli/termenv`             |                               `v0.16.0` | MIT          | `https://github.com/muesli/termenv/blob/v0.16.0/LICENSE`                  |
| `github.com/richardlehane/mscfb`        |                                `v1.0.6` | Apache-2.0   | `https://github.com/richardlehane/mscfb/blob/v1.0.6/LICENSE.txt`          |
| `github.com/richardlehane/msoleps`      |                                `v1.0.6` | Apache-2.0   | `https://github.com/richardlehane/msoleps/blob/v1.0.6/LICENSE.txt`        |
| `github.com/rivo/uniseg`                |                                `v0.4.7` | MIT          | `https://github.com/rivo/uniseg/blob/v0.4.7/LICENSE.txt`                  |
| `github.com/spf13/pflag`                |                                `v1.0.9` | BSD-3-Clause | `https://github.com/spf13/pflag/blob/v1.0.9/LICENSE`                      |
| `github.com/tiendc/go-deepcopy`         |                                `v1.7.2` | MIT          | `https://github.com/tiendc/go-deepcopy/blob/v1.7.2/LICENSE`               |
| `github.com/xo/terminfo`                |    `v0.0.0-20220910002029-abceb7e1c41e` | MIT          | `https://github.com/xo/terminfo/blob/abceb7e1c41e/LICENSE`                |
| `github.com/xuri/efp`                   |                                `v0.0.1` | BSD-3-Clause | `https://github.com/xuri/efp/blob/v0.0.1/LICENSE`                         |
| `github.com/xuri/nfp`                   |  `v0.0.2-0.20250530014748-2ddeb826f9a9` | BSD-3-Clause | `https://github.com/xuri/nfp/blob/2ddeb826f9a9/LICENSE`                   |
| `golang.org/x/crypto`                   |                               `v0.50.0` | BSD-3-Clause | `https://cs.opensource.google/go/x/crypto/+/v0.50.0:LICENSE`              |
| `golang.org/x/net`                      |                               `v0.53.0` | BSD-3-Clause | `https://cs.opensource.google/go/x/net/+/v0.53.0:LICENSE`                 |
| `golang.org/x/sys`                      |                               `v0.43.0` | BSD-3-Clause | `https://cs.opensource.google/go/x/sys/+/v0.43.0:LICENSE`                 |
| `golang.org/x/text`                     |                               `v0.36.0` | BSD-3-Clause | `https://cs.opensource.google/go/x/text/+/v0.36.0:LICENSE`                |

## Scanner output notes

`go-licenses report ./cmd/xlflow` may report packages under `github.com/harumiWeb/xlflow/...` as `Unknown` until xlflow itself has a root `LICENSE` file. That does not indicate a third-party dependency problem; it means the project licence has not been selected yet.

`github.com/mattn/go-localereader` was reported as `Unknown` by `go-licenses` because the scanner could not find a recognized licence file in the module cache. Review this module manually before the first public release. If necessary, keep this dependency notice as `Unknown` or replace/remove the dependency path by upgrading transitive dependencies.

## Licence notes

### MIT

The MIT Licence is a permissive licence. It generally allows use, copying, modification, distribution, and sublicensing, provided that the copyright notice and permission notice are included in copies or substantial portions of the software.

### BSD-3-Clause

The BSD 3-Clause Licence is a permissive licence. It generally allows redistribution and use in source and binary forms, with or without modification, provided that the licence conditions are met. It also includes a non-endorsement clause.

### Apache-2.0

The Apache License 2.0 is a permissive licence. It generally allows use, reproduction, modification, distribution, and sublicensing, subject to its conditions. It also includes an express patent licence and NOTICE-related requirements when applicable.

## Review checklist before release

Before publishing a binary release, verify this file against the `./cmd/xlflow` dependency closure or an automated licence scanner.

```bash
go list -deps ./cmd/xlflow
```

For a stricter release check, use a licence scanning tool such as `go-licenses` or an equivalent SBOM/licence scanner and compare the result with this file.

```bash
go install github.com/google/go-licenses@latest
go-licenses report ./cmd/xlflow
```

If dependency versions change, update this file together with `go.mod` and `go.sum`.
