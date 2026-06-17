# Installation

## Recommended for Windows users

### Quick install

If you want the fastest path:

```powershell
irm https://harumiweb.github.io/xlflow/install.ps1 | iex
xlflow new
```

### Uninstall

Use the same script in file mode to remove the PATH entry and the `%LOCALAPPDATA%\xlflow` installation directory:

```powershell
irm https://harumiweb.github.io/xlflow/install.ps1 -OutFile .\install.ps1
powershell -ExecutionPolicy Bypass -File .\install.ps1 -Action uninstall
```

### winget

If you just want to use xlflow on Windows with a package manager, install it with winget:

```powershell
winget install HarumiWeb.Xlflow
```

Use `upgrade` to update an existing installation:

```powershell
winget upgrade HarumiWeb.Xlflow
```

winget availability may lag behind a GitHub Release while the manifest is submitted and accepted upstream.

Scoop is also supported:

```powershell
scoop bucket add harumiweb https://github.com/harumiWeb/scoop-bucket
scoop install xlflow
```

If you prefer a manual install, download the Windows release archive from [GitHub Releases](https://github.com/harumiWeb/xlflow/releases).

The Windows ZIP contains both `xlflow.exe` and `xlflow-excel-bridge.exe`.

Verify Windows archive integrity against `checksums.txt`:

```powershell
Get-FileHash .\xlflow_windows_x86_64.zip -Algorithm SHA256
certutil -hashfile .\xlflow_windows_x86_64.zip SHA256
```

Verify GitHub artifact attestation when available:

```powershell
gh attestation verify .\xlflow_windows_x86_64.zip --repo harumiWeb/xlflow
```

These checks validate artifact integrity and provenance metadata. They are not Windows Authenticode signing.

The `.NET` bridge executable avoids PowerShell execution policy, but it can still be blocked by AppLocker, WDAC, Defender or EDR policy, antivirus reputation, or other unsigned-executable controls in managed Windows environments.

Linux release archives are built separately with native CGO on Ubuntu and contain the Go CLI only. Verify `xlflow_linux_x86_64.tar.gz` against `checksums-linux.txt`; it does not include `xlflow-excel-bridge.exe`.

## WSL development frontend

Install xlflow on both Windows and WSL. Use the Windows installer, winget, Scoop, or Windows release archive for `xlflow.exe` plus `xlflow-excel-bridge.exe`.

From a WSL shell, install the Linux frontend with:

```bash
curl -fsSL https://harumiweb.github.io/xlflow/install.sh | sh
```

To remove the WSL frontend installed by this script:

```bash
curl -fsSL https://harumiweb.github.io/xlflow/install.sh | sh -s -- --uninstall
```

The installer writes `xlflow` to `$HOME/.local/bin` by default. Override that with `XLFLOW_INSTALL_DIR` or `--install-dir`.

From a WSL shell in a source checkout, install the current frontend build with:

```bash
task wsl-install
```

Store projects under a Windows-mounted path:

```bash
cd /mnt/c/dev/my-vba-project
xlflow doctor --json
```

Excel-related commands automatically delegate to Windows xlflow. Source-only commands such as `lint`, `fmt`, and `analyze` stay in WSL.

If WSL cannot discover `xlflow.exe` through Windows PATH interoperability, configure it explicitly:

```bash
export XLFLOW_WINDOWS_EXE='C:\Users\me\AppData\Local\xlflow\xlflow.exe'
```

WSL-only paths such as `/home/user/project` are unsupported for Excel delegation. Move the project under `/mnt/c`, `/mnt/d`, or another Windows-mounted drive.

xlflow forwards its own runtime environment variables through `WSLENV`. If workbook VBA depends on additional custom WSL environment variables, add those variable names to `WSLENV` so the Windows Excel process can read them.

## Go developers

```bash
go install github.com/harumiWeb/xlflow/cmd/xlflow@latest
```

`go install` installs `xlflow` only. It does not install the packaged `.NET` bridge sidecar `xlflow-excel-bridge.exe`.

If you need `--bridge dotnet`, use the Windows release archive or build/install the bridge separately from a source checkout, for example with `task install`.

Verify the install:

```bash
xlflow version
xlflow --help
```

For source checkout development:

```bash
go run ./cmd/xlflow --help
```
