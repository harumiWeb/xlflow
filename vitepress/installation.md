# Installation

## Recommended for Windows users

### Quick install

If you want the fastest path:

```powershell
irm https://harumiweb.github.io/xlflow/install.ps1 | iex
xlflow doctor
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

Verify archive integrity against `checksums.txt`:

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
