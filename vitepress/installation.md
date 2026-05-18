# Installation

## Recommended for Windows users

If you just want to use xlflow on Windows, install it with winget:

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

## Go developers

```bash
go install github.com/harumiWeb/xlflow/cmd/xlflow@latest
```

Verify the install:

```bash
xlflow version
xlflow --help
```

For source checkout development:

```bash
go run ./cmd/xlflow --help
```
