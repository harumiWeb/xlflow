# Installation

## Go install

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

## Scoop

```powershell
scoop bucket add harumiweb https://github.com/harumiWeb/scoop-bucket
scoop install xlflow
```

## GitHub Releases

Download the Windows release archive from [GitHub Releases](https://github.com/harumiWeb/xlflow/releases).

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
