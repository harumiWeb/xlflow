param(
    [ValidateSet("install", "uninstall")]
    [string]$Action = "install",

    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "xlflow"),

    [string]$Owner = "harumiWeb",

    [string]$Repo = "xlflow"
)

$ErrorActionPreference = "Stop"

function Write-Info {
    param([string]$Message)

    Write-Output "[xlflow] $Message"
}

function Get-ReleaseApiUrl {
    param(
        [string]$RepoOwner,
        [string]$RepoName
    )

    return "https://api.github.com/repos/$RepoOwner/$RepoName/releases/latest"
}

function Get-ReleasePageUrl {
    param(
        [string]$RepoOwner,
        [string]$RepoName
    )

    return "https://github.com/$RepoOwner/$RepoName/releases/latest"
}

function Get-InstallerHeaders {
    return @{
        "Accept" = "application/vnd.github+json"
        "User-Agent" = "xlflow-install-script"
    }
}

function Get-UserPathEntries {
    $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
    if ([string]::IsNullOrWhiteSpace($userPath)) {
        return @()
    }

    return @($userPath.Split([IO.Path]::PathSeparator, [System.StringSplitOptions]::RemoveEmptyEntries))
}

function Set-UserPathEntries {
    param([string[]]$Entries)

    $normalized = [System.Collections.Generic.List[string]]::new()
    foreach ($entry in $Entries) {
        if ([string]::IsNullOrWhiteSpace($entry)) {
            continue
        }

        $candidate = $entry.Trim()
        $duplicate = $false
        foreach ($existing in $normalized) {
            if ([string]::Equals($existing, $candidate, [System.StringComparison]::OrdinalIgnoreCase)) {
                $duplicate = $true
                break
            }
        }

        if (-not $duplicate) {
            $normalized.Add($candidate)
        }
    }

    [Environment]::SetEnvironmentVariable("Path", ($normalized -join [IO.Path]::PathSeparator), "User")
}

function Add-PathEntry {
    param([string]$PathEntry)

    $entries = Get-UserPathEntries
    foreach ($entry in $entries) {
        if ([string]::Equals($entry, $PathEntry, [System.StringComparison]::OrdinalIgnoreCase)) {
            if (-not $env:Path.Split([IO.Path]::PathSeparator).Where({ [string]::Equals($_, $PathEntry, [System.StringComparison]::OrdinalIgnoreCase) }, "First")) {
                $env:Path = "$PathEntry$([IO.Path]::PathSeparator)$env:Path"
            }
            return $false
        }
    }

    Set-UserPathEntries -Entries ($entries + $PathEntry)

    if (-not $env:Path.Split([IO.Path]::PathSeparator).Where({ [string]::Equals($_, $PathEntry, [System.StringComparison]::OrdinalIgnoreCase) }, "First")) {
        $env:Path = "$PathEntry$([IO.Path]::PathSeparator)$env:Path"
    }

    return $true
}

function Remove-PathEntry {
    param([string]$PathEntry)

    $entries = Get-UserPathEntries
    $updated = [System.Collections.Generic.List[string]]::new()
    $removed = $false

    foreach ($entry in $entries) {
        if ([string]::Equals($entry, $PathEntry, [System.StringComparison]::OrdinalIgnoreCase)) {
            $removed = $true
            continue
        }

        $updated.Add($entry)
    }

    if ($removed) {
        Set-UserPathEntries -Entries $updated
    }

    $sessionEntries = [System.Collections.Generic.List[string]]::new()
    foreach ($entry in $env:Path.Split([IO.Path]::PathSeparator, [System.StringSplitOptions]::RemoveEmptyEntries)) {
        if ([string]::Equals($entry, $PathEntry, [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }

        $sessionEntries.Add($entry)
    }
    $env:Path = $sessionEntries -join [IO.Path]::PathSeparator

    return $removed
}

function Get-LatestRelease {
    param(
        [string]$RepoOwner,
        [string]$RepoName
    )

    $apiUrl = Get-ReleaseApiUrl -RepoOwner $RepoOwner -RepoName $RepoName
    Write-Info "Fetching latest release metadata from $apiUrl"
    return Invoke-RestMethod -Uri $apiUrl -Headers (Get-InstallerHeaders)
}

function Select-WindowsZipAsset {
    param($Release)

    $assets = @($Release.assets)
    if ($assets.Count -eq 0) {
        throw "Latest release does not contain downloadable assets."
    }

    $preferred = $assets | Where-Object {
        $_.name -ieq "xlflow_windows_x86_64.zip"
    } | Select-Object -First 1

    if ($preferred) {
        return $preferred
    }

    $fallback = $assets | Where-Object {
        $_.name -match "windows" -and $_.name -match "(x86_64|amd64)" -and $_.name -match "\.zip$"
    } | Select-Object -First 1

    if ($fallback) {
        return $fallback
    }

    throw "Could not find a Windows x64 ZIP asset in the latest release."
}

function Expand-ArchiveCompat {
    param(
        [string]$ZipPath,
        [string]$DestinationPath
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $DestinationPath)
}

function Resolve-ArchiveRoot {
    param([string]$ExtractPath)

    $xlflowExe = Get-ChildItem -Path $ExtractPath -Recurse -File -Filter "xlflow.exe" | Select-Object -First 1
    if (-not $xlflowExe) {
        throw "Downloaded archive did not contain xlflow.exe."
    }

    $bridgeExe = Get-ChildItem -Path $ExtractPath -Recurse -File -Filter "xlflow-excel-bridge.exe" | Select-Object -First 1
    if (-not $bridgeExe) {
        throw "Downloaded archive did not contain xlflow-excel-bridge.exe."
    }

    if (-not [string]::Equals($xlflowExe.Directory.FullName, $bridgeExe.Directory.FullName, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Downloaded archive placed xlflow.exe and xlflow-excel-bridge.exe in different directories."
    }

    return $xlflowExe.Directory.FullName
}

function Install-Xlflow {
    param(
        [string]$RepoOwner,
        [string]$RepoName,
        [string]$TargetRoot
    )

    $binDir = Join-Path $TargetRoot "bin"
    $tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("xlflow-install-" + [guid]::NewGuid().ToString("N"))

    try {
        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

        $release = Get-LatestRelease -RepoOwner $RepoOwner -RepoName $RepoName
        $asset = Select-WindowsZipAsset -Release $release
        $zipPath = Join-Path $tempRoot $asset.name
        $extractPath = Join-Path $tempRoot "extract"

        Write-Info "Downloading $($asset.name)"
        Invoke-WebRequest -Uri $asset.browser_download_url -Headers (Get-InstallerHeaders) -OutFile $zipPath

        Write-Info "Extracting archive"
        New-Item -ItemType Directory -Path $extractPath -Force | Out-Null
        Expand-ArchiveCompat -ZipPath $zipPath -DestinationPath $extractPath

        $archiveRoot = Resolve-ArchiveRoot -ExtractPath $extractPath

        Write-Info "Installing to $binDir"
        New-Item -ItemType Directory -Path $binDir -Force | Out-Null
        Copy-Item -Path (Join-Path $archiveRoot "*") -Destination $binDir -Recurse -Force

        $pathAdded = Add-PathEntry -PathEntry $binDir
        if ($pathAdded) {
            Write-Info "Added $binDir to the user PATH"
        } else {
            Write-Info "$binDir is already present in the user PATH"
        }

        $xlflowPath = Join-Path $binDir "xlflow.exe"
        Write-Info "Verifying installation"
        $versionOutput = & $xlflowPath version 2>&1
        if ($versionOutput) {
            $versionOutput | Write-Output
        }

        if ($LASTEXITCODE -ne 0) {
            throw "xlflow.exe version exited with code $LASTEXITCODE."
        }

        Write-Output ""
        Write-Info "Install complete."
        Write-Info "Next step: xlflow doctor"
        Write-Info "Release source: $(Get-ReleasePageUrl -RepoOwner $RepoOwner -RepoName $RepoName)"
    } catch {
        Write-Warning $_
        Write-Output ""
        Write-Output "Install failed."
        Write-Output "Manual recovery: download the latest Windows ZIP from $(Get-ReleasePageUrl -RepoOwner $RepoOwner -RepoName $RepoName), extract it into $binDir, and add that directory to your user PATH."
        exit 1
    } finally {
        if (Test-Path -LiteralPath $tempRoot) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Uninstall-Xlflow {
    param([string]$TargetRoot)

    $binDir = Join-Path $TargetRoot "bin"

    try {
        $pathRemoved = Remove-PathEntry -PathEntry $binDir
        if ($pathRemoved) {
            Write-Info "Removed $binDir from the user PATH"
        } else {
            Write-Info "$binDir was not present in the user PATH"
        }

        if (Test-Path -LiteralPath $TargetRoot) {
            Write-Info "Removing $TargetRoot"
            Remove-Item -LiteralPath $TargetRoot -Recurse -Force
            Write-Info "Uninstall complete."
        } else {
            Write-Info "Nothing to remove at $TargetRoot"
        }
    } catch {
        Write-Warning $_
        Write-Output ""
        Write-Output "Uninstall failed."
        Write-Output "Close any running xlflow.exe or xlflow-excel-bridge.exe processes, then remove $TargetRoot manually and delete $binDir from your user PATH if it still exists."
        exit 1
    }
}

switch ($Action) {
    "install" {
        Install-Xlflow -RepoOwner $Owner -RepoName $Repo -TargetRoot $InstallRoot
    }
    "uninstall" {
        Uninstall-Xlflow -TargetRoot $InstallRoot
    }
    default {
        throw "Unsupported action: $Action"
    }
}
