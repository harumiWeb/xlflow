[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("xlflow-analyzer-performance-test-" + [Guid]::NewGuid().ToString("N"))
$testRoot = [System.IO.Path]::GetFullPath($testRoot)
$tempRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
if (-not $testRoot.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or [System.IO.Path]::GetFileName($testRoot) -notlike "xlflow-analyzer-performance-test-*") {
    throw "Refusing unsafe test path '$testRoot'."
}

try {
    [System.IO.Directory]::CreateDirectory($testRoot) | Out-Null
    $baseLog = Join-Path $testRoot "base.txt"
    $headLog = Join-Path $testRoot "head.txt"
    $baseLines = @(
        "BenchmarkRealWorldCorpus/ronecone/analyze-only/cold-20 1 100 ns/op 10 counter_object_summary_evaluations/op 5 counter_semantic_kernel_runs/op 1000 B/op 20 allocs/op",
        "BenchmarkRealWorldCorpus/ronecone/analyze-only/cold-20 1 120 ns/op 10 counter_object_summary_evaluations/op 6 counter_semantic_kernel_runs/op 1100 B/op 22 allocs/op"
    )
    $headLines = @(
        "BenchmarkRealWorldCorpus/ronecone/analyze-only/cold-20 1 105 ns/op 12 counter_object_summary_evaluations/op 5 counter_semantic_kernel_runs/op 1000 B/op 20 allocs/op",
        "BenchmarkRealWorldCorpus/ronecone/analyze-only/cold-20 1 115 ns/op 12 counter_object_summary_evaluations/op 5 counter_semantic_kernel_runs/op 1100 B/op 22 allocs/op"
    )
    [System.IO.File]::WriteAllLines($baseLog, $baseLines, [Text.UTF8Encoding]::new($false))
    [System.IO.File]::WriteAllLines($headLog, $headLines, [Text.UTF8Encoding]::new($false))

    $output = Join-Path $testRoot "result"
    & (Join-Path $PSScriptRoot "Compare-AnalyzerPerformance.ps1") -BaseLog $baseLog -HeadLog $headLog -OutputDirectory $output
    if ($LASTEXITCODE -ne 0) {
        throw "Comparison helper failed with exit code $LASTEXITCODE."
    }
    $result = Get-Content -LiteralPath (Join-Path $output "comparison.json") -Raw | ConvertFrom-Json
    $counter = $result.comparisons | Where-Object metric -eq "counter_object_summary_evaluations/op" | Select-Object -First 1
    if ($null -eq $counter -or $counter.base -ne 10 -or $counter.head -ne 12 -or -not $counter.suspicious) {
        throw "Deterministic counter regression was not classified correctly: $($counter | ConvertTo-Json -Compress)"
    }
    $time = $result.comparisons | Where-Object metric -eq "ns/op" | Select-Object -First 1
    if ($null -eq $time -or $time.base -ne 110 -or $time.head -ne 110 -or $time.suspicious) {
        throw "Median timing comparison was not calculated correctly: $($time | ConvertTo-Json -Compress)"
    }
    $unstable = $result.unstable_counters | Select-Object -First 1
    if ($result.metadata.count -ne 2 -or $result.unstable_counters.Count -ne 1 -or $unstable.side -ne "base" -or $unstable.metric -ne "counter_semantic_kernel_runs/op") {
        throw "Sample count or deterministic stability was not recorded correctly."
    }

    $gateOutput = Join-Path $testRoot "gate-result"
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "Compare-AnalyzerPerformance.ps1") -BaseLog $baseLog -HeadLog $headLog -OutputDirectory $gateOutput -FailOnRegression *> $null
    if ($LASTEXITCODE -ne 2) {
        throw "FailOnRegression exit code = $LASTEXITCODE, want 2."
    }
    Write-Output "Analyzer performance comparison self-test passed."
} finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}
