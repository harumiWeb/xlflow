param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Load only pure functions: never dot-source the Excel COM harness.
$sourcePath = Join-Path $PSScriptRoot 'test-userform-generation-e2e.ps1'
$tokens = $null
$parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($sourcePath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }
$functionNames = @('Get-OptionalProperty', 'Convert-ObservedValue', 'Add-ObservedField', 'Get-ControlType', 'Get-ControlSnapshot', 'Assert-Value', 'Get-PropertyOrNull', 'Get-ExpectedFormMap', 'Assert-ExpectedFields', 'Assert-ExpectedForm')
foreach ($name in $functionNames) {
    $definitions = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $false))
    if ($definitions.Count -ne 1) { throw "Expected one definition of $name" }
    . ([scriptblock]::Create($definitions[0].Extent.Text))
}
$script:generatedFormNames = @('GeneratedForm', 'GeneratedEmptyForm')
$script:controlSpecs = @()
$script:checks = 0

function Assert-Test([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
    $script:checks++
}

function Assert-Rejected([scriptblock]$Action, [string]$MessagePattern) {
    $caught = $null
    try { & $Action | Out-Null } catch { $caught = $_ }
    Assert-Test ($null -ne $caught) "Expected rejection matching $MessagePattern"
    Assert-Test ($caught.Exception.Message -match $MessagePattern) "Unexpected rejection: $($caught.Exception.Message)"
}

$completeJson = '[{"form":{"name":"GeneratedForm","build":{"clientWidth":240,"clientHeight":180}},"controls":[]},{"form":{"name":"GeneratedEmptyForm","build":{"clientWidth":240,"clientHeight":180}},"controls":[]}]'
$complete = ConvertFrom-Json $completeJson
foreach ($inputValue in @(
    ,$complete
    [pscustomobject]@{ forms = $complete }
    [pscustomobject]@{ forms = [pscustomobject]@{ GeneratedForm = $complete[0]; GeneratedEmptyForm = $complete[1] } }
    [pscustomobject]@{ GeneratedForm = $complete[0]; GeneratedEmptyForm = $complete[1] }
)) {
    $map = Get-ExpectedFormMap $inputValue
    Assert-Test ($map.Count -eq 2 -and $map.Contains('GeneratedForm') -and $map.Contains('GeneratedEmptyForm')) 'Complete expected map was not retained'
}
Assert-Rejected { Get-ExpectedFormMap @() } 'must provide GeneratedForm'
Assert-Rejected { Get-ExpectedFormMap @($complete[0]) } 'must provide GeneratedEmptyForm'
Assert-Rejected { Get-ExpectedFormMap @($complete[1]) } 'must provide GeneratedForm'
Assert-Rejected { Get-ExpectedFormMap @($complete[0], $complete[0], $complete[1]) } 'duplicate form GeneratedForm'
Assert-Rejected { Get-ExpectedFormMap ([pscustomobject]@{ forms = @($complete[0], $complete[0], $complete[1]) }) } 'duplicate form GeneratedForm'
Assert-Rejected { Get-ExpectedFormMap @([pscustomobject]@{}, $complete[1]) } 'form has no name'

# Public CLR properties exercise the same reflection lookup as COM snapshots.
Add-Type -TypeDefinition @'
public class GenerationContractSpin {
    public GenerationContractSpin() { Name = "SpinButtonMain"; Delay = 50; }
    public string Name { get; set; }
    public string ProgID { get { return "Forms.SpinButton.1"; } }
    public double Left { get { return 0; } }
    public double Top { get { return 0; } }
    public double Width { get { return 12; } }
    public double Height { get { return 24; } }
    public bool Enabled { get { return true; } }
    public bool Visible { get { return true; } }
    public int Delay { get; set; }
}
public class GenerationContractScroll : GenerationContractSpin {
    public GenerationContractScroll() { ProportionalThumb = true; }
    public new string ProgID { get { return "Forms.ScrollBar.1"; } }
    public bool ProportionalThumb { get; set; }
}
'@
$spin = Get-ControlSnapshot (New-Object GenerationContractSpin)
$scrollControl = New-Object GenerationContractScroll
$scrollControl.Name = 'ScrollBarMain'
$scroll = Get-ControlSnapshot $scrollControl
Assert-Test ($spin.properties.Delay -eq 50 -and $spin.delay -eq 50) 'SpinButton Delay must be retained in both snapshot shapes'
Assert-Test ($null -eq $spin.PSObject.Properties['proportionalThumb'] -and $null -eq $spin.properties.PSObject.Properties['ProportionalThumb']) 'SpinButton must not acquire ProportionalThumb'
Assert-Test ($scroll.properties.Delay -eq 50 -and $scroll.properties.ProportionalThumb -and $scroll.proportionalThumb) 'ScrollBar Delay/ProportionalThumb must be retained'

$snapshot = [pscustomobject]@{
    name = 'GeneratedForm'
    form = [pscustomobject]@{}
    root = [pscustomobject]@{ clientWidth = 240; clientHeight = 180 }
    controls = @($spin, $scroll)
}
$expectedJson = '{"form":{"name":"GeneratedForm","build":{"clientWidth":240,"clientHeight":180}},"controls":[{"name":"SpinButtonMain","type":"SpinButton","observed":{"properties":{"Delay":50}}},{"name":"ScrollBarMain","type":"ScrollBar","observed":{"properties":{"Delay":50,"ProportionalThumb":true}}}]}'
$expected = ConvertFrom-Json $expectedJson
Assert-ExpectedForm $snapshot $expected
Assert-Test $true 'Canonical properties expectations pass'
foreach ($case in @(
    [pscustomobject]@{ index = 0; key = 'Delay'; wrong = 51 }
    [pscustomobject]@{ index = 1; key = 'Delay'; wrong = 51 }
    [pscustomobject]@{ index = 1; key = 'ProportionalThumb'; wrong = $false }
)) {
    $wrong = ConvertFrom-Json $expectedJson
    $wrong.controls[$case.index].observed.properties.($case.key) = $case.wrong
    Assert-Rejected { Assert-ExpectedForm $snapshot $wrong } ([regex]::Escape("observed.properties.$($case.key)"))
    $missing = ConvertFrom-Json ($snapshot | ConvertTo-Json -Depth 20)
    $missing.controls[$case.index].properties.PSObject.Properties.Remove($case.key)
    Assert-Rejected { Assert-ExpectedForm $missing $expected } "missing expected property $($case.key)"
}
foreach ($shape in @('control', 'observed')) {
    foreach ($key in @('delay', 'proportionalThumb')) {
        $direct = ConvertFrom-Json $completeJson
        $control = [pscustomobject]@{ name = 'ScrollBarMain'; observed = [pscustomobject]@{} }
        $target = if ($shape -eq 'control') { $control } else { $control.observed }
        $value = if ($key -eq 'delay') { 50 } else { $true }
        $target | Add-Member -NotePropertyName $key -NotePropertyValue $value
        $direct[0].controls = @($control)
        Assert-ExpectedForm $snapshot $direct[0]
        $target.($key) = if ($key -eq 'delay') { 51 } else { $false }
        Assert-Rejected { Assert-ExpectedForm $snapshot $direct[0] } ([regex]::Escape(".$key"))
        $target.($key) = $value
        $missing = ConvertFrom-Json ($snapshot | ConvertTo-Json -Depth 20)
        $missing.controls[1].PSObject.Properties.Remove($key)
        Assert-Rejected { Assert-ExpectedForm $missing $direct[0] } "missing expected property $key"
    }
}

# Preflight stays outside the COM try/finally and before its first instantiation.
$preflightCalls = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Get-ExpectedFormMap' }, $true))
$excelStarts = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.CommandAst] -and $node.Extent.Text -match '^New-Object -ComObject Excel.Application$' }, $true))
Assert-Test ($preflightCalls.Count -eq 1 -and $excelStarts.Count -eq 1 -and $preflightCalls[0].Extent.StartOffset -lt $excelStarts[0].Extent.StartOffset) 'Expected map must be validated before Excel starts'

$collisionGuards = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.IfStatementAst] -and $node.Extent.Text -match '^if \(Test-Path -LiteralPath \$observationOutput\)' }, $true))
Assert-Test ($collisionGuards.Count -eq 1 -and $collisionGuards[0].Extent.StartOffset -lt $excelStarts[0].Extent.StartOffset) 'Observation collision must be rejected before Excel starts'
$collisionGuard = [scriptblock]::Create($collisionGuards[0].Extent.Text)
$observationOutput = [IO.Path]::GetTempFileName()
try {
    Assert-Rejected { & $collisionGuard } 'Refusing to overwrite observation evidence'
} finally {
    Remove-Item -LiteralPath $observationOutput
}
& $collisionGuard
Assert-Test $true 'Fresh observation path accepted'

$observationWrites = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Write-Json' -and $node.Extent.Text -match '^Write-Json \$observationOutput ' }, $true))
$failureGuards = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.IfStatementAst] -and $node.Extent.Text -eq 'if ($null -ne $failure) { throw $failure }' }, $true))
Assert-Test ($observationWrites.Count -eq 1 -and $failureGuards.Count -eq 1 -and $observationWrites[0].Extent.StartOffset -gt $failureGuards[0].Extent.EndOffset) 'Observations must only publish after confirmed successful cleanup'
Write-Output "UserForm generation contract: $($script:checks) checks passed (no Excel COM)."
