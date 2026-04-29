# Run Harness Enhancements Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Strengthen `xlflow run` so macros can be invoked with typed CLI arguments, measured execution time, structured VBA failure details, and explicit save or save-as behavior.

**Architecture:** Extend the Cobra `run` command to collect structured run options, pass them through the Go Excel bridge as explicit script arguments, and return a richer `macro` payload plus detailed `error` metadata in the existing envelope. Reuse PowerShell helper patterns already used by `test`, but generate a temporary VBA harness for each invocation so `xlflow run` can capture `Err.Number`, `Err.Description`, `Erl`, and elapsed time while defaulting to no workbook persistence.

**Tech Stack:** Go, Cobra, PowerShell, Excel COM/VBIDE, JSON envelope tests, `go test`, `task verify`

---

## Locked Decisions

- `--arg` is repeatable and requires explicit type prefixes: `string:`, `int:`, `bool:`.
- `--input` overrides `excel.path` for one invocation only.
- The default run does **not** save the workbook.
- `--save` and `--save-as` are mutually exclusive.
- `--save-as` writes a copy after a successful run and must keep the same workbook extension as the opened workbook.

## File Structure

| Path | Responsibility |
| --- | --- |
| `internal/cli/root.go` | Add `run` flags, parse typed CLI args, validate mutually exclusive save flags, and build `excel.RunOptions`. |
| `internal/cli/root_test.go` | Lock `run` command flag registration and Go-side option parsing behavior. |
| `internal/excel/bridge.go` | Introduce `RunArgument` and `RunOptions`, serialize them for `run.ps1`, and keep exit-code mapping stable. |
| `internal/excel/bridge_test.go` | Verify run option serialization, workbook override handling, and `macro_failed` exit-code classification. |
| `internal/output/output.go` | Extend `output.Error` with an optional `line` field without changing existing envelopes. |
| `internal/output/output_test.go` | Verify JSON envelopes preserve the new error line metadata. |
| `scripts/common.ps1` | Add typed run-argument conversion, VBA literal rendering, save-as extension validation, and temporary run-harness code generation that uses `Application.Run` for arbitrary macro names. |
| `scripts/run.ps1` | Open the selected workbook, fail cleanly on VBIDE access problems, inject and execute the temporary VBA harness, measure duration, and apply save or save-as behavior. |
| `scripts/scripts_test.go` | Lock helper parsing, VBA harness generation, and save-as validation behavior with `pwsh`-level tests. |
| `docs/specs/cli-contract.md` | Define the CLI, JSON, and exit-code contract for the stronger `run` behavior. |
| `README.md` | Update user-facing examples and describe the new `run` flags and failure metadata. |

### Task 1: Go-side run contract and CLI plumbing

**Files:**
- Modify: `internal/cli/root.go`
- Modify: `internal/cli/root_test.go`
- Modify: `internal/excel/bridge.go`
- Modify: `internal/excel/bridge_test.go`

- [ ] **Step 1: Write the failing Go tests for run flags and option serialization**

```go
package cli

import (
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/excel"
)

func TestRootCommandIncludesRunFlags(t *testing.T) {
	a := &app{}
	root := a.rootCommand()

	cmd, _, err := root.Find([]string{"run"})
	if err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"arg", "input", "save", "save-as"} {
		if cmd.Flags().Lookup(name) == nil {
			t.Fatalf("expected run command to define --%s", name)
		}
	}
}

func TestBuildRunOptionsRejectsConflictingSaveFlags(t *testing.T) {
	cfg := config.Default()
	_, err := buildRunOptions(cfg, "Main.Run", "", []string{"string:hello"}, true, "build\\result.xlsm")
	if err == nil || !strings.Contains(err.Error(), "--save and --save-as cannot be combined") {
		t.Fatalf("expected save conflict error, got %v", err)
	}
}

func TestBuildRunOptionsParsesTypedArguments(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptions(cfg, "", "fixtures\\Book.xlsm", []string{"string:hello", "int:7", "bool:true"}, false, "")
	if err != nil {
		t.Fatal(err)
	}

	want := []excel.RunArgument{
		{Type: "string", Value: "hello"},
		{Type: "int", Value: "7"},
		{Type: "bool", Value: "true"},
	}
	if opts.Macro != "Main.Run" {
		t.Fatalf("macro = %q, want Main.Run", opts.Macro)
	}
	if opts.WorkbookPath != "fixtures\\Book.xlsm" {
		t.Fatalf("workbook path = %q", opts.WorkbookPath)
	}
	if !reflect.DeepEqual(opts.Args, want) {
		t.Fatalf("run args = %#v, want %#v", opts.Args, want)
	}
}

func TestBuildRunOptionsAllowsEmptyStringArguments(t *testing.T) {
	cfg := config.Default()
	opts, err := buildRunOptions(cfg, "Main.Run", "", []string{"string:"}, false, "")
	if err != nil {
		t.Fatal(err)
	}
	if len(opts.Args) != 1 || opts.Args[0].Type != "string" || opts.Args[0].Value != "" {
		t.Fatalf("run args = %#v", opts.Args)
	}
}

func TestBuildRunOptionsRejectsMalformedTypedArguments(t *testing.T) {
	cfg := config.Default()
	for _, literal := range []string{"int:not-a-number", "bool:maybe"} {
		_, err := buildRunOptions(cfg, "Main.Run", "", []string{literal}, false, "")
		if err == nil {
			t.Fatalf("expected %q to fail", literal)
		}
	}
}
```

```go
package excel

import (
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/output"
)

func TestBuildRunScriptArgsSerializesArgumentsAndSaveAs(t *testing.T) {
	root := t.TempDir()
	cfg := config.Default()
	args, err := buildRunScriptArgs(root, cfg, RunOptions{
		Macro:        "Report.Generate",
		WorkbookPath: "fixtures\\Book.xlsm",
		Args: []RunArgument{
			{Type: "string", Value: "fixtures\\sample.xlsx"},
			{Type: "int", Value: "3"},
			{Type: "bool", Value: "true"},
		},
		SaveAs: "build\\Result.xlsm",
	})
	if err != nil {
		t.Fatal(err)
	}
	if args["MacroName"] != "Report.Generate" {
		t.Fatalf("macro name = %q", args["MacroName"])
	}
	if args["WorkbookPath"] != filepath.Join(root, "fixtures", "Book.xlsm") {
		t.Fatalf("workbook path = %q", args["WorkbookPath"])
	}
	if args["SaveWorkbook"] != "false" {
		t.Fatalf("save flag = %q", args["SaveWorkbook"])
	}
	if args["SaveAsPath"] != filepath.Join(root, "build", "Result.xlsm") {
		t.Fatalf("save-as path = %q", args["SaveAsPath"])
	}
	wantJSON := `[{"type":"string","value":"fixtures\\sample.xlsx"},{"type":"int","value":"3"},{"type":"bool","value":"true"}]`
	if args["MacroArgsJson"] != wantJSON {
		t.Fatalf("macro args json = %s", args["MacroArgsJson"])
	}
}

func TestMacroFailureIsValidationFailure(t *testing.T) {
	result := ScriptResult{
		Status: output.StatusFailed,
		Error:  &output.Error{Code: "macro_failed", Message: "boom"},
	}
	if got := exitCodeForScriptResult(result); got != output.ExitValidation {
		t.Fatalf("exitCodeForScriptResult(macro_failed) = %d, want %d", got, output.ExitValidation)
	}
}
```

- [ ] **Step 2: Run the focused Go tests to verify the contract is missing**

```bash
go test ./internal/cli ./internal/excel -run 'TestRootCommandIncludesRunFlags|TestBuildRunOptionsRejectsConflictingSaveFlags|TestBuildRunOptionsParsesTypedArguments|TestBuildRunOptionsAllowsEmptyStringArguments|TestBuildRunOptionsRejectsMalformedTypedArguments|TestBuildRunScriptArgsSerializesArgumentsAndSaveAs|TestMacroFailureIsValidationFailure'
```

Expected: FAIL with missing run flags, undefined `buildRunOptions`, undefined `RunOptions`, and undefined `buildRunScriptArgs`.

- [ ] **Step 3: Implement the minimal Go-side run option model and flag plumbing**

```go
package excel

import (
	"bytes"
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strconv"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/output"
)

type RunArgument struct {
	Type  string `json:"type"`
	Value string `json:"value"`
}

type RunOptions struct {
	Macro        string
	WorkbookPath string
	Args         []RunArgument
	Save         bool
	SaveAs       string
}

func buildRunScriptArgs(root string, cfg config.Config, opts RunOptions) (map[string]string, error) {
	if opts.Macro == "" {
		opts.Macro = cfg.Project.Entry
	}
	workbook := cfg.Excel.Path
	if opts.WorkbookPath != "" {
		workbook = opts.WorkbookPath
	}
	argsJSON, err := json.Marshal(opts.Args)
	if err != nil {
		return nil, err
	}
	scriptArgs := map[string]string{
		"WorkbookPath":  workbookPath(root, workbook),
		"MacroName":     opts.Macro,
		"MacroArgsJson": string(argsJSON),
		"Visible":       strconv.FormatBool(cfg.Excel.Visible),
		"DisplayAlerts": strconv.FormatBool(cfg.Excel.DisplayAlerts),
		"SaveWorkbook":  strconv.FormatBool(opts.Save),
	}
	if opts.SaveAs != "" {
		scriptArgs["SaveAsPath"] = workbookPath(root, opts.SaveAs)
	}
	return scriptArgs, nil
}

func (r Runner) Run(cfg config.Config, opts RunOptions) (output.Envelope, int, error) {
	scriptArgs, err := buildRunScriptArgs(r.RootDir, cfg, opts)
	if err != nil {
		return output.Failure("run", output.Error{Code: "run_args_invalid", Message: err.Error(), Source: "xlflow"}), output.ExitConfig, nil
	}
	return r.run("run", scriptArgs)
}
```

```go
package cli

import (
	"errors"
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/spf13/cobra"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/excel"
	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/output"
	"github.com/harumiWeb/xlflow/internal/project"
)

func buildRunOptions(cfg config.Config, macro, input string, argLiterals []string, save bool, saveAs string) (excel.RunOptions, error) {
	if save && saveAs != "" {
		return excel.RunOptions{}, fmt.Errorf("--save and --save-as cannot be combined")
	}
	if macro == "" {
		macro = cfg.Project.Entry
	}
	args := make([]excel.RunArgument, 0, len(argLiterals))
	for _, literal := range argLiterals {
		parts := strings.SplitN(literal, ":", 2)
		if len(parts) != 2 {
			return excel.RunOptions{}, fmt.Errorf("invalid --arg %q: expected type:value", literal)
		}
		switch parts[0] {
		case "string":
		case "int", "bool":
			if parts[1] == "" {
				return excel.RunOptions{}, fmt.Errorf("invalid --arg %q: %s values cannot be empty", literal, parts[0])
			}
			if parts[0] == "int" {
				if _, err := strconv.Atoi(parts[1]); err != nil {
					return excel.RunOptions{}, fmt.Errorf("invalid --arg %q: int values must parse as base-10 integers", literal)
				}
			}
			if parts[0] == "bool" && parts[1] != "true" && parts[1] != "false" {
				return excel.RunOptions{}, fmt.Errorf("invalid --arg %q: bool values must be true or false", literal)
			}
		default:
			return excel.RunOptions{}, fmt.Errorf("unsupported --arg type prefix %q", parts[0])
		}
		args = append(args, excel.RunArgument{Type: parts[0], Value: parts[1]})
	}
	return excel.RunOptions{
		Macro:        macro,
		WorkbookPath: input,
		Args:         args,
		Save:         save,
		SaveAs:       saveAs,
	}, nil
}

func (a *app) runCommand() *cobra.Command {
	var argLiterals []string
	var input string
	var save bool
	var saveAs string

	cmd := &cobra.Command{
		Use:   "run [macro]",
		Short: "Run a workbook macro",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			cfg, err := a.loadConfig("run")
			if err != nil {
				return err
			}
			macro := ""
			if len(args) == 1 {
				macro = args[0]
			}
			opts, err := buildRunOptions(cfg, macro, input, argLiterals, save, saveAs)
			if err != nil {
				return a.writeFailure("run", output.ExitConfig, "run_args_invalid", err)
			}
			env, code, err := excel.Runner{RootDir: a.cwd}.Run(cfg, opts)
			if err != nil {
				return err
			}
			return a.write(env, code)
		},
	}
	cmd.Flags().StringArrayVar(&argLiterals, "arg", nil, "pass a typed macro argument such as string:hello, int:7, or bool:true")
	cmd.Flags().StringVar(&input, "input", "", "override workbook path for this run")
	cmd.Flags().BoolVar(&save, "save", false, "save the opened workbook after a successful run")
	cmd.Flags().StringVar(&saveAs, "save-as", "", "write the successful workbook result to a new path")
	return cmd
}
```

- [ ] **Step 4: Run the focused Go tests to verify the Go-side contract now passes**

```bash
go test ./internal/cli ./internal/excel -run 'TestRootCommandIncludesRunFlags|TestBuildRunOptionsRejectsConflictingSaveFlags|TestBuildRunOptionsParsesTypedArguments|TestBuildRunOptionsAllowsEmptyStringArguments|TestBuildRunOptionsRejectsMalformedTypedArguments|TestBuildRunScriptArgsSerializesArgumentsAndSaveAs|TestMacroFailureIsValidationFailure'
```

Expected: PASS.

- [ ] **Step 5: Commit the Go-side contract work**

```bash
git add internal/cli/root.go internal/cli/root_test.go internal/excel/bridge.go internal/excel/bridge_test.go
git commit -m "feat: add structured run options"
```

### Task 2: Structured run result and PowerShell harness

**Files:**
- Modify: `internal/output/output.go`
- Modify: `internal/output/output_test.go`
- Modify: `scripts/common.ps1`
- Modify: `scripts/run.ps1`
- Modify: `scripts/scripts_test.go`

- [ ] **Step 1: Write the failing tests for error line metadata and run harness helpers**

```go
package output

import (
	"bytes"
	"encoding/json"
	"testing"
)

func TestWriteJSONEnvelopeIncludesErrorLine(t *testing.T) {
	env := Failure("run", Error{
		Code:    "macro_failed",
		Message: "inputPath is required",
		Source:  "Main",
		Number:  5,
		Line:    10,
	})
	var buf bytes.Buffer
	if err := Write(&buf, env, true); err != nil {
		t.Fatal(err)
	}
	var decoded map[string]any
	if err := json.Unmarshal(buf.Bytes(), &decoded); err != nil {
		t.Fatal(err)
	}
	errorMap, ok := decoded["error"].(map[string]any)
	if !ok {
		t.Fatalf("error payload = %#v", decoded["error"])
	}
	if errorMap["line"] != float64(10) {
		t.Fatalf("error line = %#v", errorMap["line"])
	}
}
```

```go
package scripts_test

import (
	"os/exec"
	"strings"
	"testing"
)

func TestRunArgumentConversionSupportsExplicitTypes(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $json = '[{\"type\":\"string\",\"value\":\"hello\"},{\"type\":\"int\",\"value\":\"7\"},{\"type\":\"bool\",\"value\":\"true\"}]'; $values = ConvertFrom-XlflowRunArgumentsJson -Json $json; ConvertTo-Json -InputObject $values -Compress",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("run argument conversion failed: %v\n%s", err, out)
	}
	got := strings.TrimSpace(string(out))
	if got != "[\"hello\",7,true]" {
		t.Fatalf("converted values = %s", got)
	}
}

func TestRunHarnessCodeIncludesApplicationRunAndErrorLine(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; $args = @([pscustomobject]@{ type = 'string'; value = 'fixtures\\sample.xlsx' }, [pscustomobject]@{ type = 'int'; value = '3' }, [pscustomobject]@{ type = 'bool'; value = 'true' }); New-XlflowRunHarnessCode -MacroName 'Report.Generate' -Arguments $args",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("run harness code generation failed: %v\n%s", err, out)
	}
	got := string(out)
	for _, want := range []string{"Application.Run \"Report.Generate\"", "\"fixtures\\sample.xlsx\"", "CLng(3)", "CBool(True)", "Err.Description", "Erl"} {
		if !strings.Contains(got, want) {
			t.Fatalf("expected run harness code to contain %q:\n%s", want, got)
		}
	}
}

func TestSaveAsExtensionValidationRejectsMismatchedTargets(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; try { Assert-XlflowSaveAsExtension -WorkbookPath 'build\\Book.xlsm' -SaveAsPath 'build\\Book.xlsx'; 'unexpected success' } catch { $_.Exception.Message }",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("save-as validation command failed: %v\n%s", err, out)
	}
	got := strings.TrimSpace(string(out))
	if !strings.Contains(got, "does not match workbook extension") {
		t.Fatalf("validation output = %q", got)
	}
}

func TestFormatMacroFailureMessageIncludesLineAndErrNumber(t *testing.T) {
	cmd := exec.Command(
		"pwsh",
		"-NoProfile",
		"-Command",
		". ./common.ps1; Format-XlflowMacroFailureMessage -ModuleName 'Main' -Line 10 -Number 5 -Description 'inputPath is required'",
	)
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("macro failure message formatting failed: %v\n%s", err, out)
	}
	got := strings.TrimSpace(string(out))
	if got != "Main line 10 Err 5: inputPath is required" {
		t.Fatalf("failure message = %q", got)
	}
}
```

- [ ] **Step 2: Run the focused tests to verify the richer run harness is not implemented yet**

```bash
go test ./internal/output ./scripts -run 'TestWriteJSONEnvelopeIncludesErrorLine|TestRunArgumentConversionSupportsExplicitTypes|TestRunHarnessCodeIncludesApplicationRunAndErrorLine|TestSaveAsExtensionValidationRejectsMismatchedTargets|TestFormatMacroFailureMessageIncludesLineAndErrNumber'
```

Expected: FAIL because `output.Error` has no `Line` field and the new PowerShell helpers do not exist.

- [ ] **Step 3: Implement error line support, typed PowerShell helpers, and the run harness**

```go
package output

type Error struct {
	Code    string `json:"code,omitempty"`
	Message string `json:"message"`
	Source  string `json:"source,omitempty"`
	Number  int    `json:"number,omitempty"`
	Line    int    `json:"line,omitempty"`
}
```

```powershell
function Set-XlflowError {
  param(
    $Result,
    [string]$Code,
    [string]$Message,
    [string]$Source = "",
    [int]$Number = 0,
    [int]$Line = 0
  )
  $Result.status = "failed"
  $Result.error = [ordered]@{
    code = $Code
    message = $Message
    source = $Source
    number = $Number
    line = $Line
  }
}

function ConvertFrom-XlflowRunArgumentsJson {
  param([string]$Json)

  if ([string]::IsNullOrWhiteSpace($Json)) {
    return @()
  }
  $specs = @(ConvertFrom-Json -InputObject $Json)
  $values = New-Object System.Collections.Generic.List[object]
  foreach ($spec in $specs) {
    switch ([string]$spec.type) {
      "string" {
        $values.Add([string]$spec.value)
      }
      "int" {
        $parsed = 0
        if (-not [int]::TryParse([string]$spec.value, [ref]$parsed)) {
          throw "invalid int run argument: $($spec.value)"
        }
        $values.Add($parsed)
      }
      "bool" {
        if ($spec.value -ne "true" -and $spec.value -ne "false") {
          throw "invalid bool run argument: $($spec.value)"
        }
        $values.Add((ConvertTo-XlflowBool ([string]$spec.value)))
      }
      default {
        throw "unsupported run argument type: $($spec.type)"
      }
    }
  }
  return ,$values.ToArray()
}

function ConvertTo-XlflowVBALiteral {
  param([string]$Type, [string]$Value)

  switch ($Type) {
    "string" { return """" + $Value.Replace("""", """""") + """" }
    "int" { return "CLng(" + $Value + ")" }
    "bool" {
      if ($Value -eq "true") {
        return "CBool(True)"
      }
      return "CBool(False)"
    }
    default { throw "unsupported run argument type: $Type" }
  }
}

function Get-XlflowMacroModuleName {
  param([string]$MacroName)

  $parts = $MacroName.Split(".")
  if ($parts.Count -lt 2) {
    return $MacroName
  }
  return ($parts[0..($parts.Count - 2)] -join ".")
}

function Assert-XlflowSaveAsExtension {
  param([string]$WorkbookPath, [string]$SaveAsPath)

  if ([string]::IsNullOrWhiteSpace($SaveAsPath)) {
    return
  }
  $workbookExtension = [System.IO.Path]::GetExtension($WorkbookPath)
  $saveAsExtension = [System.IO.Path]::GetExtension($SaveAsPath)
  if ($workbookExtension -ne $saveAsExtension) {
    throw "save-as extension $saveAsExtension does not match workbook extension $workbookExtension"
  }
}

function Format-XlflowMacroFailureMessage {
  param(
    [string]$ModuleName,
    [int]$Line,
    [int]$Number,
    [string]$Description
  )

  $parts = New-Object System.Collections.Generic.List[string]
  if (-not [string]::IsNullOrWhiteSpace($ModuleName)) {
    $parts.Add($ModuleName)
  }
  if ($Line -gt 0) {
    $parts.Add("line " + $Line)
  }
  if ($Number -ne 0) {
    $parts.Add("Err " + $Number)
  }
  if ([string]::IsNullOrWhiteSpace($Description)) {
    return ($parts -join " ")
  }
  return (($parts -join " ") + ": " + $Description).Trim()
}

function New-XlflowRunHarnessCode {
  param([string]$MacroName, $Arguments)

  $builder = New-Object System.Text.StringBuilder
  $moduleName = Get-XlflowMacroModuleName -MacroName $MacroName
  $literals = New-Object System.Collections.Generic.List[string]
  foreach ($argument in $Arguments) {
    $literals.Add((ConvertTo-XlflowVBALiteral -Type ([string]$argument.type) -Value ([string]$argument.value)))
  }
  $escapedMacroName = $MacroName.Replace("""", """""")
  $invocation = "Application.Run """ + $escapedMacroName + """"
  if ($literals.Count -gt 0) {
    $invocation += ", " + ($literals -join ", ")
  }

  [void]$builder.AppendLine("Option Explicit")
  [void]$builder.AppendLine("")
  [void]$builder.AppendLine("Public Function RunMacro() As Variant")
  [void]$builder.AppendLine("  Dim startedAt As Double")
  [void]$builder.AppendLine("  startedAt = Timer")
  [void]$builder.AppendLine("  On Error GoTo Handler")
  [void]$builder.AppendLine("  " + $invocation)
  [void]$builder.AppendLine("  RunMacro = Array(True, """ + $moduleName + """, CLng(0), """", CLng(0), CLng((Timer - startedAt) * 1000))")
  [void]$builder.AppendLine("  Exit Function")
  [void]$builder.AppendLine("Handler:")
  [void]$builder.AppendLine("  RunMacro = Array(False, """ + $moduleName + """, CLng(Err.Number), CStr(Err.Description), CLng(Erl), CLng((Timer - startedAt) * 1000))")
  [void]$builder.AppendLine("  Err.Clear")
  [void]$builder.AppendLine("End Function")
  return $builder.ToString()
}
```

```powershell
param(
  [string]$WorkbookPath,
  [string]$MacroName,
  [string]$MacroArgsJson = "[]",
  [string]$Visible = "false",
  [string]$DisplayAlerts = "false",
  [string]$SaveWorkbook = "false",
  [string]$SaveAsPath = ""
)

. "$PSScriptRoot/common.ps1"

$result = New-XlflowResult -Command "run"
$excel = $null
$workbook = $null
$vbProject = $null
$runnerComponent = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = ConvertTo-XlflowBool $Visible
  $excel.DisplayAlerts = ConvertTo-XlflowBool $DisplayAlerts
  $workbook = $excel.Workbooks.Open($WorkbookPath)

  $argumentSpecs = @()
  if (-not [string]::IsNullOrWhiteSpace($MacroArgsJson)) {
    $argumentSpecs = @(ConvertFrom-Json -InputObject $MacroArgsJson)
  }
  $typedValues = @(ConvertFrom-XlflowRunArgumentsJson -Json $MacroArgsJson)

  try {
    $vbProject = $workbook.VBProject
    $runnerComponent = $vbProject.VBComponents.Add(1)
  } catch {
    Set-XlflowError -Result $result -Code "vbide_access_denied" -Message $_.Exception.Message -Source "vbide"
    throw
  }

  $runnerName = "XlflowRunHarness_" + [Guid]::NewGuid().ToString("N")
  $runnerComponent.Name = $runnerName
  $runnerComponent.CodeModule.AddFromString((New-XlflowRunHarnessCode -MacroName $MacroName -Arguments $argumentSpecs))

  $runResult = $excel.Run($runnerName + ".RunMacro")
  $successLog = "ran " + $MacroName + " in " + ([int]$runResult[5]) + "ms"
  if ($null -ne $runnerComponent) {
    $vbProject.VBComponents.Remove($runnerComponent)
    $runnerComponent = $null
  }
  $result.macro = [ordered]@{
    name = $MacroName
    args = @($typedValues)
    duration_ms = [int]$runResult[5]
  }

  if (-not [bool]$runResult[0]) {
    $failureMessage = Format-XlflowMacroFailureMessage -ModuleName ([string]$runResult[1]) -Line ([int]$runResult[4]) -Number ([int]$runResult[2]) -Description ([string]$runResult[3])
    Set-XlflowError -Result $result -Code "macro_failed" -Message $failureMessage -Source ([string]$runResult[1]) -Number ([int]$runResult[2]) -Line ([int]$runResult[4])
  } elseif (ConvertTo-XlflowBool $SaveWorkbook) {
    $workbook.Save()
    $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $true; save_as = $null }
    $result.logs = @($successLog, "saved workbook in place")
  } elseif (-not [string]::IsNullOrWhiteSpace($SaveAsPath)) {
    Assert-XlflowSaveAsExtension -WorkbookPath $WorkbookPath -SaveAsPath $SaveAsPath
    $targetDir = Split-Path -Parent $SaveAsPath
    if (-not [string]::IsNullOrWhiteSpace($targetDir)) {
      New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
    }
    $workbook.SaveCopyAs($SaveAsPath)
    $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $false; save_as = $SaveAsPath }
    $result.logs = @($successLog, "wrote workbook copy to " + $SaveAsPath)
  } else {
    $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $false; save_as = $null }
    $result.logs = @($successLog, "left workbook unchanged on disk")
  }
} catch {
  if ($result.error -eq $null) {
    Set-XlflowError -Result $result -Code "macro_failed" -Message $_.Exception.Message -Source $_.Exception.Source -Number $_.Exception.HResult
  }
  $result.macro = [ordered]@{ name = $MacroName; args = @(); duration_ms = 0 }
  $result.workbook = [ordered]@{ path = $WorkbookPath; saved = $false; save_as = $null }
} finally {
  if ($null -ne $runnerComponent -and $null -ne $vbProject) {
    try { $vbProject.VBComponents.Remove($runnerComponent) } catch {}
  }
  Close-XlflowCom -Workbook $workbook -Excel $excel -Save $false
}

Write-XlflowJson -Result $result
```

- [ ] **Step 4: Run the focused tests to verify the richer run harness now passes**

```bash
go test ./internal/output ./scripts -run 'TestWriteJSONEnvelopeIncludesErrorLine|TestRunArgumentConversionSupportsExplicitTypes|TestRunHarnessCodeIncludesApplicationRunAndErrorLine|TestSaveAsExtensionValidationRejectsMismatchedTargets|TestFormatMacroFailureMessageIncludesLineAndErrNumber'
```

Expected: PASS.

- [ ] **Step 5: Commit the run harness implementation**

```bash
git add internal/output/output.go internal/output/output_test.go scripts/common.ps1 scripts/run.ps1 scripts/scripts_test.go
git commit -m "feat: enrich run harness results"
```

### Task 3: Documentation and full verification

**Files:**
- Modify: `docs/specs/cli-contract.md`
- Modify: `README.md`

- [ ] **Step 1: Update the CLI contract for the stronger run command**

```md
## Commands

```text
xlflow [--json] run [macro] [--input <workbook>] [--arg <type:value>]... [--save | --save-as <path>]
```

`run` uses the positional macro argument when provided. Otherwise it uses `project.entry` from `xlflow.toml`. `--input` overrides `excel.path` for one invocation. `--arg` may be repeated and must use explicit prefixes: `string:hello`, `string:`, `int:7`, and `bool:true`. Empty values are valid only for `string:` arguments. Malformed `int:` and `bool:` values are rejected by the CLI before Excel starts and exit with code `2`. The default run never saves. `--save` persists the opened workbook in place after a successful run. `--save-as` writes a copy after a successful run and must keep the same workbook extension as the opened workbook. `--save` and `--save-as` cannot be combined.

`run` adds a `macro` object with `name`, `args`, and `duration_ms`. Failed macro runs return `macro_failed` with `error.source`, `error.number`, `error.line`, and `error.message` populated from the VBA failure. Plain-text success output must include the elapsed duration and whether the workbook was saved, copied, or left unchanged. Plain-text failure output must use the formatted message `Module line <n> Err <n>: <description>` when line and error number are available. Because `run` injects a temporary VBA harness to measure duration and collect `Erl`, VBIDE access failures return an environment error such as `vbide_access_denied` and exit code `3`.
```

- [ ] **Step 2: Update the README examples and user guidance**

```md
## MVP Commands

```bash
xlflow run Main.Run --json
xlflow run Report.Generate --arg string:fixtures\sample.xlsx --arg int:3 --arg bool:true --save --json
xlflow run Report.Generate --input build\Book.xlsm --arg string:fixtures\sample.xlsx --save-as build\Result.xlsm --json
```

`xlflow run` accepts repeatable typed arguments with the `string:`, `int:`, and `bool:` prefixes. `string:` may carry an empty value. Successful runs report `macro.duration_ms` in JSON and print the elapsed time plus save behavior in plain text. Failing runs return `macro_failed` with VBA error metadata including module name, line number, `Err.Number`, and `Err.Description`, and the non-JSON message must stay readable as `Main line 10 Err 5: inputPath is required`. The default run does not save the workbook; use `--save` or `--save-as` explicitly.
```

- [ ] **Step 3: Run the repository verification commands**

```bash
go test ./...
task verify
```

Expected: PASS.

- [ ] **Step 4: Run a real Excel COM check in a disposable workspace**

```vb
Attribute VB_Name = "Main"
Option Explicit

Public Sub GenerateReport(ByVal inputPath As String, ByVal copies As Long, ByVal overwrite As Boolean)
10  If Len(inputPath) = 0 Then Err.Raise 5, "Main", "inputPath is required"
20  ThisWorkbook.Worksheets("Sheet1").Range("A1").Value = inputPath
30  ThisWorkbook.Worksheets("Sheet1").Range("A2").Value = copies
40  ThisWorkbook.Worksheets("Sheet1").Range("A3").Value = overwrite
End Sub
```

```bash
xlflow new RunHarness.xlsm
@'
Attribute VB_Name = "Main"
Option Explicit

Public Sub GenerateReport(ByVal inputPath As String, ByVal copies As Long, ByVal overwrite As Boolean)
10  If Len(inputPath) = 0 Then Err.Raise 5, "Main", "inputPath is required"
20  ThisWorkbook.Worksheets("Sheet1").Range("A1").Value = inputPath
30  ThisWorkbook.Worksheets("Sheet1").Range("A2").Value = copies
40  ThisWorkbook.Worksheets("Sheet1").Range("A3").Value = overwrite
End Sub
'@ | Set-Content -LiteralPath src\modules\Main.bas
xlflow push --json
xlflow run Main.GenerateReport --arg string:fixtures\sample.xlsx --arg int:3 --arg bool:true
xlflow run Main.GenerateReport --arg string:fixtures\sample.xlsx --arg int:3 --arg bool:true --save --json
xlflow run Main.GenerateReport --arg string:fixtures\sample.xlsx --arg int:3 --arg bool:true --save-as build\RunHarness-copy.xlsm --json
xlflow run Main.GenerateReport --arg string: --arg int:3 --arg bool:true
```

This writes the VBA snippet above into `src\modules\Main.bas` before running `xlflow push --json`.

Expected:
- First run prints `ran Main.GenerateReport in <n>ms` and `left workbook unchanged on disk`.
- Second run returns `status: ok`, `macro.duration_ms > 0`, and `workbook.saved = true`.
- Third run returns `status: ok`, `workbook.save_as = "build\\RunHarness-copy.xlsm"`, and leaves the opened workbook path unchanged.
- Fourth run exits with code `1` and prints `Main line 10 Err 5: inputPath is required`.

Use the repo-local `xlflow-tmp-workspace-e2e` skill for this disposable workbook flow so the main tree stays untouched.

- [ ] **Step 5: Commit the docs and verification updates**

```bash
git add docs/specs/cli-contract.md README.md
git commit -m "docs: describe enhanced run harness"
```

## Coverage Check

- Spec item: run a named macro from the CLI -> Task 1 and Task 2.
- Spec item: support explicit macro arguments -> Task 1 and Task 2.
- Spec item: report execution duration -> Task 2 and Task 3.
- Spec item: show module name, line number, `Err.Number`, and `Err.Description` -> Task 2 and Task 3.
- Spec item: choose whether to save or save-as -> Task 1, Task 2, and Task 3.

No open gaps remain against `docs/design.md` section 2.
