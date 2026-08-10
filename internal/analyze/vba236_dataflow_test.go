package analyze

import (
	"encoding/json"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestVBA236OwnsProcessLaunchFlowsAndAddsCommandContext(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    Dim raw As String
    raw = InputBox("command")
    Shell "cmd /c " & raw
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA236")
	if len(got) != 1 {
		t.Fatalf("VBA236 findings = %+v, want one finding", got)
	}
	if len(findingsByCode(findings, "VBA224")) != 0 {
		t.Fatalf("process launch must not also be reported as VBA224: %+v", findings)
	}
	if got[0].CommandExecution == nil || got[0].CommandExecution.RiskClass != "injection" || got[0].CommandExecution.OriginState != "tainted" {
		t.Fatalf("command context = %+v", got[0].CommandExecution)
	}
	if got[0].DataFlow == nil || got[0].DataFlow.Sink.Kind != "shell" || !strings.Contains(got[0].Message, "Potential command-injection") {
		t.Fatalf("flow projection = %+v", got[0])
	}
	encoded, err := json.Marshal(got[0])
	if err != nil || !strings.Contains(string(encoded), `"command_execution"`) {
		t.Fatalf("command_execution JSON context = %s, err=%v", encoded, err)
	}

	realtime, err := SourceRealtimeFindings(dir, filepath.Join(dir, "src", "modules", "Main.bas"), config.Default(), []byte(`Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Shell InputBox("command")
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	if len(findingsByCode(realtime, "VBA236")) != 1 || len(findingsByCode(realtime, "VBA224")) != 0 {
		t.Fatalf("realtime process launch projection = %+v", realtime)
	}
}

func TestVBA236DoesNotClaimInjectionForUnknownTransformation(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(raw As String)
    raw = CustomTransform(raw)
    Shell raw
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA236")
	if len(got) != 1 || got[0].CommandExecution == nil {
		t.Fatalf("unknown-origin command finding = %+v", got)
	}
	if got[0].CommandExecution.RiskClass != "process_launch" || got[0].CommandExecution.OriginState != "unknown" {
		t.Fatalf("unknown-origin context = %+v", got[0].CommandExecution)
	}
	if strings.Contains(strings.ToLower(got[0].Message), "injection") {
		t.Fatalf("unknown origin was presented as confirmed injection: %s", got[0].Message)
	}
}

func TestVBA224RetainsNonProcessSinkFlows(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(raw As String)
    Dim cn As Object
    cn.Execute raw
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if len(findingsByCode(findings, "VBA224")) != 1 || len(findingsByCode(findings, "VBA236")) != 0 {
		t.Fatalf("non-process sink ownership = %+v", findings)
	}
}

func TestVBA236RecognizesRunExecInterpretersAndUnquotedPaths(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    Dim raw As String
    Dim sh As Object
    raw = InputBox("command")
    Set sh = CreateObject("WScript.Shell")
    Shell "cmd.exe /c echo " & raw
    sh.Run "powershell.exe -Command " & raw
    sh.Run "powershell.exe -File fixed.ps1 " & raw
    sh.Run "wscript.exe fixed.vbs " & raw
    sh.Exec raw
    Shell "C:\Program Files\Tool\tool.exe"
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA236")
	if len(got) < 4 {
		t.Fatalf("VBA236 findings = %+v, want interpreter, Exec, and unquoted-path findings", got)
	}
	seen := map[string]bool{}
	for _, finding := range got {
		if finding.CommandExecution != nil {
			seen[finding.CommandExecution.Interpreter+":"+finding.CommandExecution.RiskKind] = true
		}
	}
	if !seen["cmd.exe:tainted_command_text"] || !seen["powershell:tainted_command_text"] || !seen[":unquoted_executable_path"] {
		t.Fatalf("command contexts = %+v", got)
	}
	if !seen["powershell:unknown_origin"] {
		t.Fatalf("PowerShell -File argument should remain process-launch risk = %+v", got)
	}
	if !seen["script_host:tainted_command_text"] {
		t.Fatalf("script-host command text should be classified conservatively = %+v", got)
	}
}

func TestVBA236RecognizesShellApplicationAndWin32ShellExecute(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Declare Function ShellExecuteA Lib "shell32.dll" (ByVal hwnd As Long, ByVal operation As String, ByVal file As String, ByVal parameters As String, ByVal directory As String, ByVal show As Long) As Long

Public Sub Run(raw As String)
    Dim sh As Object
    Set sh = CreateObject("Shell.Application")
    sh.ShellExecute raw
    ShellExecuteA 0, "open", raw, "", "", 1
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA236")
	if len(got) < 2 {
		t.Fatalf("ShellExecute findings = %+v, want Shell.Application and Win32 calls", got)
	}
	for _, finding := range got {
		if finding.CommandExecution == nil || finding.CommandExecution.Launcher == "" {
			t.Fatalf("missing ShellExecute context = %+v", finding)
		}
	}
}

func TestVBA236ReportsHiddenUnobservedExecutionAndAcceptsSafeConstant(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    Shell "notepad.exe"
    Shell "notepad.exe", vbHide
    Shell """C:\Program Files\Tool\tool.exe"""
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA236")
	if len(got) != 1 || got[0].CommandExecution == nil || got[0].CommandExecution.RiskKind != "observability" {
		t.Fatalf("hidden/constant findings = %+v", got)
	}
}

func TestVBA236IgnoresUnrelatedRunAndExecMembers(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(raw As String)
    Dim other As Object
    Dim app As Object
    Set app = CreateObject("Shell.Application")
    other.Run raw
    other.Exec raw
    other.ShellExecuteEx raw
    app.Run raw
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA236"); len(got) != 0 {
		t.Fatalf("unrelated Run/Exec findings = %+v", got)
	}
}

func TestVBA236ExcludesShellExecuteEx(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(raw As String)
    Dim other As Object
    other.ShellExecuteEx raw
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA236"); len(got) != 0 {
		t.Fatalf("ShellExecuteEx must remain out of scope: %+v", got)
	}
}

func TestVBA236ExcludesUserDefinedShell(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Function Shell(raw As String) As Long
    Shell = 1
End Function

Public Sub Run(raw As String)
    Shell raw
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA236"); len(got) != 0 {
		t.Fatalf("user-defined Shell must not be treated as a process sink: %+v", got)
	}
}

func TestVBA236FlagsCredentialLikeCommandArguments(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    Shell "tool.exe --password=secret"
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA236")
	if len(got) != 1 || got[0].CommandExecution == nil || got[0].CommandExecution.RiskKind != "credential_exposure" {
		t.Fatalf("credential findings = %+v", got)
	}
	for _, line := range got[0].NearbyCode {
		if strings.Contains(line, "secret") {
			t.Fatalf("credential literal leaked in nearby code: %q", line)
		}
	}
}

func TestVBA236CanBeDisabled(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(raw As String)
    Shell raw
End Sub
`)
	cfg := config.Default()
	cfg.Analyze.DetectUnsafeCommandConstruction = false
	cfg.Analyze.DetectUntrustedDataFlow = false
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA236"); len(got) != 0 {
		t.Fatalf("disabled VBA236 findings = %+v", got)
	}
}

func TestVBA236HonorsInlineSuppression(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(raw As String)
    ' xlflow:disable-next-line VBA236
    Shell raw
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA236"); len(got) != 0 {
		t.Fatalf("suppressed VBA236 findings = %+v", got)
	}
}
