package analyze

import (
	"encoding/json"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	vbadf "github.com/harumiWeb/xlflow/internal/vba/dataflow"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestVBA236FallbackFlowUsesDeclaredCommandRoles(t *testing.T) {
	t.Parallel()
	flow := vbadf.Finding{
		State: vbadf.StateTainted,
		Sink:  vbadf.Sink{Kind: vbadf.SinkShell, Label: "Shell"},
	}
	if class, kind := commandRiskClassification(flow, string(vbadf.CommandRoleExecutable)); class != "injection" || kind != "tainted_command_text" {
		t.Fatalf("known executable role classification = (%q, %q)", class, kind)
	}
	launcher, _, role := commandLauncherDetails(flow, procedureir.CallSite{}, sourceProcedure{})
	if launcher != "Shell" || role != string(vbadf.CommandRoleUnknown) {
		t.Fatalf("fallback command details = (%q, %q), want Shell/unknown", launcher, role)
	}
	if class, kind := commandRiskClassification(flow, role); class != "process_launch" || kind != "unknown_origin" {
		t.Fatalf("unknown fallback classification = (%q, %q)", class, kind)
	}
}

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
	realtimeVBA236 := findingsByCode(realtime, "VBA236")
	if len(realtimeVBA236) != 1 || len(findingsByCode(realtime, "VBA224")) != 0 {
		t.Fatalf("realtime process launch projection = %+v", realtime)
	}
	context := realtimeVBA236[0].CommandExecution
	if context == nil || context.RiskClass != "injection" || context.RiskKind != "tainted_command_text" || context.CommandRole != "executable" || context.Interpreter != "" {
		t.Fatalf("realtime command context = %+v", context)
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

func TestVBA236TreatsNumericShellArgumentsAsProcessLaunchRisk(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(port As Long)
    Shell "tool.exe --port=" & port
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA236")
	if len(got) != 1 || got[0].CommandExecution == nil {
		t.Fatalf("numeric argument finding = %+v", got)
	}
	context := got[0].CommandExecution
	if context.RiskClass != "process_launch" || context.RiskKind != "unknown_origin" || context.CommandRole != "arguments" {
		t.Fatalf("numeric argument context = %+v", context)
	}
	if strings.Contains(strings.ToLower(got[0].Message), "injection") {
		t.Fatalf("numeric argument was presented as injection: %s", got[0].Message)
	}
}

func TestVBA236TreatsFixedExecutableArgumentConcatenationAsProcessLaunchRisk(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(raw As String)
    Shell "tool.exe --name=" & raw
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA236")
	if len(got) != 1 || got[0].CommandExecution == nil {
		t.Fatalf("argument concatenation finding = %+v", got)
	}
	context := got[0].CommandExecution
	if context.RiskClass != "process_launch" || context.RiskKind != "unknown_origin" || context.CommandRole != "arguments" {
		t.Fatalf("argument concatenation context = %+v", context)
	}
	if strings.Contains(strings.ToLower(got[0].Message), "injection") {
		t.Fatalf("ordinary argument was presented as injection: %s", got[0].Message)
	}
}

func TestVBA236TreatsVbNullStringInitializerOnColonLineAsClean(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    Dim commandStr As String, serviceArgs As String
    commandStr = vbNullString: serviceArgs = vbNullString
    commandStr = commandStr & "tool.exe" & serviceArgs
    Shell commandStr
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA236"); len(got) != 0 {
		t.Fatalf("vbNullString initializer findings = %+v", got)
	}
}

func TestVBA236TreatsModuleLevelConstantQuotingAsClean(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit

Private commandStr As String
Private serviceArgs As String

Public Sub Run(driverPath As String, port As Long, logPath As String)
    commandStr = vbNullString: serviceArgs = vbNullString
    serviceArgs = " --log-path=" & logPath
    commandStr = commandStr & Chr$(34) & driverPath & Chr$(34) & " --port=" & port & serviceArgs
    Shell commandStr
End Sub
`
	writeModule(t, dir, "Main.bas", source)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA236")
	for _, finding := range got {
		if finding.DataFlow != nil && finding.DataFlow.Source.Kind == "unknown" {
			t.Fatalf("fixed Chr$(34) quoting fragment propagated as unknown: path=%+v context=%+v", finding.DataFlow.Path, finding.CommandExecution)
		}
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
	seen := make(map[string]*CommandExecutionContext, len(got))
	for _, finding := range got {
		if finding.CommandExecution == nil || finding.CommandExecution.Launcher == "" {
			t.Fatalf("missing ShellExecute context = %+v", finding)
		}
		seen[finding.CommandExecution.Launcher] = finding.CommandExecution
	}
	app := seen["Shell.Application.ShellExecute"]
	if app == nil || app.CommandRole != "document" || app.RiskClass != "process_launch" || app.RiskKind != "unknown_origin" {
		t.Fatalf("Shell.Application.ShellExecute context = %+v", app)
	}
	win32 := seen["ShellExecuteA"]
	if win32 == nil || win32.CommandRole != "executable" || win32.RiskClass != "injection" || win32.RiskKind != "tainted_command_text" {
		t.Fatalf("ShellExecuteA context = %+v", win32)
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
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA236"); len(got) != 0 {
		t.Fatalf("disabled VBA236 findings = %+v", got)
	}
	if got := findingsByCode(findings, "VBA224"); len(got) != 1 {
		t.Fatalf("VBA224 fallback findings = %+v, want one finding", findings)
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
