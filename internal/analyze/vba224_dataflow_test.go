package analyze

import (
	"context"
	"encoding/json"
	"path/filepath"
	"sort"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestAnalyzerRunResultContextReturnsCancellation(t *testing.T) {
	ctx, cancel := context.WithCancel(context.Background())
	cancel()

	_, err := (Analyzer{RootDir: t.TempDir(), Config: config.Default()}).RunResultContext(ctx)
	if err != context.Canceled {
		t.Fatalf("RunResultContext error = %v, want context.Canceled", err)
	}
}

func TestAnalyzerRunResultContextReturnsCancellationDuringFileLoop(t *testing.T) {
	dir := t.TempDir()
	for _, module := range []string{"First", "Second", "Third"} {
		writeModule(t, dir, module+".bas", `Attribute VB_Name = "`+module+`"
Option Explicit

Public Sub Run()
End Sub
`)
	}

	ctx := &cancelAfterChecksContext{Context: context.Background(), remaining: 2}
	_, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunResultContext(ctx)
	if err != context.Canceled {
		t.Fatalf("RunResultContext error = %v, want context.Canceled", err)
	}
}

type cancelAfterChecksContext struct {
	context.Context
	remaining int
}

func (c *cancelAfterChecksContext) Err() error {
	if c.remaining <= 0 {
		return context.Canceled
	}
	c.remaining--
	return nil
}

func TestVBA224DetectsDirectAliasAndConcatenation(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run()
    Dim raw As String
    Dim aliasValue As String
    raw = InputBox("command")
    aliasValue = raw
    Shell "cmd /c " & aliasValue
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA224")
	if len(got) != 1 {
		t.Fatalf("VBA224 findings = %+v, want one finding", got)
	}
	if got[0].DataFlow == nil || got[0].DataFlow.Source.Kind != "inputbox" || got[0].DataFlow.Sink.Kind != "shell" {
		t.Fatalf("data flow context = %+v", got[0].DataFlow)
	}
	if len(got[0].DataFlow.Path) < 2 || !strings.Contains(got[0].Message, "Conservative analysis") || !strings.Contains(got[0].Reason, "InputBox") {
		t.Fatalf("incomplete diagnostic context = %+v", got[0])
	}
	encoded, err := json.Marshal(got[0])
	if err != nil || !strings.Contains(string(encoded), `"data_flow"`) {
		t.Fatalf("JSON data_flow context = %s, err=%v", encoded, err)
	}

	realtime, err := SourceRealtimeFindings(dir, filepath.Join(dir, "src", "modules", "Main.bas"), config.Default(), []byte(`Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim raw As String
    raw = InputBox("command")
    Shell raw
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	if len(findingsByCode(realtime, "VBA224")) != 1 {
		t.Fatalf("realtime VBA224 findings = %+v", realtime)
	}
}

func TestVBA224AcceptsConstantsAndExplicitURLSanitization(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(raw As String)
    Dim url As String
    Shell "cmd /c echo fixed"
    url = EncodeURL(raw)
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA224")
	if len(got) != 0 {
		t.Fatalf("safe constant/EncodeURL findings = %+v", got)
	}
}

func TestVBA224KeepsUnknownTransformationsAndHonorsDisable(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(raw As String)
    raw = Trim(raw)
    raw = CustomTransform(raw)
    Shell raw
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA224"); len(got) != 1 || got[0].DataFlow == nil {
		t.Fatalf("unknown transform findings = %+v", got)
	}

	cfg := config.Default()
	cfg.Analyze.DetectUntrustedDataFlow = false
	findings, err = (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA224"); len(got) != 0 {
		t.Fatalf("disabled VBA224 findings = %+v", got)
	}
}

func TestVBA224AcceptsOnlyTheValidatedAllowlistBranch(t *testing.T) {
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(raw As String)
    If raw = "safe" Then
        Shell raw
    Else
        Shell raw
    End If
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA224")
	if len(got) != 1 {
		t.Fatalf("allowlist branch findings = %+v", got)
	}
}

func TestVBA224DetectsProcedureParameter(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(raw As String)
    Shell raw
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA224"); len(got) != 1 {
		t.Fatalf("parameter findings = %+v", got)
	}
}

func TestVBA224HonorsInlineSuppression(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(raw As String)
    ' xlflow:disable-next-line VBA224
    Shell raw
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA224"); len(got) != 0 {
		t.Fatalf("suppressed findings = %+v", got)
	}
}

func TestVBA224AcceptsOnlyTheValidatedSelectCaseBranch(t *testing.T) {
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(raw As String)
    Select Case raw
    Case "safe"
        Shell raw
    Case Else
        Shell raw
    End Select
End Sub
`
	writeModule(t, dir, "Main.bas", source)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA224"); len(got) != 1 {
		t.Fatalf("select case findings = %+v", got)
	}
}

func TestVBA224RecognizesInitialSourceAndSinkCatalogMembers(t *testing.T) {
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(raw As String)
    Dim value As String
    Dim http As Object
    Dim sh As Object
    Dim cn As Object
    Dim rs As Object
    Dim wb As Object
    Set sh = CreateObject("WScript.Shell")
    value = Range("A1").Value
    Shell value
    value = Environ$("PATH")
    sh.Run value
    value = Command$
    sh.Exec value
    value = http.ResponseText
    cn.Execute value
    Kill value
    Workbooks.Open value
    wb.SaveAs value
    http.Open "GET", value
    http.SetRequestHeader "X-Test", value
    Open "input.txt" For Input As #1
    Line Input #1, value
    Shell value
    value = rs.Fields(0).Value
    Shell value
End Sub
	`
	writeModule(t, dir, "Main.bas", source)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA224")
	if len(got) != 11 {
		t.Fatalf("catalog findings = %d, %+v", len(got), got)
	}
	kindSet := map[string]bool{}
	for _, finding := range got {
		if finding.DataFlow == nil {
			t.Fatalf("catalog finding lacks data flow context = %+v", finding)
		}
		kindSet[finding.DataFlow.Sink.Kind] = true
	}
	kinds := make([]string, 0, len(kindSet))
	for kind := range kindSet {
		kinds = append(kinds, kind)
	}
	sort.Strings(kinds)
	wantKinds := []string{"destructive_file_operation", "http_request_header", "http_request_url", "save_as", "shell", "sql_execution", "workbooks_open", "wscript_shell_exec", "wscript_shell_run"}
	if !sameStrings(kinds, wantKinds) {
		t.Fatalf("catalog sink kinds = %v, want %v", kinds, wantKinds)
	}
}

func sameStrings(got, want []string) bool {
	if len(got) != len(want) {
		return false
	}
	for i := range got {
		if got[i] != want[i] {
			return false
		}
	}
	return true
}

func TestVBA224DoesNotMatchUnrelatedMembersOrUnreachableCode(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(raw As String)
    Dim other As Object
    other.Run raw
    other.Execute raw
    ' Shell raw is only a comment
    Exit Sub
    Shell raw
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA224"); len(got) != 0 {
		t.Fatalf("unrelated/unreachable findings = %+v", got)
	}
}
