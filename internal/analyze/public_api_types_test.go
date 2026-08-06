package analyze

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/typedb"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestVBA222ChecksPublicProceduresPropertiesEventsAndProjectTypes(t *testing.T) {
	dir := t.TempDir()
	t.Setenv(typedb.EnvDir, filepath.Join(dir, "missing-typelib-db"))
	writeClass(t, dir, "PrivateThing.cls", `Attribute VB_Name = "PrivateThing"
Attribute VB_Exposed = False
Option Explicit
Public Sub InternalMember()
End Sub
`)
	writeClass(t, dir, "IContract.cls", `Attribute VB_Name = "IContract"
Attribute VB_Exposed = True
Option Explicit
Public Function GetPrivate() As PrivateThing
End Function
`)
	writeModule(t, dir, "Main.bas", `Option Explicit
Private Type PrivatePayload
    Value As Long
End Type
Private Enum PrivateStatus
    StatusUnknown = 0
End Enum

Public Function FromPrivate() As PrivateThing
End Function

Public Sub UsePrivate(ByVal value As PrivateThing)
End Sub

Public Sub UsePrivateTypes(ByVal payload As PrivatePayload, ByVal status As PrivateStatus)
End Sub

Public Function KnownExcel(ByVal value As Long) As Workbook
End Function

Public Sub UseContract(ByVal contract As IContract)
End Sub

Public Function MissingExternal(ByVal value As AcmeUnavailable.Widget) As Variant
End Function

Public Sub Auto_Open(ByVal value As UnavailableHostType)
End Sub

Public Property Get Value() As PrivatePayload
End Property

Private Property Let Value(ByVal value As PrivatePayload)
End Property

Private Property Get InternalValue() As MissingPrivateType
End Property

Private Property Let InternalValue(ByVal value As MissingPrivateType)
End Property

Public Event Changed(ByVal value As PrivatePayload)
`)

	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA222")
	if len(got) < 8 {
		t.Fatalf("VBA222 findings = %+v, want procedure/property/event and exposed-interface findings", got)
	}
	for _, finding := range got {
		if finding.Severity != "warning" || !strings.Contains(finding.Message, "type") && !strings.Contains(finding.Message, "Type") {
			t.Fatalf("unexpected VBA222 finding: %+v", finding)
		}
		if finding.Line <= 0 || finding.File == "" || finding.Reason == "" || finding.Suggestion == "" {
			t.Fatalf("VBA222 finding lacks location or explanation: %+v", finding)
		}
	}
	assertFindingMentions(t, got, "PrivateThing")
	assertFindingMentions(t, got, "PrivatePayload")
	assertFindingMentions(t, got, "PrivateStatus")
	assertFindingMentions(t, got, "AcmeUnavailable.Widget")
	if known := findingsWithProcedure(got, "KnownExcel"); len(known) != 0 {
		t.Fatalf("known built-in/Excel types must be allowed: %+v", known)
	}
	if contract := findingsWithProcedure(got, "UseContract"); len(contract) != 0 {
		t.Fatalf("VB_Exposed interface type must be allowed: %+v", contract)
	}
	if host := findingsWithProcedure(got, "Auto_Open"); len(host) != 0 {
		t.Fatalf("host-required Auto_Open signature must be excluded: %+v", host)
	}
	if internal := findingsWithProcedure(got, "InternalMember"); len(internal) != 0 {
		t.Fatalf("members of an unexposed class must not be treated as public API: %+v", internal)
	}
	if internal := findingsWithProcedure(got, "InternalValue"); len(internal) != 0 {
		t.Fatalf("private property accessors must not be treated as public API: %+v", internal)
	}
	if blocking := findingsByCode(BlockingFindings(got), "VBA222"); len(blocking) != 0 {
		t.Fatalf("VBA222 must not block preflight: %+v", blocking)
	}
	if realtime, err := runRealtimeForSingleModule(t, dir, config.Default(), "Main.bas"); err != nil {
		t.Fatal(err)
	} else if got := findingsByCode(realtime, "VBA222"); len(got) != 0 {
		t.Fatalf("VBA222 is batch-only: %+v", got)
	}
}

func TestVBA222ReportsAmbiguousProjectTypes(t *testing.T) {
	dir := t.TempDir()
	t.Setenv(typedb.EnvDir, filepath.Join(dir, "missing-typelib-db"))
	writeModule(t, dir, "First.bas", `Public Type Duplicate
    Value As Long
End Type
`)
	writeModule(t, dir, "Second.bas", `Public Type Duplicate
    Value As Long
End Type
`)
	writeModule(t, dir, "Main.bas", `Public Sub Use(ByVal value As Duplicate)
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA222")
	if len(got) != 1 || !strings.Contains(got[0].Message, "ambiguous") || !strings.Contains(got[0].Message, "Duplicate") {
		t.Fatalf("ambiguous project type finding = %+v", got)
	}
}

func TestVBA222CarriesTypeDBLoadWarnings(t *testing.T) {
	dir := t.TempDir()
	typeDBDir := filepath.Join(dir, "typelib")
	if err := os.MkdirAll(typeDBDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(typeDBDir, "broken.json"), []byte("{"), 0o644); err != nil {
		t.Fatal(err)
	}
	t.Setenv(typedb.EnvDir, typeDBDir)
	writeModule(t, dir, "Main.bas", `Public Sub Run()
End Sub
`)
	result, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunResult()
	if err != nil {
		t.Fatal(err)
	}
	found := false
	for _, warning := range result.Warnings {
		if warning["code"] == "type_db_load_warning" && strings.Contains(warning["message"].(string), "generated type database could not be loaded") {
			found = true
			break
		}
	}
	if !found {
		t.Fatalf("type DB load warning was not carried into the analysis result: %+v", result.Warnings)
	}
}

func TestVBA222ExcludesHostRequiredEventHandlers(t *testing.T) {
	dir := t.TempDir()
	writeWorkbookModule(t, dir, "Sheet1.bas")
	path := filepath.Join(dir, "src", "workbook", "Sheet1.bas")
	body := `Attribute VB_Name = "Sheet1"
Option Explicit
Private Sub Worksheet_Change(ByVal Target As UnavailableHostType)
End Sub
`
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA222"); len(got) != 0 {
		t.Fatalf("host-required event handler signature must be excluded: %+v", got)
	}
}

func TestVBA222HonorsDisableAndInlineSuppression(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Function Suppressed() As MissingType
End Function

' xlflow:disable-next-line VBA222
Public Function InlineSuppressed() As MissingType
End Function
`)
	findings, err := (Analyzer{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA222")
	if len(got) != 1 || got[0].Procedure != "Suppressed" {
		t.Fatalf("inline suppression should keep only Suppressed: %+v", got)
	}
	cfg := config.Default()
	cfg.Analyze.DetectPublicAPITypeSafety = false
	findings, err = (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA222"); len(got) != 0 {
		t.Fatalf("disabled VBA222 should not report: %+v", got)
	}
}

func assertFindingMentions(t *testing.T, findings []Finding, typeName string) {
	t.Helper()
	for _, finding := range findings {
		if strings.Contains(finding.Message, typeName) && strings.Contains(finding.Reason, typeName) && strings.Contains(finding.Suggestion, typeName) {
			return
		}
	}
	t.Fatalf("no VBA222 finding mentioned %s in message/reason/suggestion: %+v", typeName, findings)
}

func findingsWithProcedure(findings []Finding, procedure string) []Finding {
	var out []Finding
	for _, finding := range findings {
		if finding.Procedure == procedure {
			out = append(out, finding)
		}
	}
	return out
}

func runRealtimeForSingleModule(t *testing.T, dir string, cfg config.Config, name string) ([]Finding, error) {
	t.Helper()
	path := filepath.Join(dir, "src", "modules", name)
	source, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	doc, err := vbaast.ParseDocument(path, source)
	if err != nil {
		return nil, err
	}
	defer doc.Close()
	ir, err := procedureir.BuildParsed(procedureir.BuildOptions{RootDir: dir, Path: path, ModuleKind: "standard"}, doc)
	if err != nil {
		return nil, err
	}
	controlFlow := vbacfg.BuildDocument(ir)
	return SourceRealtimeFindingsParsedIRCFG(dir, cfg, doc, ir, controlFlow)
}
