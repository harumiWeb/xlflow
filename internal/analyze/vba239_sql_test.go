package analyze

import (
	"encoding/json"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func vba239TestConfig() config.Config {
	cfg := config.Default()
	cfg.Analyze.DetectUnsafeCommandConstruction = false
	cfg.Analyze.DetectUntrustedDataFlow = true
	cfg.Analyze.DetectUnsafeSQLConstruction = true
	return cfg
}

func TestVBA239DetectsConnectionRecordsetAndCommandTextSQL(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(raw As String)
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim cmd As ADODB.Command
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    Set cmd = CreateObject("ADODB.Command")
    cn.Execute "SELECT * FROM Users WHERE Name = '" & raw & "'"
    rs.Open "SELECT * FROM Users WHERE Name LIKE '%" & raw & "%'", cn
    cmd.CommandText = "SELECT * FROM Users WHERE Id = " & raw
    cmd.Execute
End Sub
`)

	findings, err := (Analyzer{RootDir: dir, Config: vba239TestConfig()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA239")
	if len(got) != 3 {
		t.Fatalf("VBA239 findings = %+v, want three", got)
	}
	apis := map[string]bool{}
	risks := map[string]bool{}
	for _, finding := range got {
		if finding.SQLExecution == nil || finding.DataFlow == nil {
			t.Fatalf("missing SQL context = %+v", finding)
		}
		apis[finding.SQLExecution.API] = true
		risks[finding.SQLExecution.RiskKind] = true
		if strings.Contains(strings.ToLower(finding.Message), "confirmed") || strings.Contains(strings.ToLower(finding.Message), "proof") {
			t.Fatalf("overclaiming SQL message = %q", finding.Message)
		}
	}
	for _, api := range []string{"ADODB.Connection.Execute", "ADODB.Recordset.Open", "ADODB.Command.Execute"} {
		if !apis[api] {
			t.Fatalf("missing API %q in %v", api, apis)
		}
	}
	if !risks["manual_quoting"] || !risks["wildcard_like_input"] {
		t.Fatalf("risk kinds = %v", risks)
	}
	encoded, err := json.Marshal(got[0])
	if err != nil || !strings.Contains(string(encoded), `"sql_execution"`) {
		t.Fatalf("SQL JSON context = %s, err=%v", encoded, err)
	}
}

func TestVBA239TracksCellAliasAndProcedureParameter(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(raw As String)
    Dim cn As ADODB.Connection
    Dim cellValue As Variant
    Set cn = CreateObject("ADODB.Connection")
    cellValue = Range("A1").Value
    cn.Execute "SELECT * FROM Users WHERE Name = '" & cellValue & "'"
    cn.Execute "SELECT * FROM Users WHERE Id = " & raw
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: vba239TestConfig()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	got := findingsByCode(findings, "VBA239")
	if len(got) != 2 {
		t.Fatalf("cell/parameter findings = %+v", got)
	}
	kinds := map[string]bool{}
	for _, finding := range got {
		if finding.DataFlow == nil {
			t.Fatalf("missing data flow = %+v", finding)
		}
		kinds[finding.DataFlow.Source.Kind] = true
	}
	if !kinds["worksheet_cell"] || !kinds["parameter"] {
		t.Fatalf("source kinds = %v", kinds)
	}
}

func TestVBA239AcceptsParameterizedCommandAndUnexecutedCommandText(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit

Public Sub Run(raw As String)
    Dim cmd As ADODB.Command
    Set cmd = CreateObject("ADODB.Command")
    cmd.CommandText = "SELECT * FROM Users WHERE Id = ?"
    cmd.Parameters.Append cmd.CreateParameter("id", 3, 1, 20, raw)
    cmd.Execute
    cmd.CommandText = "SELECT * FROM Users WHERE Name = " & raw
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: vba239TestConfig()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA239"); len(got) != 0 {
		t.Fatalf("parameterized/unexecuted findings = %+v", got)
	}
}

func TestVBA239FallsBackToVBA224WhenDisabled(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(raw As String)
    Dim cn As ADODB.Connection
    Set cn = CreateObject("ADODB.Connection")
    cn.Execute "SELECT * FROM Users WHERE Id = " & raw
End Sub
`)
	cfg := vba239TestConfig()
	cfg.Analyze.DetectUnsafeSQLConstruction = false
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if len(findingsByCode(findings, "VBA239")) != 0 || len(findingsByCode(findings, "VBA224")) != 1 {
		t.Fatalf("fallback findings = %+v", findings)
	}
}

func TestVBA239RealtimeAndInlineSuppression(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	source := `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(raw As String)
    ' xlflow:disable-next-line VBA239
    CurrentDb.Execute "SELECT * FROM Users WHERE Id = " & raw
End Sub
`
	path := filepath.Join(dir, "src", "modules", "Main.bas")
	writeModule(t, dir, "Main.bas", source)
	findings, err := (Analyzer{RootDir: dir, Config: vba239TestConfig()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA239"); len(got) != 0 {
		t.Fatalf("suppressed batch findings = %+v", got)
	}
	realtime, err := SourceRealtimeFindings(dir, path, vba239TestConfig(), []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(realtime, "VBA239"); len(got) != 0 {
		t.Fatalf("suppressed realtime findings = %+v", got)
	}
}

func TestVBA239CommandExecuteRecordsAffectedIsNotSQLText(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run(raw As String)
    Dim cmd As ADODB.Command
    Dim recordsAffected As Variant
    Set cmd = CreateObject("ADODB.Command")
    cmd.CommandText = "SELECT * FROM Users WHERE Id = ?"
    cmd.Execute recordsAffected
End Sub
`)
	findings, err := (Analyzer{RootDir: dir, Config: vba239TestConfig()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA239"); len(got) != 0 {
		t.Fatalf("records-affected argument was treated as SQL = %+v", got)
	}
}

func TestVBA224FallbackDoesNotTreatCommandRecordsAffectedAsSQL(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim commandObject As ADODB.Command
    Dim recordsAffected As Variant
    Set commandObject = CreateObject("ADODB.Command")
    commandObject.CommandText = "SELECT 1"
    commandObject.Execute recordsAffected
End Sub
`)
	cfg := vba239TestConfig()
	cfg.Analyze.DetectUnsafeSQLConstruction = false
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(findings, "VBA224"); len(got) != 0 {
		t.Fatalf("records-affected argument leaked through VBA224 fallback = %+v", got)
	}
}
