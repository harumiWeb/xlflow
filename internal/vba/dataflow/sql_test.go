package dataflow

import (
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func analyzeSQLFixture(t *testing.T, source string) Result {
	t.Helper()
	document, err := procedureir.BuildSource(procedureir.BuildOptions{Path: "Main.bas", ModuleKind: "standard"}, []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	procedure := document.Procedures[0]
	result, err := AnalyzeProcedureContext(t.Context(), procedure, cfg.Build(procedure), Options{Conservative: true})
	if err != nil {
		t.Fatal(err)
	}
	return result
}

func TestAnalyzeProcedureSQLFindsConnectionAndRecordsetFlows(t *testing.T) {
	result := analyzeSQLFixture(t, `Option Explicit
Sub Run(raw As String)
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    cn.Execute "SELECT * FROM Users WHERE Name = '" & raw & "'"
    rs.Open "SELECT * FROM Users WHERE Name LIKE '%" & raw & "%'", cn
End Sub
`)
	if len(result.SQLFindings) != 2 {
		t.Fatalf("SQL findings = %#v, want two", result.SQLFindings)
	}
	if result.SQLFindings[0].Execution.API != "ADODB.Connection.Execute" && result.SQLFindings[1].Execution.API != "ADODB.Connection.Execute" {
		t.Fatalf("missing Connection.Execute finding: %#v", result.SQLFindings)
	}
	seen := map[SQLRiskKind]bool{}
	for _, finding := range result.SQLFindings {
		seen[finding.RiskKind] = true
	}
	if !seen[SQLRiskManualQuoting] || !seen[SQLRiskWildcardLikeInput] {
		t.Fatalf("risk kinds = %#v", seen)
	}
}

func TestAnalyzeProcedureSQLFindsCommandTextOnlyWhenExecuted(t *testing.T) {
	result := analyzeSQLFixture(t, `Option Explicit
Sub Run(raw As String)
    Dim cmd As ADODB.Command
    Set cmd = CreateObject("ADODB.Command")
    cmd.CommandText = "SELECT * FROM Users WHERE Id = " & raw
End Sub
`)
	if len(result.SQLFindings) != 0 {
		t.Fatalf("unexecuted CommandText findings = %#v, want none", result.SQLFindings)
	}

	result = analyzeSQLFixture(t, `Option Explicit
Sub Run(raw As String)
    Dim cmd As ADODB.Command
    Set cmd = CreateObject("ADODB.Command")
    cmd.CommandText = "SELECT * FROM Users WHERE Id = " & raw
    cmd.Execute
End Sub
`)
	if len(result.SQLFindings) != 1 || result.SQLFindings[0].Execution.API != "ADODB.Command.Execute" {
		t.Fatalf("executed CommandText findings = %#v", result.SQLFindings)
	}

	result = analyzeSQLFixture(t, `Option Explicit
Sub Run(raw As String)
    Dim original As ADODB.Command
    Dim aliasCommand As ADODB.Command
    Set original = New ADODB.Command
    Set aliasCommand = original
    aliasCommand.CommandText = "SELECT * FROM Users WHERE Id = " & raw
    aliasCommand.Execute
End Sub
`)
	if len(result.SQLFindings) != 1 || result.SQLFindings[0].Execution.API != "ADODB.Command.Execute" {
		t.Fatalf("early-bound/alias CommandText findings = %#v", result.SQLFindings)
	}

	result = analyzeSQLFixture(t, `Option Explicit
Sub Run(raw As String)
    Dim original As ADODB.Command
    Dim aliasCommand As ADODB.Command
    Set original = New ADODB.Command
    Set aliasCommand = original
    original.CommandText = "SELECT * FROM Users WHERE Id = " & raw
    aliasCommand.Execute
End Sub
`)
	if len(result.SQLFindings) != 1 || result.SQLFindings[0].Execution.API != "ADODB.Command.Execute" {
		t.Fatalf("aliased CommandText findings = %#v", result.SQLFindings)
	}
}

func TestAnalyzeProcedureSQLAcceptsParameterizedCommandAndConstants(t *testing.T) {
	result := analyzeSQLFixture(t, `Option Explicit
Sub Run(raw As String)
    Dim cmd As ADODB.Command
    Set cmd = CreateObject("ADODB.Command")
    cmd.CommandText = "SELECT * FROM Users WHERE Id = ?"
    cmd.Parameters.Append cmd.CreateParameter("id", 3, 1, 20, raw)
    cmd.Execute
		CurrentDb.Execute "SELECT 1"
End Sub
`)
	if len(result.SQLFindings) != 0 {
		t.Fatalf("parameterized/constant findings = %#v", result.SQLFindings)
	}
}

func TestAnalyzeProcedureSQLClassifiesLocaleSensitiveAndIdentifiers(t *testing.T) {
	result := analyzeSQLFixture(t, `Option Explicit
Sub Run(tableName As String, whenValue As Date)
    Dim sql As String
    sql = "SELECT * FROM " & tableName & " WHERE CreatedAt = '" & whenValue & "'"
    CurrentDb.Execute sql
End Sub
`)
	if len(result.SQLFindings) != 2 {
		t.Fatalf("SQL findings = %#v, want two source flows", result.SQLFindings)
	}
	joined := ""
	for _, finding := range result.SQLFindings {
		joined += string(finding.RiskKind) + " " + string(finding.Execution.Role) + " "
	}
	if !strings.Contains(joined, string(SQLRiskDynamicIdentifier)) || !strings.Contains(joined, string(SQLRiskLocaleSensitiveValue)) {
		t.Fatalf("risk classification = %q findings=%#v", joined, result.SQLFindings)
	}
}

func TestAnalyzeProcedureSQLClassifiesManualQuotingThroughVariable(t *testing.T) {
	result := analyzeSQLFixture(t, `Option Explicit
Sub Run(name As String)
    Dim sql As String
    sql = "SELECT * FROM Users WHERE Name = '" & name & "'"
    CurrentDb.Execute sql
End Sub
`)
	if len(result.SQLFindings) != 1 || result.SQLFindings[0].RiskKind != SQLRiskManualQuoting {
		t.Fatalf("manual quoting finding = %#v, want one manual_quoting finding", result.SQLFindings)
	}
}

func TestAnalyzeProcedureSQLDoesNotConfuseColumnNameWithIdentifierSource(t *testing.T) {
	result := analyzeSQLFixture(t, `Option Explicit
Sub Run(id As String)
    CurrentDb.Execute "SELECT Id FROM Users WHERE Id = " & id
End Sub
`)
	if len(result.SQLFindings) != 1 {
		t.Fatalf("column/source name findings = %#v", result.SQLFindings)
	}
	if result.SQLFindings[0].Execution.Role != SQLInputRoleValue || result.SQLFindings[0].RiskKind != SQLRiskExternalValueConcatenation {
		t.Fatalf("column/source name classification = %#v", result.SQLFindings[0])
	}
}

func TestAnalyzeProcedureSQLClassifiesCellIdentifierWhenExpressionIsKnown(t *testing.T) {
	result := analyzeSQLFixture(t, `Option Explicit
Sub Run()
    Dim tableName As String
    tableName = Range("A1").Value
    CurrentDb.Execute "SELECT * FROM " & tableName
End Sub
`)
	if len(result.SQLFindings) != 1 || result.SQLFindings[0].Execution.Role != SQLInputRoleIdentifier || result.SQLFindings[0].RiskKind != SQLRiskDynamicIdentifier {
		t.Fatalf("cell identifier classification = %#v", result.SQLFindings)
	}
}

func TestAnalyzeProcedureSQLClassifiesDynamicColumnAndTypedLocal(t *testing.T) {
	result := analyzeSQLFixture(t, `Option Explicit
Sub Run(columnName As String)
    Dim amount As Currency
    amount = InputBox("amount")
    CurrentDb.Execute "SELECT " & columnName & " FROM Users WHERE Amount = " & amount
End Sub
`)
	if len(result.SQLFindings) != 2 {
		t.Fatalf("column/local findings = %#v", result.SQLFindings)
	}
	seen := map[SQLRiskKind]bool{}
	for _, finding := range result.SQLFindings {
		seen[finding.RiskKind] = true
	}
	if !seen[SQLRiskDynamicIdentifier] || !seen[SQLRiskLocaleSensitiveValue] {
		t.Fatalf("column/local risk kinds = %#v", seen)
	}
}

func TestAnalyzeProcedureSQLKeepsRiskyBranchAndDropsSafeOverwrite(t *testing.T) {
	risky := analyzeSQLFixture(t, `Option Explicit
Sub Run(raw As String, useRaw As Boolean)
    Dim cmd As ADODB.Command
    Set cmd = CreateObject("ADODB.Command")
    If useRaw Then
        cmd.CommandText = "SELECT * FROM Users WHERE Id = " & raw
    Else
        cmd.CommandText = "SELECT 1"
    End If
    cmd.Execute
End Sub
`)
	if len(risky.SQLFindings) != 1 {
		t.Fatalf("branch join findings = %#v, want one risky path", risky.SQLFindings)
	}

	safe := analyzeSQLFixture(t, `Option Explicit
Sub Run(raw As String)
    Dim cmd As ADODB.Command
    Set cmd = CreateObject("ADODB.Command")
    cmd.CommandText = "SELECT * FROM Users WHERE Id = " & raw
    cmd.CommandText = "SELECT 1"
    cmd.Execute
End Sub
`)
	if len(safe.SQLFindings) != 0 {
		t.Fatalf("safe overwrite findings = %#v, want none", safe.SQLFindings)
	}
}

func TestAnalyzeProcedureSQLLoopWithConflictingObjectKindsConverges(t *testing.T) {
	result := analyzeSQLFixture(t, `Option Explicit
Sub Run(raw As String, choose As Boolean)
    Dim obj As Object
    Dim i As Long
    For i = 1 To 2
        If choose Then
            Set obj = CreateObject("ADODB.Command")
        Else
            Set obj = CreateObject("ADODB.Recordset")
        End If
    Next i
    obj.CommandText = "SELECT * FROM Users WHERE Id = " & raw
    obj.Execute
End Sub
`)
	if len(result.SQLFindings) != 0 {
		t.Fatalf("conflicting loop object findings = %#v, want none", result.SQLFindings)
	}
}

func TestJoinSQLObjectStateKeepsUnknownAndEmptyIdentity(t *testing.T) {
	command := sqlObjectState{kind: sqlObjectCommand, identity: "left"}
	recordset := sqlObjectState{kind: sqlObjectRecordset, identity: "right"}
	merged := joinSQLObjectState(command, recordset)
	if merged.kind != sqlObjectUnknown || merged.identity != "" {
		t.Fatalf("conflicting object join = %#v, want unknown kind and empty identity", merged)
	}
	recovered := joinSQLObjectState(merged, command)
	if recovered.kind != sqlObjectUnknown || recovered.identity != "" {
		t.Fatalf("absorbing object join = %#v, want unknown kind and empty identity", recovered)
	}
}

func TestAnalyzeProcedureSQLRecognizesDAOSinks(t *testing.T) {
	result := analyzeSQLFixture(t, `Option Explicit
Sub Run(tableName As String, sqlText As String)
    CurrentDb.Execute "DELETE FROM " & tableName
    CurrentDb.OpenRecordset sqlText
    DoCmd.RunSQL "DELETE FROM Users WHERE Id = " & sqlText
End Sub
`)
	if len(result.SQLFindings) != 3 {
		t.Fatalf("DAO SQL findings = %#v, want three", result.SQLFindings)
	}
	apis := map[string]bool{}
	for _, finding := range result.SQLFindings {
		apis[finding.Execution.API] = true
	}
	for _, api := range []string{"Database.Execute", "DAO.OpenRecordset", "DoCmd.RunSQL"} {
		if !apis[api] {
			t.Fatalf("missing DAO API %q in %#v", api, apis)
		}
	}
}

func TestAnalyzeProcedureSQLRecognizesConstantCallByNameMembers(t *testing.T) {
	result := analyzeSQLFixture(t, `Option Explicit
Sub Run(raw As String)
	Dim cn As ADODB.Connection
	Dim cmd As ADODB.Command
	Dim rs As ADODB.Recordset
	Set cn = CreateObject("ADODB.Connection")
	Set cmd = CreateObject("ADODB.Command")
	Set rs = CreateObject("ADODB.Recordset")
	CallByName cn, "Execute", VbMethod, "SELECT * FROM Users WHERE Id = " & raw
	CallByName cmd, "CommandText", VbLet, "SELECT * FROM Users WHERE Id = " & raw
	CallByName cmd, "Execute", VbMethod
	CallByName rs, "Open", VbMethod, "SELECT * FROM Users WHERE Id = " & raw, cn
End Sub
`)
	if len(result.SQLFindings) != 3 {
		t.Fatalf("CallByName SQL findings = %#v", result.SQLFindings)
	}
	apis := map[string]bool{}
	for _, finding := range result.SQLFindings {
		apis[finding.Execution.API] = true
	}
	for _, api := range []string{"ADODB.Connection.Execute", "ADODB.Command.Execute", "ADODB.Recordset.Open"} {
		if !apis[api] {
			t.Fatalf("CallByName API %q missing from %#v", api, apis)
		}
	}
}
