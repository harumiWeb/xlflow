package procedureir

import (
	"context"
	"encoding/json"
	"errors"
	"reflect"
	"strconv"
	"strings"
	"testing"
	"time"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

func TestParameterJSONOmitsUnsetOptionalRanges(t *testing.T) {
	data, err := json.Marshal(Parameter{Name: "value"})
	if err != nil {
		t.Fatal(err)
	}
	encoded := string(data)
	for _, field := range []string{"defaultRange", "boundsRange"} {
		if strings.Contains(encoded, field) {
			t.Fatalf("unset %s should be omitted: %s", field, encoded)
		}
	}
}

type cancelAfterChecksContext struct{ remaining int }

func (c *cancelAfterChecksContext) Deadline() (time.Time, bool) { return time.Time{}, false }
func (c *cancelAfterChecksContext) Done() <-chan struct{}       { return nil }
func (c *cancelAfterChecksContext) Value(any) any               { return nil }
func (c *cancelAfterChecksContext) Err() error {
	if c.remaining <= 0 {
		return context.Canceled
	}
	c.remaining--
	return nil
}

func TestBuildSourceContextReturnsCancellationWithoutPartialIR(t *testing.T) {
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	result, err := BuildSourceContext(ctx, BuildOptions{Path: "Main.bas"}, []byte("Sub Main()\nEnd Sub\n"))
	if !errors.Is(err, context.Canceled) || !reflect.DeepEqual(result, DocumentIR{}) {
		t.Fatalf("canceled result = (%+v, %v)", result, err)
	}
}

func TestBuildParsedContextDiscardsPartiallyBuiltIR(t *testing.T) {
	var source strings.Builder
	for i := 0; i < 400; i++ {
		source.WriteString("Sub Procedure")
		source.WriteString(strconv.Itoa(i))
		source.WriteString("()\n  Dim value As Long\n  value = 1\nEnd Sub\n")
	}
	doc, err := vbaast.ParseDocument("Main.bas", []byte(source.String()))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	result, err := BuildParsedContext(&cancelAfterChecksContext{remaining: 4}, BuildOptions{Path: "Main.bas"}, doc)
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("build error = %v, want context.Canceled", err)
	}
	if !reflect.DeepEqual(result, DocumentIR{}) {
		t.Fatalf("canceled build returned partial IR: %+v", result)
	}
}

func TestBuildSourceNormalizesProceduresStatementsCallsAndScopes(t *testing.T) {
	t.Parallel()
	source := []byte(`Attribute VB_Name = "Module1"
Option Explicit
Private moduleValue As Long

Public Function Compute(ByVal inputValue As Long, Optional label As String = "x") As Long
    Dim localValue As Long
    localValue = inputValue + moduleValue
    Call Helper(localValue)
    Compute = localValue
    Exit Function
Handler:
    Resume Next
End Function
`)
	doc, err := BuildSource(BuildOptions{Path: "src/modules/Module1.bas"}, source)
	if err != nil {
		t.Fatal(err)
	}
	if doc.ModuleName != "Module1" || doc.ModuleKind != "standard" {
		t.Fatalf("unexpected module: %#v", doc)
	}
	if len(doc.Procedures) != 1 {
		t.Fatalf("procedures = %d, want 1", len(doc.Procedures))
	}
	procedure := doc.Procedures[0]
	if procedure.Symbol.Name != "Compute" || procedure.Symbol.Kind != ProcedureFunction {
		t.Fatalf("unexpected procedure symbol: %#v", procedure.Symbol)
	}
	if procedure.Symbol.Visibility != "Public" || procedure.Symbol.ReturnType != "Long" {
		t.Fatalf("unexpected procedure signature: %#v", procedure.Symbol)
	}
	if got := len(procedure.Symbol.Parameters); got != 2 {
		t.Fatalf("parameters = %d, want 2", got)
	}
	if procedure.Symbol.Parameters[0].Passing != "ByVal" || !procedure.Symbol.Parameters[1].Optional {
		t.Fatalf("unexpected parameters: %#v", procedure.Symbol.Parameters)
	}
	requireStatementKind(t, procedure.Statements, StatementDeclaration)
	requireStatementKind(t, procedure.Statements, StatementAssignment)
	requireStatementKind(t, procedure.Statements, StatementCall)
	requireStatementKind(t, procedure.Statements, StatementExit)
	requireStatementKind(t, procedure.Statements, StatementLabel)
	requireStatementKind(t, procedure.Statements, StatementResume)
	if len(procedure.Calls) != 1 || procedure.Calls[0].Callee.BaseName != "Helper" {
		t.Fatalf("unexpected calls: %#v", procedure.Calls)
	}
	assertAccess(t, procedure.Accesses, "inputValue", ScopeParameter, AccessRead)
	assertAccess(t, procedure.Accesses, "moduleValue", ScopeModule, AccessRead)
	assertAccess(t, procedure.Accesses, "localValue", ScopeLocal, AccessWrite)
	if procedure.Symbol.DeclarationRange.StartLine != 5 || procedure.Symbol.BodyRange.StartLine < 5 {
		t.Fatalf("unexpected ranges: %#v", procedure.Symbol)
	}
}

func TestProcedureSignatureRetainsArrayDefaultsAndRecoveryFacts(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Public Sub Signature(ByRef dynamic() As Long, ByVal bounded(1 To 2, 0 To 3) As String, Optional label As String = "x", ParamArray rest() As Variant)
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Procedures) == 0 {
		t.Fatalf("procedures = %#v", doc.Procedures)
	}
	params := doc.Procedures[0].Symbol.Parameters
	if len(params) != 4 {
		t.Fatalf("parameters = %#v", params)
	}
	if !params[0].IsArray || params[0].ArrayShape != ArrayShapeDynamic || params[0].BoundsRange == nil || params[0].BoundsRange.StartByte >= params[0].BoundsRange.EndByte {
		t.Fatalf("dynamic array facts = %#v", params[0])
	}
	if !params[1].IsArray || params[1].ArrayShape != ArrayShapeBounded || len(params[1].ArrayBounds) != 2 {
		t.Fatalf("bounded array facts = %#v", params[1])
	}
	if got := params[1].ArrayBounds[0]; got.Lower != "1" || got.Upper != "2" || got.LowerRange == nil || got.UpperRange == nil || got.LowerRange.StartByte >= got.UpperRange.EndByte {
		t.Fatalf("first bound = %#v", got)
	}
	if !params[2].Optional || !params[2].HasDefault || params[2].Default != `"x"` || params[2].DefaultRange == nil || params[2].DefaultRange.StartByte >= params[2].DefaultRange.EndByte {
		t.Fatalf("optional default facts = %#v", params[2])
	}
	if !params[3].ParamArray || !params[3].IsArray || params[3].ArrayShape != ArrayShapeDynamic || params[3].Passing != "ByRef" || params[3].PassingExplicit {
		t.Fatalf("ParamArray facts = %#v", params[3])
	}
}

func TestProcedureSignatureMarksMalformedArrayBoundsAsInvalid(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Broken.bas"}, []byte("Public Sub Broken(ByVal values(1 To) As Long)\nEnd Sub\n"))
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Procedures) == 0 || len(doc.Procedures[0].Symbol.Parameters) != 1 {
		t.Fatalf("recovered procedure parameters = %#v", doc.Procedures)
	}
	param := doc.Procedures[0].Symbol.Parameters[0]
	if param.ArrayShape != ArrayShapeInvalid || !param.Recovered || (!doc.Parse.HasError && !doc.Parse.HasMissing) {
		t.Fatalf("malformed array facts = %#v parse=%#v", param, doc.Parse)
	}
}

func TestConditionalProcedureSignatureRetainsBranchMetadata(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`#If VBA7 Then
Public Function ConditionalRun(ByVal value As Long) As Long
End Function
#ElseIf Win64 Then
Public Function ConditionalRun(ByVal value As Integer) As Integer
End Function
#Else
Public Function ConditionalRun(Optional value As Long = 1) As Long
End Function
#End If
`))
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Procedures) != 3 {
		t.Fatalf("conditional procedures = %#v", doc.Procedures)
	}
	for i, procedure := range doc.Procedures {
		symbol := procedure.Symbol
		if symbol.Name != "ConditionalRun" || len(symbol.ConditionalBranches) != 1 {
			t.Fatalf("conditional symbol %d = %#v", i, symbol)
		}
		if len(symbol.Parameters) != 1 || (i == 0 && symbol.Parameters[0].Optional) || (i == 1 && symbol.Parameters[0].Optional) || (i == 2 && !symbol.Parameters[0].Optional) {
			t.Fatalf("conditional parameter %d = %#v", i, symbol.Parameters)
		}
		if symbol.ConditionalBranches[0].Branch != i || ((i == 0 || i == 1) && symbol.ConditionalBranches[0].Condition == "") || (i == 2 && symbol.ConditionalBranches[0].Condition != "") {
			t.Fatalf("conditional branch %d = %#v", i, symbol.ConditionalBranches)
		}
	}
}

func TestBuildSourceClassifiesPropertiesAndEvents(t *testing.T) {
	t.Parallel()
	tests := []struct {
		name      string
		opts      BuildOptions
		source    string
		kind      ProcedureKind
		event     bool
		eventKind string
	}{
		{"property get", BuildOptions{Path: "Thing.cls", ModuleKind: "class"}, "Public Property Get Value() As Long\nValue = 1\nEnd Property\n", ProcedurePropertyGet, false, ""},
		{"property let", BuildOptions{Path: "Thing.cls", ModuleKind: "class"}, "Public Property Let Value(ByVal rhs As Long)\nEnd Property\n", ProcedurePropertyLet, false, ""},
		{"property set", BuildOptions{Path: "Thing.cls", ModuleKind: "class"}, "Public Property Set Value(ByVal rhs As Object)\nEnd Property\n", ProcedurePropertySet, false, ""},
		{"workbook", BuildOptions{Path: "ThisWorkbook.cls", ModuleKind: "document"}, "Private Sub Workbook_Open()\nEnd Sub\n", ProcedureSub, true, "workbook"},
		{"workbook after xml export", BuildOptions{Path: "ThisWorkbook.cls", ModuleKind: "document"}, "Private Sub Workbook_AfterXmlExport()\nEnd Sub\n", ProcedureSub, true, "workbook"},
		{"workbook after xml import", BuildOptions{Path: "ThisWorkbook.cls", ModuleKind: "document"}, "Private Sub Workbook_AfterXmlImport()\nEnd Sub\n", ProcedureSub, true, "workbook"},
		{"workbook before xml export", BuildOptions{Path: "ThisWorkbook.cls", ModuleKind: "document"}, "Private Sub Workbook_BeforeXmlExport()\nEnd Sub\n", ProcedureSub, true, "workbook"},
		{"workbook before xml import", BuildOptions{Path: "ThisWorkbook.cls", ModuleKind: "document"}, "Private Sub Workbook_BeforeXmlImport()\nEnd Sub\n", ProcedureSub, true, "workbook"},
		{"workbook sheet pivot table change sync", BuildOptions{Path: "ThisWorkbook.cls", ModuleKind: "document"}, "Private Sub Workbook_SheetPivotTableChangeSync()\nEnd Sub\n", ProcedureSub, true, "workbook"},
		{"worksheet", BuildOptions{Path: "Sheet1.cls", ModuleKind: "document"}, "Private Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\n", ProcedureSub, true, "worksheet"},
		{"workbook helper is not an event", BuildOptions{Path: "ThisWorkbook.cls", ModuleKind: "document"}, "Private Sub Workbook_Helper()\nEnd Sub\n", ProcedureSub, false, ""},
		{"worksheet helper is not an event", BuildOptions{Path: "Sheet1.cls", ModuleKind: "document"}, "Private Sub Worksheet_Helper()\nEnd Sub\n", ProcedureSub, false, ""},
		{"form", BuildOptions{Path: "Dialog.frm", ModuleKind: "form"}, "Private Sub Submit_Click()\nEnd Sub\n", ProcedureSub, true, "userform"},
		{"form test", BuildOptions{Path: "Dialog.frm", ModuleKind: "form"}, "Private Sub Test_Click()\nEnd Sub\n", ProcedureSub, false, ""},
		{"auto", BuildOptions{Path: "Module1.bas"}, "Public Sub Auto_Open()\nEnd Sub\n", ProcedureSub, true, "auto"},
	}
	for _, test := range tests {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			doc, err := BuildSource(test.opts, []byte(test.source))
			if err != nil {
				t.Fatal(err)
			}
			if len(doc.Procedures) != 1 {
				t.Fatalf("procedures = %d", len(doc.Procedures))
			}
			symbol := doc.Procedures[0].Symbol
			if symbol.Kind != test.kind || symbol.IsEventHandler != test.event || symbol.EventKind != test.eventKind {
				t.Fatalf("unexpected symbol: %#v", symbol)
			}
		})
	}
}

func TestBuildParsedDoesNotCloseDocumentAndIRSurvivesClose(t *testing.T) {
	t.Parallel()
	doc, err := vbaast.ParseDocument("Module1.bas", []byte("Public Sub Run()\nMsgBox \"ok\"\nEnd Sub\n"))
	if err != nil {
		t.Fatal(err)
	}
	ir, err := BuildParsed(BuildOptions{}, doc)
	if err != nil {
		t.Fatal(err)
	}
	if err := doc.Read(func(vbaast.ParsedView) error { return nil }); err != nil {
		t.Fatalf("BuildParsed closed caller document: %v", err)
	}
	doc.Close()
	if len(ir.Procedures) != 1 || len(ir.Procedures[0].Calls) != 1 {
		t.Fatalf("IR unusable after close: %#v", ir)
	}
	if containsTreeSitterType(reflect.TypeOf(ir), map[reflect.Type]bool{}) {
		t.Fatal("DocumentIR retains a tree-sitter type")
	}
}

func TestBuildSourceReturnsPartialIRForRecoveredSource(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Broken.bas"}, []byte("Public Sub Broken(\nDim value As Long\nvalue =\nEnd Sub\n"))
	if err != nil {
		t.Fatal(err)
	}
	if !doc.Parse.HasError && !doc.Parse.HasMissing {
		t.Fatalf("expected parser recovery metadata: %#v", doc.Parse)
	}
	if len(doc.Procedures) == 0 {
		t.Fatal("expected recoverable procedure facts")
	}
	if !doc.Procedures[0].Symbol.Recovered && !hasRecoveredStatement(doc.Procedures[0].Statements) {
		t.Fatal("expected recovered IR facts")
	}
}

func TestTypeReferencesAndEventDeclaration(t *testing.T) {
	t.Parallel()
	source := []byte(`Public Event Changed(ByVal value As Long)
Implements IFoo
Private service As IService
Public Sub Run()
    Dim item As New Widget
    Set item = New Widget
End Sub
`)
	doc, err := BuildSource(BuildOptions{Path: "Thing.cls", ModuleKind: "class"}, source)
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Procedures) != 1 || doc.Procedures[0].Symbol.Name != "Run" {
		t.Fatalf("Event declaration was treated as a procedure: %#v", doc.Procedures)
	}
	if len(doc.Declarations) == 0 || len(doc.Declarations[0].Parameters) != 1 || doc.Declarations[0].Parameters[0].Type != "Long" {
		t.Fatalf("Event parameters were not retained in the IR: %#v", doc.Declarations)
	}
	assertTypeReference(t, doc.TypeReferences, "implements", "IFoo")
	assertTypeReference(t, doc.TypeReferences, "uses_type", "IService")
	assertTypeReference(t, doc.TypeReferences, "uses_type", "Widget")
	assertTypeReference(t, doc.TypeReferences, "constructs", "Widget")
}

func TestModuleAttributesAndPublicTypesAreRetained(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Contract.cls", ModuleKind: "class"}, []byte(`Attribute VB_Name = "Contract"
Attribute VB_Exposed = True
Private Type PrivatePayload
    Value As Long
End Type
Public Type PublicPayload
    Value As Long
End Type
Public Enum StatusCode
    StatusOK = 1
End Enum
`))
	if err != nil {
		t.Fatal(err)
	}
	if !moduleAttributeValue(doc, "VB_Exposed", "True") {
		t.Fatalf("module attributes = %#v", doc.ModuleAttributes)
	}
	seen := map[string]string{}
	for _, declaration := range doc.Declarations {
		seen[declaration.Name] = declaration.Kind + ":" + declaration.Visibility
	}
	for name, want := range map[string]string{
		"PrivatePayload": "type:Private",
		"PublicPayload":  "type:Public",
		"StatusCode":     "enum:Public",
	} {
		if seen[name] != want {
			t.Fatalf("declaration %s = %q, want %q; all=%#v", name, seen[name], want, seen)
		}
	}
}

func moduleAttributeValue(doc DocumentIR, name, want string) bool {
	for _, attribute := range doc.ModuleAttributes {
		if strings.EqualFold(attribute.Name, name) && strings.EqualFold(strings.Trim(attribute.Value, `"`), want) {
			return true
		}
	}
	return false
}

func TestStatementHierarchyAndErrorHandlingKinds(t *testing.T) {
	t.Parallel()
	source := []byte(`Public Sub Run()
    On Error GoTo Handler
    If Ready Then
        For Each item In items
            Call Work(item)
        Next
    Else
        GoTo Done
    End If
    Exit Sub
Handler:
    Resume Next
Done:
End Sub
`)
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, source)
	if err != nil {
		t.Fatal(err)
	}
	statements := doc.Procedures[0].Statements
	for _, kind := range []StatementKind{
		StatementOnError, StatementIf, StatementForEach, StatementCall,
		StatementElse, StatementGoTo, StatementExit, StatementLabel, StatementResume,
	} {
		requireStatementKind(t, statements, kind)
	}
	for _, statement := range statements {
		if statement.Kind == StatementCall && statement.ParentID == 0 {
			t.Fatalf("nested call has no parent: %#v", statement)
		}
		if statement.Kind == StatementOnError && !strings.EqualFold(statement.Label, "Handler") {
			t.Fatalf("On Error target = %q", statement.Label)
		}
	}
}

func TestNonLabelErrorHandlingFormsHaveNoLabel(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Public Sub Run()
    On Error Resume Next
    On Error GoTo 0
    On Error GoTo [Next]
    Resume
    Resume Next
    Resume [Next]
    Resume Handler
Handler:
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	var onErrorLabels, resumeLabels []string
	for _, statement := range doc.Procedures[0].Statements {
		switch statement.Kind {
		case StatementOnError:
			onErrorLabels = append(onErrorLabels, statement.Label)
		case StatementResume:
			resumeLabels = append(resumeLabels, statement.Label)
		}
	}
	if !reflect.DeepEqual(onErrorLabels, []string{"", "", "Next"}) {
		t.Fatalf("On Error labels = %#v, want non-label forms empty", onErrorLabels)
	}
	if !reflect.DeepEqual(resumeLabels, []string{"", "", "Next", "Handler"}) {
		t.Fatalf("Resume labels = %#v, want only explicit handler label", resumeLabels)
	}
}

func TestControlFlowMetadataNormalizesBranchesCasesLoopsAndTransfers(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Public Function Run() As Long
    If Ready Then Work 1: GoTo Done Else Work 2: End
    Select Case value
    Case 1
        Work 3
    Case Else
        Work 4
    End Select
    Do While Ready
        Exit Do
    Loop
    Do Until Ready
    Loop
    Do
    Loop While Ready
    Do
    Loop Until Ready
    For index = 1 To 2
        Exit For
    Next
    On Error GoTo Handler
    On Error Resume Next
    On Error GoTo 0
    Resume
    Resume Next
    Resume [Next]
    Resume Handler
    Exit Function
Done:
    Exit Sub
Handler:
    Exit Property
End Function
`))
	if err != nil {
		t.Fatal(err)
	}
	statements := doc.Procedures[0].Statements
	assertControl(t, statements, "Work 1", BranchThen, "", "", "", "")
	assertControl(t, statements, "GoTo Done", BranchThen, "", TransferGoto, "Done", "")
	assertControl(t, statements, "Work 2", BranchElse, "", "", "", "")
	assertControl(t, statements, "End", BranchElse, "", TransferTerminate, "", "")
	assertControl(t, statements, "Case Else", "", "", "", "", "case_else")
	assertControl(t, statements, "Do While Ready", "", LoopPreWhile, "", "", "")
	assertControl(t, statements, "Do Until Ready", "", LoopPreUntil, "", "", "")
	assertControl(t, statements, "Loop While Ready", "", LoopPostWhile, "", "", "")
	assertControl(t, statements, "Loop Until Ready", "", LoopPostUntil, "", "", "")
	assertControl(t, statements, "Exit Do", "", "", TransferExitDo, "", "")
	assertControl(t, statements, "Exit For", "", "", TransferExitFor, "", "")
	assertControl(t, statements, "On Error GoTo Handler", "", "", TransferOnErrorGoto, "Handler", "")
	assertControl(t, statements, "On Error Resume Next", "", "", TransferOnErrorResumeNext, "", "")
	assertControl(t, statements, "On Error GoTo 0", "", "", TransferOnErrorDisable, "", "")
	assertControl(t, statements, "Resume", "", "", TransferResumeRetry, "", "")
	assertControl(t, statements, "Resume Next", "", "", TransferResumeNext, "", "")
	assertControl(t, statements, "Resume [Next]", "", "", TransferResumeLabel, "Next", "")
	assertControl(t, statements, "Resume Handler", "", "", TransferResumeLabel, "Handler", "")
	assertControl(t, statements, "Exit Function", "", "", TransferExitFunction, "", "")
	assertControl(t, statements, "Exit Sub", "", "", TransferExitSub, "", "")
	assertControl(t, statements, "Exit Property", "", "", TransferExitProperty, "", "")
}

func TestNumericLineLabelKeepsNestedExecutableStatement(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Public Sub Run()
100 GoTo 200
200 Work
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	statements := doc.Procedures[0].Statements
	if len(statements) != 4 {
		t.Fatalf("statements = %+v, want two labels and two nested executable statements", statements)
	}
	if statements[0].Kind != StatementLabel || statements[0].Label != "100" ||
		statements[1].Kind != StatementGoTo || statements[1].ParentID != statements[0].ID ||
		statements[1].Control == nil || statements[1].Control.Target != "200" {
		t.Fatalf("first numbered statement = %+v / %+v", statements[0], statements[1])
	}
	if statements[0].Text != "100" || statements[0].Range.EndByte > statements[1].Range.StartByte {
		t.Fatalf("numeric label range overlaps nested statement: %+v / %+v", statements[0], statements[1])
	}
	if statements[2].Kind != StatementLabel || statements[2].Label != "200" ||
		statements[3].Kind != StatementCall || statements[3].ParentID != statements[2].ID {
		t.Fatalf("second numbered statement = %+v / %+v", statements[2], statements[3])
	}
}

func TestSetAssignmentRecordsObjectWriteAndRead(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Public Sub Run()
    Dim target As Object
    Dim source As Object
    Set target = source
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	procedure := doc.Procedures[0]
	requireStatementKind(t, procedure.Statements, StatementSet)
	assertAccess(t, procedure.Accesses, "target", ScopeLocal, AccessWrite)
	assertAccess(t, procedure.Accesses, "source", ScopeLocal, AccessRead)
}

func TestCompositeAssignmentTargetsKeepReceiverAndIndexesAsReads(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Public Sub Run()
    Dim obj As Object
    Dim values() As Long
    Dim index As Long
    Dim source As Long
    obj.Value = source
    values(index) = source
    Let values(index) = source
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	modes := map[string][]AccessMode{}
	for _, access := range doc.Procedures[0].Accesses {
		key := strings.ToLower(access.Name)
		modes[key] = append(modes[key], access.Mode)
	}
	if !reflect.DeepEqual(modes["obj"], []AccessMode{AccessRead}) {
		t.Fatalf("object receiver modes = %v, want read", modes["obj"])
	}
	if !reflect.DeepEqual(modes["values"], []AccessMode{AccessWrite, AccessWrite}) {
		t.Fatalf("indexed assignment base modes = %v, want writes; accesses=%+v",
			modes["values"], doc.Procedures[0].Accesses)
	}
	if !reflect.DeepEqual(modes["index"], []AccessMode{AccessRead, AccessRead}) {
		t.Fatalf("index modes = %v, want reads", modes["index"])
	}
	if !reflect.DeepEqual(modes["source"], []AccessMode{AccessRead, AccessRead, AccessRead}) {
		t.Fatalf("assignment value modes = %v, want reads", modes["source"])
	}
}

func TestParenthesizedComparisonArgumentRemainsCallAndReads(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Public Sub Run()
    Dim actual As Long
    Dim expected As Long
    Check (actual) = expected
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	procedure := doc.Procedures[0]
	var statement *Statement
	for i := range procedure.Statements {
		if procedure.Statements[i].Range.StartLine == 4 {
			statement = &procedure.Statements[i]
			break
		}
	}
	if statement == nil || statement.Kind != StatementCall {
		t.Fatalf("comparison argument statement = %+v, want call", statement)
	}
	if len(procedure.Calls) != 1 || !strings.EqualFold(procedure.Calls[0].Callee.BaseName, "Check") {
		t.Fatalf("comparison argument call was lost: %+v", procedure.Calls)
	}
	assertAccess(t, procedure.Accesses, "actual", ScopeLocal, AccessRead)
	assertAccess(t, procedure.Accesses, "expected", ScopeLocal, AccessRead)
	for _, access := range procedure.Accesses {
		if strings.EqualFold(access.Name, "Check") {
			t.Fatalf("callee became a variable access: %+v", access)
		}
	}
}

func TestCallAccessesExcludeOnlyCalleeAndNamedArgumentLabels(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Public Sub Run()
    Dim Foo As Long
    Dim value As Long
    Call Foo(Foo)
    Call Target(Bar:=value)
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	counts := map[string]int{}
	for _, access := range doc.Procedures[0].Accesses {
		counts[strings.ToLower(access.Name)]++
	}
	if counts["foo"] != 1 {
		t.Fatalf("Foo accesses = %d, want only the argument read: %+v", counts["foo"], doc.Procedures[0].Accesses)
	}
	if counts["bar"] != 0 {
		t.Fatalf("named argument label became a variable read: %+v", doc.Procedures[0].Accesses)
	}
	if counts["value"] != 1 {
		t.Fatalf("value accesses = %d, want argument read: %+v", counts["value"], doc.Procedures[0].Accesses)
	}
}

func TestRangesPreserveUTF8ByteColumnsAndCRLFOffsets(t *testing.T) {
	t.Parallel()
	source := []byte("Public Sub Run()\r\n    MsgBox \"😀\", Title:=\"日本語\"\r\nEnd Sub\r\n")
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, source)
	if err != nil {
		t.Fatal(err)
	}
	call := doc.Procedures[0].Calls[0]
	if call.Range.StartLine != 2 || call.Range.StartColumn != 5 {
		t.Fatalf("unexpected call range: %#v", call.Range)
	}
	if got := string(source[call.Range.StartByte:call.Range.EndByte]); !strings.Contains(got, "😀") {
		t.Fatalf("byte offsets do not address original UTF-8 source: %q", got)
	}
	if call.Arguments.Count != 2 || len(call.Arguments.Named) != 1 || call.Arguments.Named[0].Name != "Title" {
		t.Fatalf("unexpected arguments: %#v", call.Arguments)
	}
}

func TestResolveIsRepeatableAndDoesNotMutateSyntaxIR(t *testing.T) {
	t.Parallel()
	source, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Private projectValue As Long
Public Sub Run()
    MissingCall projectValue
    MsgBox "ok"
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	resolver := NewSymbolResolver([]ResolverSymbol{
		{Name: "MissingCall", Module: "Module2", Kind: "sub", Visibility: "Public", File: "Module2.bas", Line: 4},
		{Name: "projectName", Module: "Module2", Kind: "module_variable", Visibility: "Public", File: "Module2.bas", Line: 1},
	})
	resolved := Resolve(source, resolver)
	if source.Procedures[0].Calls[0].Resolution.Status != ResolutionNotAttempted {
		t.Fatal("Resolve mutated source")
	}
	if got := resolved.Procedures[0].Calls[0].Resolution.Status; got != ResolutionMatched {
		t.Fatalf("call status = %q", got)
	}
	if got := resolved.Procedures[0].Calls[1].Resolution.Status; got != ResolutionBuiltinLike {
		t.Fatalf("builtin status = %q", got)
	}
	again := Resolve(resolved, NewSymbolResolver(nil))
	if got := again.Procedures[0].Calls[0].Resolution.Status; got != ResolutionUnresolved {
		t.Fatalf("re-resolution status = %q", got)
	}
}

func TestResolverClassifiesEffectfulBuiltins(t *testing.T) {
	t.Parallel()
	resolver := NewResolver(nil)
	for _, name := range []string{"Shell", "Error"} {
		resolution := resolver.ResolveCall(CallSite{Callee: Callee{Text: name, BaseName: name, Member: name}})
		if resolution.Status != ResolutionBuiltinLike {
			t.Fatalf("%s status = %q, want %q", name, resolution.Status, ResolutionBuiltinLike)
		}
	}
}

func TestResolverRequiresReceiverForCrossModuleClassProcedure(t *testing.T) {
	t.Parallel()
	resolver := NewResolver([]ResolverSymbol{{
		Name: "RestoreEvents", Module: "StateClass", ModuleKind: "class",
		Kind: "sub", Visibility: "Public", File: "StateClass.cls", Line: 2,
	}})
	bare := resolver.ResolveCall(CallSite{
		Caller: ProcedureRef{Name: "Run", QualifiedName: "Main.Run"},
		Callee: Callee{Text: "RestoreEvents", BaseName: "RestoreEvents", Member: "RestoreEvents"},
	})
	if bare.Status != ResolutionUnresolved {
		t.Fatalf("bare class call status = %q, want %q", bare.Status, ResolutionUnresolved)
	}
	receiver := "StateClass"
	qualified := resolver.ResolveCall(CallSite{
		Caller: ProcedureRef{Name: "Run", QualifiedName: "Main.Run"},
		Callee: Callee{Text: "StateClass.RestoreEvents", BaseName: "RestoreEvents", Receiver: &receiver, Member: "RestoreEvents"},
	})
	if qualified.Status != ResolutionMatched {
		t.Fatalf("qualified class call status = %q, want %q", qualified.Status, ResolutionMatched)
	}
}

func TestResolveSymbolOverlayAndPrivateVisibility(t *testing.T) {
	t.Parallel()
	source, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Public Sub Run()
    projectName = externalValue + implicitHidden + implicitField + implicitWithEvents
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	resolver := NewResolver([]ResolverSymbol{
		{Name: "projectName", Module: "Module2", Kind: "module_variable", Visibility: "Public", File: "Module2.bas", Line: 1},
		{Name: "externalValue", Module: "Module2", Kind: "module_variable", Visibility: "Private", File: "Module2.bas", Line: 2},
		{Name: "implicitHidden", Module: "Module2", Kind: "module_variable", File: "Module2.bas", Line: 3},
		{Name: "implicitField", Module: "Class1", Kind: "field", File: "Class1.cls", Line: 1},
		{Name: "implicitWithEvents", Module: "Class1", Kind: "withevents_field", File: "Class1.cls", Line: 2},
	})
	resolved := Resolve(source, resolver)
	assertAccess(t, resolved.Procedures[0].Accesses, "projectName", ScopeProject, AccessWrite)
	assertAccess(t, resolved.Procedures[0].Accesses, "externalValue", ScopeUnresolved, AccessRead)
	assertAccess(t, resolved.Procedures[0].Accesses, "implicitHidden", ScopeUnresolved, AccessRead)
	assertAccess(t, resolved.Procedures[0].Accesses, "implicitField", ScopeUnresolved, AccessRead)
	assertAccess(t, resolved.Procedures[0].Accesses, "implicitWithEvents", ScopeUnresolved, AccessRead)
	sameModule := resolver.ResolveSymbol(SymbolReference{
		Name:   "implicitHidden",
		Caller: ProcedureRef{Name: "Run", Kind: ProcedureSub, QualifiedName: "Module2.Run"},
	})
	if sameModule.Scope != ScopeProject || len(sameModule.Candidates) != 1 {
		t.Fatalf("same-module implicit private resolution = %+v", sameModule)
	}
}

func TestModuleDeclarationsAfterProcedureParticipateInScopeResolution(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Public Sub Run()
    privateAfter = publicAfter
End Sub

Private privateAfter As Long
Public publicAfter As Long
`))
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Declarations) != 2 {
		t.Fatalf("module declarations = %d, want 2: %#v", len(doc.Declarations), doc.Declarations)
	}
	var privateDecl, publicDecl Declaration
	for _, declaration := range doc.Declarations {
		switch declaration.Name {
		case "privateAfter":
			privateDecl = declaration
		case "publicAfter":
			publicDecl = declaration
		}
	}
	if privateDecl.Scope != ScopeModule || publicDecl.Scope != ScopeProject {
		t.Fatalf("unexpected declaration scopes: private=%q public=%q", privateDecl.Scope, publicDecl.Scope)
	}
	assertAccess(t, doc.Procedures[0].Accesses, "privateAfter", ScopeModule, AccessWrite)
	assertAccess(t, doc.Procedures[0].Accesses, "publicAfter", ScopeModule, AccessRead)
}

func TestDeclarationValueShapesRetainArrayBoundsAndConstness(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Shapes.bas"}, []byte(`Option Explicit
Public Const Limit As Long = 2
Public fixed(1 To Limit, 0 To 3) As Long
Public dynamic() As Long
Public scalar As Long
Public Sub Run()
  Dim local(2 To 4) As Long
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	var fixed, dynamic, scalar, constant Declaration
	for _, declaration := range doc.Declarations {
		switch declaration.Name {
		case "fixed":
			fixed = declaration
		case "dynamic":
			dynamic = declaration
		case "scalar":
			scalar = declaration
		case "Limit":
			constant = declaration
		}
	}
	if fixed.ValueShape != ValueShapeFixedArray || len(fixed.ArrayBounds) != 2 || fixed.ArrayBounds[0].Lower != "1" || fixed.ArrayBounds[0].Upper != "Limit" {
		t.Fatalf("fixed declaration shape = %#v", fixed)
	}
	if dynamic.ValueShape != ValueShapeDynamicArray || len(dynamic.ArrayBounds) != 0 {
		t.Fatalf("dynamic declaration shape = %#v", dynamic)
	}
	if scalar.ValueShape != ValueShapeScalar || constant.ValueShape != ValueShapeScalar || !constant.IsConst {
		t.Fatalf("scalar/constant shapes = scalar=%#v const=%#v", scalar, constant)
	}
	if len(doc.Procedures) != 1 || len(doc.Procedures[0].Declarations) != 1 || doc.Procedures[0].Declarations[0].ValueShape != ValueShapeFixedArray {
		t.Fatalf("local declaration shape = %#v", doc.Procedures)
	}
}

func TestProcedureReturnValueShapeRetainsArraySyntax(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Returns.bas"}, []byte(`Option Explicit
Public Function FixedValues() As Long(1 To 2)
End Function
Public Function DynamicValues() As Variant()
End Function
`))
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Procedures) != 2 {
		t.Fatalf("procedures = %#v", doc.Procedures)
	}
	fixed, dynamic := doc.Procedures[0].Symbol, doc.Procedures[1].Symbol
	if !fixed.IsArray || fixed.ValueShape != ValueShapeFixedArray || len(fixed.ArrayBounds) != 1 {
		t.Fatalf("fixed return shape = %#v", fixed)
	}
	if !dynamic.IsArray || dynamic.ValueShape != ValueShapeDynamicArray {
		t.Fatalf("dynamic return shape = %#v", dynamic)
	}
}

func TestClonePreservesArrayShapeBoundsOwnership(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Shapes.bas"}, []byte(`Public fixed(1 To 2) As Long
Public Function Values() As Long(1 To 2)
End Function
`))
	if err != nil {
		t.Fatal(err)
	}
	clone := Clone(doc)
	if len(clone.Declarations) == 0 || len(clone.Declarations[0].ArrayBounds) != 1 || len(clone.Procedures[0].Symbol.ArrayBounds) != 1 {
		t.Fatalf("clone lost array bounds: %#v", clone)
	}
	clone.Declarations[0].ArrayBounds[0].LowerRange.StartLine = 99
	clone.Procedures[0].Symbol.ArrayBounds[0].UpperRange.StartLine = 99
	if doc.Declarations[0].ArrayBounds[0].LowerRange.StartLine == 99 || doc.Procedures[0].Symbol.ArrayBounds[0].UpperRange.StartLine == 99 {
		t.Fatal("Clone shares array-bound range pointers")
	}
}

func TestVisibilityOnlyUsesDeclarationHeader(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Const AccessLabel As String = "Public"

Sub Run(Optional value As String = "Private")
    Debug.Print "Friend"
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	if got := doc.Declarations[0]; got.Visibility != "" || got.Scope != ScopeModule {
		t.Fatalf("module declaration visibility leaked from initializer: %+v", got)
	}
	if got := doc.Procedures[0].Symbol.Visibility; got != "" {
		t.Fatalf("procedure visibility leaked from signature/body text: %q", got)
	}
	parameter := doc.Procedures[0].Symbol.Parameters[0]
	if parameter.Passing != "ByRef" || !parameter.Optional {
		t.Fatalf("parameter modifiers leaked from default text: %+v", parameter)
	}
}

func TestEmptyProcedureHasNoBodyFacts(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte("Public Sub Empty()\nEnd Sub\n"))
	if err != nil {
		t.Fatal(err)
	}
	procedure := doc.Procedures[0]
	if len(procedure.Statements) != 0 || len(procedure.Expressions) != 0 ||
		len(procedure.Calls) != 0 || len(procedure.Accesses) != 0 {
		t.Fatalf("empty procedure contains body facts: %+v", procedure)
	}
	if procedure.Symbol.BodyRange.StartByte != procedure.Symbol.BodyRange.EndByte ||
		procedure.Symbol.BodyRange.StartLine != 2 {
		t.Fatalf("empty body range = %+v, want zero-width range at End Sub", procedure.Symbol.BodyRange)
	}
}

func TestCurrentModulePublicDeclarationShadowsProjectSymbols(t *testing.T) {
	t.Parallel()
	source, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Public sharedValue As Long
Public Sub Run()
    sharedValue = sharedValue + 1
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	resolved := Resolve(source, NewResolver([]ResolverSymbol{{
		Name: "sharedValue", Module: "Other", Kind: "module_variable",
		Visibility: "Public", File: "Other.bas", Line: 1,
	}}))
	for _, access := range resolved.Procedures[0].Accesses {
		if strings.EqualFold(access.Name, "sharedValue") &&
			(access.Scope != ScopeModule || len(access.Resolution.Candidates) != 0) {
			t.Fatalf("current-module binding was replaced by project resolution: %+v", access)
		}
	}
}

func TestCanonicalExpressionLinksIncludeWrappedCallArguments(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Public Sub Run(ByVal a As Long, ByVal b As Long)
    Dim result As Long
    result = Foo(a + 1, named:=Bar(b))
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	procedure := doc.Procedures[0]
	var assignment *Statement
	for i := range procedure.Statements {
		if procedure.Statements[i].Kind == StatementAssignment {
			assignment = &procedure.Statements[i]
			break
		}
	}
	if assignment == nil || assignment.TargetID == 0 || assignment.ValueID == 0 {
		t.Fatalf("assignment lacks canonical target/value links: %+v", assignment)
	}
	if assignment.Target == nil || assignment.Target.ID != assignment.TargetID ||
		assignment.Value == nil || assignment.Value.ID != assignment.ValueID {
		t.Fatalf("statement pointers are detached from canonical expressions: %+v", assignment)
	}
	foo := procedure.Calls[0]
	if foo.Arguments.Count != 2 || len(foo.Arguments.ExpressionIDs) != 2 ||
		foo.Arguments.ExpressionIDs[0] == 0 || foo.Arguments.ExpressionIDs[1] == 0 {
		t.Fatalf("wrapped argument expressions are not linked: %+v", foo.Arguments)
	}
	if len(foo.Arguments.Named) != 1 || foo.Arguments.Named[0].ExpressionID == 0 {
		t.Fatalf("named argument value lacks expression link: %+v", foo.Arguments.Named)
	}
}

func TestCallArgumentsIgnoreLineContinuationNodes(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Private Sub Consume(ByRef values() As Byte, ByVal count As Long)
End Sub

Public Sub Run()
    Dim values() As Byte
    ReDim values(0 To 1)
    Consume values, _
        2
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Procedures) != 2 || len(doc.Procedures[1].Calls) != 1 {
		t.Fatalf("line-continuation call was not captured: %+v", doc.Procedures)
	}
	call := doc.Procedures[1].Calls[0]
	if call.Arguments.Count != 2 || len(call.Arguments.ExpressionIDs) != 2 ||
		call.Arguments.ExpressionIDs[0] == 0 || call.Arguments.ExpressionIDs[1] == 0 {
		t.Fatalf("line continuation must not become an argument slot: %+v", call.Arguments)
	}
}

func TestRecoveredUnparenthesizedCallWithQualifiedArguments(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Private Sub NDArrayToCollections(ByRef arr() As Variant, ByRef src As Variant, ByRef bounds As Variant, ByVal totalElem As Long, ByRef cDimsPtr As LongPtr)
End Sub

Public Sub Run()
    NDArrayToCollections .arrItems, vars.arr(0), bounds _
                         , totalElem, ptrs.arr(0)
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Procedures) != 2 || len(doc.Procedures[1].Calls) != 1 {
		t.Fatalf("recovered call was not captured: %+v", doc.Procedures)
	}
	call := doc.Procedures[1].Calls[0]
	if !strings.EqualFold(call.Callee.BaseName, "NDArrayToCollections") {
		t.Fatalf("callee = %+v, want NDArrayToCollections", call.Callee)
	}
	if call.Arguments.Count != 5 || len(call.Arguments.ExpressionIDs) != 5 {
		t.Fatalf("recovered call arguments = %+v, want five arguments", call.Arguments)
	}
	want := []string{".arrItems", "vars.arr(0)", "bounds", "totalElem", "ptrs.arr(0)"}
	for i, text := range want {
		expressionID := call.Arguments.ExpressionIDs[i]
		if expressionID == 0 || expressionID > len(doc.Procedures[1].Expressions) {
			t.Fatalf("argument %d has invalid expression ID %d: %+v", i, expressionID, call.Arguments)
		}
		expression := doc.Procedures[1].Expressions[expressionID-1]
		if expression.Text != text || !expression.Recovered {
			t.Fatalf("argument %d = %+v, want recovered expression %q", i, expression, text)
		}
	}
}

func TestRecoveredUnparenthesizedCallPreservesOmittedArgumentSlots(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Private Sub Target(Optional first As Variant, Optional second As Variant, Optional third As Variant)
End Sub

Public Sub Run()
    Target .first, _
           , third
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Procedures) != 2 || len(doc.Procedures[1].Calls) != 1 {
		t.Fatalf("recovered call was not captured: %+v", doc.Procedures)
	}
	call := doc.Procedures[1].Calls[0]
	if call.Arguments.Count != 3 || len(call.Arguments.ExpressionIDs) != 3 {
		t.Fatalf("recovered call arguments = %+v, want three positional slots", call.Arguments)
	}
	if call.Arguments.ExpressionIDs[0] == 0 || call.Arguments.ExpressionIDs[1] != 0 || call.Arguments.ExpressionIDs[2] == 0 {
		t.Fatalf("omitted argument slot was not preserved: %+v", call.Arguments)
	}
	for index, want := range []string{".first", "third"} {
		expressionID := call.Arguments.ExpressionIDs[[]int{0, 2}[index]]
		expression := doc.Procedures[1].Expressions[expressionID-1]
		if expression.Text != want || !expression.Recovered {
			t.Fatalf("recovered argument %d = %+v, want %q", index, expression, want)
		}
	}
}

func TestRecoveredCallArgumentsPreservesEmptyPositionalSegments(t *testing.T) {
	t.Parallel()
	source := []byte("first, , third")
	arguments := recoveredCallArguments(string(source), 0, len(source), 0, source)
	if len(arguments) != 3 {
		t.Fatalf("recovered arguments = %#v, want three positional slots", arguments)
	}
	if arguments[0].text != "first" || arguments[1].text != "" || arguments[2].text != "third" {
		t.Fatalf("recovered argument texts = %#v, want first/empty/third", arguments)
	}
	if arguments[1].rng.EndByte > arguments[1].rng.StartByte {
		t.Fatalf("omitted argument received a source range: %#v", arguments[1].rng)
	}
}

func TestPrintOutputListPositionsAreNotArguments(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Public Sub Run()
	Debug.Print first; second("value"), third;
	End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Procedures) != 1 || len(doc.Procedures[0].Calls) != 2 {
		t.Fatalf("Debug.Print call was not captured: %+v", doc.Procedures)
	}
	call := doc.Procedures[0].Calls[0]
	if call.Arguments.Count != 3 || len(call.Arguments.ExpressionIDs) != 3 {
		t.Fatalf("output positions must not become argument slots: %+v", call.Arguments)
	}
	for _, expressionID := range call.Arguments.ExpressionIDs {
		if expressionID == 0 {
			t.Fatalf("output expression lacks canonical expression link: %+v", call.Arguments)
		}
	}
	if nested := doc.Procedures[0].Calls[1]; !strings.EqualFold(nested.Callee.BaseName, "second") {
		t.Fatalf("nested output expression call was not captured: %+v", nested)
	}
}

func TestFunctionAndPropertyGetNamesResolveAsLocalReturnSlots(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Thing.cls", ModuleKind: "class"}, []byte(`Public Function Compute(ByVal input As Long) As Long
    Compute = input
End Function

Public Property Get Value() As Long
    Value = 2
End Property
`))
	if err != nil {
		t.Fatal(err)
	}
	declarations := doc.Procedures[0].Declarations
	if len(declarations) < 2 || declarations[0].Kind != "return_slot" || declarations[1].Kind != "parameter" ||
		declarations[0].Range.StartByte >= declarations[1].Range.StartByte {
		t.Fatalf("function declarations are not in source order: %+v", declarations)
	}
	resolved := Resolve(doc, NewResolver([]ResolverSymbol{
		{Name: "Compute", Module: "Other", Kind: "module_variable", Visibility: "Public"},
		{Name: "Value", Module: "Other", Kind: "module_variable", Visibility: "Public"},
	}))
	for _, procedure := range resolved.Procedures {
		assertAccess(t, procedure.Accesses, procedure.Symbol.Name, ScopeLocal, AccessWrite)
		for _, access := range procedure.Accesses {
			if strings.EqualFold(access.Name, procedure.Symbol.Name) && len(access.Resolution.Candidates) != 0 {
				t.Fatalf("return slot flowed to project resolver: %+v", access)
			}
		}
	}
}

func TestExplicitLetReturnSlotResolvesAsLocalWrite(t *testing.T) {
	doc, err := BuildSource(BuildOptions{Path: "Thing.cls", ModuleKind: "class"}, []byte(`Public Function Compute() As Long
    Let Compute = 2
End Function
`))
	if err != nil {
		t.Fatal(err)
	}
	assertAccess(t, doc.Procedures[0].Accesses, "Compute", ScopeLocal, AccessWrite)
}

func TestReDimAccessModes(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`Public Sub Resize()
    Dim values() As Long
    ReDim values(10)
    ReDim Preserve values(20)
End Sub
`))
	if err != nil {
		t.Fatal(err)
	}
	var modes []AccessMode
	for _, access := range doc.Procedures[0].Accesses {
		if strings.EqualFold(access.Name, "values") {
			modes = append(modes, access.Mode)
		}
	}
	if len(modes) != 2 || modes[0] != AccessWrite || modes[1] != AccessReadWrite {
		t.Fatalf("ReDim access modes = %v, want [write read_write]", modes)
	}
}

func TestConditionalProcedureRetainsCallerAndBodyCalls(t *testing.T) {
	t.Parallel()
	doc, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(`#If VBA7 Then
Public Function ConditionalRun() As Long
    ConditionalRun = TargetCall(1)
End Function
#End If
`))
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Procedures) != 1 {
		t.Fatalf("conditional procedures = %d: %+v", len(doc.Procedures), doc.Procedures)
	}
	procedure := doc.Procedures[0]
	if procedure.Symbol.Name != "ConditionalRun" || len(procedure.Calls) != 1 {
		t.Fatalf("conditional procedure facts lost: %+v", procedure)
	}
	if procedure.Calls[0].Caller.Name != "ConditionalRun" ||
		procedure.Calls[0].Caller.QualifiedName != "Module1.ConditionalRun" {
		t.Fatalf("conditional caller metadata = %+v", procedure.Calls[0].Caller)
	}
}

func TestCloneIsDeep(t *testing.T) {
	t.Parallel()
	source, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte("Public Sub Run(ByVal x As Long)\nCall Foo(x)\nEnd Sub\n"))
	if err != nil {
		t.Fatal(err)
	}
	clone := Clone(source)
	clone.Procedures[0].Symbol.Parameters[0].Name = "changed"
	clone.Procedures[0].Statements[0].ExpressionIDs = append(clone.Procedures[0].Statements[0].ExpressionIDs, 999)
	clone.Procedures[0].Calls[0].Callee.Text = "changed"
	if source.Procedures[0].Symbol.Parameters[0].Name == "changed" ||
		source.Procedures[0].Calls[0].Callee.Text == "changed" {
		t.Fatal("Clone shares mutable storage")
	}
}

func TestCloneAndRebasePreserveOwnedSignatureRanges(t *testing.T) {
	t.Parallel()
	source, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte("#If VBA7 Then\nPublic Sub Run(ByVal values(1 To 2) As Long, Optional label As String = \"x\")\nEnd Sub\n#End If\n"))
	if err != nil {
		t.Fatal(err)
	}
	if len(source.Procedures) != 1 || len(source.Procedures[0].Symbol.Parameters[0].ArrayBounds) != 1 {
		t.Fatalf("signature source = %#v", source.Procedures)
	}
	clone := Clone(source)
	clone.Procedures[0].Symbol.Parameters[0].ArrayBounds[0].Lower = "changed"
	clone.Procedures[0].Symbol.ConditionalBranches[0].Condition = "changed"
	if source.Procedures[0].Symbol.Parameters[0].ArrayBounds[0].Lower == "changed" || source.Procedures[0].Symbol.ConditionalBranches[0].Condition == "changed" {
		t.Fatal("Clone shares signature storage")
	}
	oldBase := clone.Procedures[0].Symbol.DeclarationRange
	newBase := oldBase
	newBase.StartLine++
	newBase.EndLine++
	newBase.StartByte += 10
	newBase.EndByte += 10
	rebased := RebaseProcedure(clone.Procedures[0], oldBase, newBase)
	param := rebased.Symbol.Parameters[0]
	if param.Range.StartByte <= clone.Procedures[0].Symbol.Parameters[0].Range.StartByte || param.ArrayBounds[0].Range.StartByte <= clone.Procedures[0].Symbol.Parameters[0].ArrayBounds[0].Range.StartByte {
		t.Fatalf("signature ranges were not rebased: before=%#v after=%#v", clone.Procedures[0].Symbol.Parameters[0], param)
	}
}

func TestCloneControlFlowMetadataIsDeep(t *testing.T) {
	t.Parallel()
	source, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte("Public Sub Run()\nGoTo Done\nDone:\nEnd Sub\n"))
	if err != nil {
		t.Fatal(err)
	}
	clone := Clone(source)
	clone.Procedures[0].Statements[0].Control.Target = "Changed"
	if source.Procedures[0].Statements[0].Control.Target != "Done" {
		t.Fatal("Clone shares control-flow metadata")
	}
}

func TestResumeNextControlMetadataNormalizesWhitespace(t *testing.T) {
	t.Parallel()
	source, err := BuildSource(BuildOptions{Path: "Module1.bas"}, []byte(
		"Public Sub Run()\nResume    Next\nResume\tNext\nEnd Sub\n",
	))
	if err != nil {
		t.Fatal(err)
	}
	var resumeNext int
	for _, statement := range source.Procedures[0].Statements {
		if statement.Kind == StatementResume && statement.Control != nil &&
			statement.Control.Transfer == TransferResumeNext {
			resumeNext++
		}
	}
	if resumeNext != 2 {
		t.Fatalf("Resume Next statements = %d, want 2; statements=%+v",
			resumeNext, source.Procedures[0].Statements)
	}
}

func TestBuildParsedClosedDocument(t *testing.T) {
	t.Parallel()
	doc, err := vbaast.ParseDocument("Module1.bas", []byte("Sub Run()\nEnd Sub\n"))
	if err != nil {
		t.Fatal(err)
	}
	doc.Close()
	_, err = BuildParsed(BuildOptions{}, doc)
	if !errors.Is(err, vbaast.ErrParsedDocumentClosed) {
		t.Fatalf("error = %v", err)
	}
}

func requireStatementKind(t *testing.T, statements []Statement, want StatementKind) {
	t.Helper()
	for _, statement := range statements {
		if statement.Kind == want {
			return
		}
	}
	t.Fatalf("missing statement kind %q in %#v", want, statements)
}

func assertControl(
	t *testing.T,
	statements []Statement,
	text string,
	branch BranchRole,
	loop LoopTest,
	transfer TransferKind,
	target string,
	flag string,
) {
	t.Helper()
	var found *Statement
	for i := range statements {
		if statementText := strings.TrimSpace(statements[i].Text); statementText == text {
			found = &statements[i]
			break
		}
	}
	if found == nil {
		bestWidth := -1
		for i := range statements {
			if strings.Contains(statements[i].Text, text) {
				width := len(statements[i].Text)
				if bestWidth < 0 || width < bestWidth {
					found = &statements[i]
					bestWidth = width
				}
			}
		}
	}
	if found == nil {
		t.Fatalf("missing statement containing %q in %+v", text, statements)
		return
	}
	if found.Control == nil {
		t.Fatalf("%q has no control metadata: %+v", text, *found)
	}
	if found.Control.Branch != branch || found.Control.Loop != loop ||
		found.Control.Transfer != transfer || found.Control.Target != target ||
		(flag == "case_else") != found.Control.CaseElse {
		t.Fatalf("%q control = %+v, want branch=%q loop=%q transfer=%q target=%q flag=%q",
			text, found.Control, branch, loop, transfer, target, flag)
	}
	if found.Control.Range != found.Range {
		t.Fatalf("%q control range = %+v, want statement range %+v", text, found.Control.Range, found.Range)
	}
}

func assertAccess(t *testing.T, accesses []VariableAccess, name string, scope SymbolScope, mode AccessMode) {
	t.Helper()
	for _, access := range accesses {
		if strings.EqualFold(access.Name, name) && access.Scope == scope && access.Mode == mode {
			return
		}
	}
	t.Fatalf("missing access %s/%s/%s in %#v", name, scope, mode, accesses)
}

func assertTypeReference(t *testing.T, refs []TypeReference, kind, target string) {
	t.Helper()
	for _, ref := range refs {
		if ref.Kind == kind && strings.EqualFold(ref.Target, target) {
			return
		}
	}
	t.Fatalf("missing type reference %s/%s in %#v", kind, target, refs)
}

func hasRecoveredStatement(statements []Statement) bool {
	for _, statement := range statements {
		if statement.Recovered {
			return true
		}
	}
	return false
}

func containsTreeSitterType(typ reflect.Type, visited map[reflect.Type]bool) bool {
	if typ == nil || visited[typ] {
		return false
	}
	visited[typ] = true
	treeSitterNode := reflect.TypeOf(tree_sitter.Node{})
	treeSitterTree := reflect.TypeOf(tree_sitter.Tree{})
	if typ == treeSitterNode || typ == treeSitterTree {
		return true
	}
	switch typ.Kind() {
	case reflect.Pointer, reflect.Slice, reflect.Array:
		return containsTreeSitterType(typ.Elem(), visited)
	case reflect.Map:
		return containsTreeSitterType(typ.Key(), visited) || containsTreeSitterType(typ.Elem(), visited)
	case reflect.Struct:
		for i := 0; i < typ.NumField(); i++ {
			if containsTreeSitterType(typ.Field(i).Type, visited) {
				return true
			}
		}
	}
	return false
}
