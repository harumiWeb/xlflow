package analyze

import (
	"fmt"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/sourceproject"
)

func runParameterPassingAnalysis(t *testing.T, cfg config.Config, modules map[string]string) []Finding {
	t.Helper()
	dir := t.TempDir()
	for name, source := range modules {
		writeModule(t, dir, name, source)
	}
	findings, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	return findings
}

func parameterPassingConfig() config.Config {
	cfg := config.Default()
	cfg.Analyze.DetectAssignedByValParameters = true
	cfg.Analyze.DetectByRefParametersCanBeByVal = true
	cfg.Analyze.DetectMisleadingPropertyValueByRef = true
	return cfg
}

func TestParameterPassingReportsWritesAndReadOnlyByRefParameters(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Main.bas": `Option Explicit
Public Sub Run(ByVal localCopy As Long, ByRef outputValue As Long, ByRef inputValue As Long)
    Dim observed As Long
    localCopy = 1
    outputValue = 2
    observed = inputValue
End Sub
`})
	assigned := findingsByCode(findings, "VBA274")
	if len(assigned) != 1 || !strings.Contains(assigned[0].Message, "localCopy") {
		t.Fatalf("VBA274 findings = %+v", assigned)
	}
	canByVal := findingsByCode(findings, "VBA275")
	if len(canByVal) != 1 || !strings.Contains(canByVal[0].Message, "inputValue") {
		t.Fatalf("VBA275 findings = %+v", canByVal)
	}
}

func TestParameterPassingPropagatesResolvedByRefWrites(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Main.bas": `Option Explicit
Private Sub Mutate(ByRef value As Long)
    value = 1
End Sub

Private Sub Forward(ByRef target As Long)
    Mutate target
End Sub

Private Sub Observe(ByVal value As Long)
    Debug.Print value
End Sub

Private Sub ForwardReadOnly(ByRef target As Long)
    Observe target
End Sub
`})
	for _, finding := range findingsByCode(findings, "VBA275") {
		if finding.Procedure == "Forward" {
			t.Fatalf("VBA275 reported transitively written parameter: %+v", finding)
		}
	}
	got := findingsByCode(findings, "VBA275")
	found := false
	for _, finding := range got {
		if finding.Procedure == "ForwardReadOnly" && strings.Contains(finding.Message, "target") {
			found = true
		}
	}
	if !found {
		t.Fatalf("VBA275 findings = %+v, want ForwardReadOnly.target", got)
	}
}

func TestParameterPassingTreatsUnknownByRefCallsConservatively(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Main.bas": `Option Explicit
Private Sub ForwardUnknown(ByRef target As Long)
    ExternalHelper target
End Sub
`})
	if got := findingsByCode(findings, "VBA275"); len(got) != 0 {
		t.Fatalf("VBA275 findings = %+v, want none for unknown call", got)
	}
}

func TestParameterPassingReportsPropertyValueByRefOnlyOnFinalParameter(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Widget.cls": `Option Explicit
Private stored As Long

Public Property Let Item(ByRef index As Long, ByRef newValue As Long)
    stored = newValue + index
End Property
`})
	got := findingsByCode(findings, "VBA276")
	if len(got) != 1 || !strings.Contains(got[0].Message, "newValue") {
		t.Fatalf("VBA276 findings = %+v", got)
	}
	if got[0].Column <= 0 || got[0].EndColumn <= got[0].Column {
		t.Fatalf("VBA276 range = %+v", got[0])
	}
}

func TestParameterPassingStyleRulesAreIndependent(t *testing.T) {
	modules := map[string]string{"Main.bas": `Option Explicit
Public Sub Run(implicitValue As Long, ByRef explicitValue As Long)
    implicitValue = 1
    explicitValue = 2
End Sub
`}
	implicitCfg := config.Default()
	implicitCfg.Analyze.DetectImplicitByRefParameters = true
	implicit := findingsByCode(runParameterPassingAnalysis(t, implicitCfg, modules), "VBA273")
	if len(implicit) != 1 || !strings.Contains(implicit[0].Message, "implicitValue") {
		t.Fatalf("VBA273 findings = %+v", implicit)
	}

	redundantCfg := config.Default()
	redundantCfg.Analyze.DetectRedundantByRefModifiers = true
	redundant := findingsByCode(runParameterPassingAnalysis(t, redundantCfg, modules), "VBA277")
	if len(redundant) != 1 || !strings.Contains(redundant[0].Message, "explicitValue") {
		t.Fatalf("VBA277 findings = %+v", redundant)
	}
}

func TestParameterPassingCanBeByValSkipsArrayParameters(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Main.bas": `Option Explicit
Public Sub Run(ByRef values() As Long)
    Debug.Print UBound(values)
End Sub
`})
	if got := findingsByCode(findings, "VBA275"); len(got) != 0 {
		t.Fatalf("VBA275 findings = %+v, want none for array parameter", got)
	}
}

func TestParameterPassingDistinguishesObjectMutationFromParameterReassignment(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Main.bas": `Option Explicit
Public Sub MutateMember(ByRef target As Object)
    target.Caption = "updated"
End Sub

Public Sub ReplaceObject(ByRef target As Object)
    Set target = New Collection
End Sub
`})
	got := findingsByCode(findings, "VBA275")
	if len(got) != 1 || got[0].Procedure != "MutateMember" {
		t.Fatalf("VBA275 findings = %+v, want only MutateMember.target", got)
	}
}

func TestParameterPassingTreatsUDTMemberAssignmentAsCallerVisibleMutation(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Main.bas": `Option Explicit
Private Type Pair
    Left As Long
    Right As Long
End Type

Private Sub Mutate(ByRef value As Pair)
    value.Left = 1
End Sub
`})
	if got := findingsByCode(findings, "VBA275"); len(got) != 0 {
		t.Fatalf("VBA275 UDT findings = %+v, want none", got)
	}
}

func TestParameterPassingTreatsUDTMemberByRefCallAsCallerVisibleMutation(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Main.bas": `Option Explicit
Private Type Pair
    Left As Long
End Type

Private Sub Replace(ByRef value As Long)
    value = 1
End Sub

Private Sub MutateUDT(ByRef pairValue As Pair)
    Replace pairValue.Left
End Sub

Private Sub MutateObjectMember(ByRef objectValue As Object)
    Replace objectValue.Tag
End Sub
`})
	got := findingsByCode(findings, "VBA275")
	foundObject := false
	for _, finding := range got {
		if finding.Procedure == "MutateUDT" {
			t.Fatalf("VBA275 reported UDT member passed through ByRef: %+v", finding)
		}
		if finding.Procedure == "MutateObjectMember" {
			foundObject = true
		}
	}
	if !foundObject {
		t.Fatalf("VBA275 findings = %+v, want MutateObjectMember.objectValue", got)
	}
}

func TestParameterPassingTreatsWithBlockMemberWriteAsCallerVisibleMutation(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Main.bas": `Option Explicit
Private Type Pair
    Left As Long
    Right As Long
End Type

Private Type Outer
    Inner As Pair
End Type

Private Sub MutateDirect(ByRef value As Pair)
    With value
        .Left = 1
    End With
End Sub

Private Sub MutateNested(ByRef value As Outer)
    With value.Inner
        .Left = 1
    End With
End Sub

Private Sub MutateChained(ByRef value As Outer)
    With value
        .Inner.Left = 1
    End With
End Sub

Private Sub MutateUnknown(ByRef value As SomeExternal)
    With value
        .Left = 1
    End With
End Sub

Private Sub MutateObject(ByRef value As Collection)
    With value
        .Add 1
    End With
End Sub
`})
	got := findingsByCode(findings, "VBA275")
	for _, finding := range got {
		if finding.Procedure != "MutateObject" {
			t.Fatalf("VBA275 reported With-block written parameter: %+v", finding)
		}
	}
	if len(got) != 1 {
		t.Fatalf("VBA275 findings = %+v, want only MutateObject.value", got)
	}
}

func TestParameterPassingTreatsWithBlockImplicitCallArgumentAsMutation(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Main.bas": `Option Explicit
Private Type Pair
    Left As Long
End Type

Private Sub Replace(ByRef value As Long)
    value = 9
End Sub

Private Sub MutateCallKeyword(ByRef value As Pair)
    With value
        Call Replace(.Left)
    End With
End Sub

Private Sub MutateSpacedCallee(ByRef value As Pair)
    With value
        Replace .Left
    End With
End Sub
`})
	if got := findingsByCode(findings, "VBA275"); len(got) != 0 {
		t.Fatalf("VBA275 findings = %+v, want none for implicit member call arguments", got)
	}
}

func TestParameterPassingTreatsFileStatementsAsParameterWrites(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Main.bas": `Option Explicit
Private Type Pair
    Left As Long
End Type

Private Sub ResetValue(ByRef value As Variant)
    Erase value
End Sub

Private Sub ReadInput(ByRef text As String)
    Input #1, text
End Sub

Private Sub ReadLine(ByRef text As String)
    Line Input #1, text
End Sub

Private Sub ReadRecord(ByRef value As Variant)
    Get #1, , value
End Sub

Private Sub ReadInputMulti(ByRef fileNumber As Long, ByRef first As String, ByRef second As String)
    Input #fileNumber, first, second
End Sub

Private Sub ReadInputMember(ByRef pairValue As Pair)
    Input #1, pairValue.Left
End Sub

Private Sub WriteRecord(ByRef value As Variant)
    Put #1, , value
End Sub
`})
	got := findingsByCode(findings, "VBA275")
	found := map[string]bool{}
	for _, finding := range got {
		found[finding.Procedure+":"+finding.Message] = true
	}
	for _, procedure := range []string{"ResetValue", "ReadInput", "ReadLine", "ReadRecord", "ReadInputMember"} {
		for _, finding := range got {
			if finding.Procedure == procedure {
				t.Fatalf("VBA275 reported %s parameter written by file statement: %+v", procedure, got)
			}
		}
	}
	wantFileNumber, wantSecond, wantPut := false, false, false
	for _, finding := range got {
		switch {
		case finding.Procedure == "ReadInputMulti" && strings.Contains(finding.Message, "fileNumber"):
			wantFileNumber = true
		case finding.Procedure == "WriteRecord" && strings.Contains(finding.Message, "value"):
			wantPut = true
		case finding.Procedure == "ReadInputMulti" && strings.Contains(finding.Message, "second"):
			wantSecond = true
		}
	}
	if !wantFileNumber || wantSecond || !wantPut {
		t.Fatalf("VBA275 findings = %+v, want only ReadInputMulti.fileNumber and WriteRecord.value", got)
	}
}

func TestParameterPassingSkipsParamArrayForStyleRules(t *testing.T) {
	modules := map[string]string{"Main.bas": `Option Explicit
Public Sub Run(ParamArray args())
    Debug.Print UBound(args)
End Sub
`}
	implicitCfg := config.Default()
	implicitCfg.Analyze.DetectImplicitByRefParameters = true
	if got := findingsByCode(runParameterPassingAnalysis(t, implicitCfg, modules), "VBA273"); len(got) != 0 {
		t.Fatalf("VBA273 findings = %+v, want none for ParamArray", got)
	}
}

func TestParameterPassingIndexedWriteIsNotBindingReplacement(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Main.bas": `Option Explicit
Private Sub WriteElement(ByVal value As Variant)
    value(0) = 1
End Sub

Private Sub WriteElementByRef(ByRef value As Variant)
    value(0) = 1
End Sub
`})
	if got := findingsByCode(findings, "VBA274"); len(got) != 0 {
		t.Fatalf("VBA274 findings = %+v, want none for element write", got)
	}
	if got := findingsByCode(findings, "VBA275"); len(got) != 0 {
		t.Fatalf("VBA275 findings = %+v, want none for caller-visible element write", got)
	}
}

func TestParameterPassingParenthesizedArgumentIsForcedByVal(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Main.bas": `Option Explicit
Private Sub Mutate(ByRef value As Long)
    value = 1
End Sub

Private Sub Forward(ByRef target As Long)
    Call Mutate((target))
End Sub
`})
	got := findingsByCode(findings, "VBA275")
	if len(got) != 1 || got[0].Procedure != "Forward" {
		t.Fatalf("VBA275 findings = %+v, want Forward.target for forced ByVal argument", got)
	}
}

func TestParameterPassingSkipsImplementsPropertyValueByRef(t *testing.T) {
	cfg := config.Default()
	cfg.Analyze.DetectMisleadingPropertyValueByRef = true
	result, err := (Analyzer{Config: cfg}).AnalyzeProject(t.Context(), sourceproject.SourceProject{Files: []sourceproject.SourceFile{
		{Path: "virtual/Worker.cls", ModuleKind: sourceproject.ModuleKindClass, Source: []byte(`Option Explicit
Implements IWorker
Private Property Let IWorker_Value(ByRef newValue As Long)
End Property
`)},
	}})
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(result.Findings, "VBA276"); len(got) != 0 {
		t.Fatalf("VBA276 constrained findings = %+v, want none", got)
	}
}

func TestParameterPassingPropagatesNamedByRefArguments(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Main.bas": `Option Explicit
Private Sub Mutate(ByRef value As Long)
    value = 1
End Sub

Private Sub Forward(ByRef target As Long)
    Mutate value:=target
End Sub
`})
	for _, finding := range findingsByCode(findings, "VBA275") {
		if finding.Procedure == "Forward" {
			t.Fatalf("VBA275 reported named transitively written parameter: %+v", finding)
		}
	}
}

func parameterFindingMentions(finding Finding, name string) bool {
	return strings.Contains(finding.Message, "parameter "+name+" ")
}

func TestParameterPassingPropagatesNamedByRefArgumentsRegardlessOfWidth(t *testing.T) {
	// A named argument binds its value expression: alpha:=x must not lose the
	// value to the wider label, and a:=x must not lose it to the label on a
	// width tie.
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Main.bas": `Option Explicit
Private Sub Mutate(ByRef a As Long, ByRef b As Long)
    a = 1
End Sub

Private Sub MutateLong(ByRef alpha As Long)
    alpha = 1
End Sub

Private Sub Forward(ByRef x As Long, ByRef y As Long)
    Mutate a:=x, b:=y
End Sub

Private Sub ForwardLong(ByRef z As Long)
    MutateLong alpha:=z
End Sub
`})
	got := findingsByCode(findings, "VBA275")
	for _, finding := range got {
		if finding.Procedure == "Forward" && parameterFindingMentions(finding, "x") {
			t.Fatalf("VBA275 reported parameter written through a:=x: %+v", finding)
		}
		if finding.Procedure == "ForwardLong" {
			t.Fatalf("VBA275 reported parameter written through alpha:=z: %+v", finding)
		}
	}
	found := false
	for _, finding := range got {
		if finding.Procedure == "Forward" && parameterFindingMentions(finding, "y") {
			found = true
		}
	}
	if !found {
		t.Fatalf("VBA275 findings = %+v, want Forward.y bound to never-written b", got)
	}
}

func TestParameterPassingPropagatesIndexedElementByRefArguments(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Main.bas": `Option Explicit
Private Sub Mutate(ByRef v As Long)
    v = 1
End Sub

Private Sub ForwardElem(ByRef arr As Variant)
    Mutate arr(0)
End Sub

Private Sub ForwardElemCall(ByRef items As Variant)
    Call Mutate(items(0))
End Sub

Private Sub SwapElem(ByRef v As Variant, ByRef i As Long, ByRef j As Long)
    Mutate v(i)
    Mutate v(j)
End Sub

Private Sub DoubleParenElem(ByRef arr As Variant)
    Call Mutate((arr(0)))
End Sub

Private Sub ForwardMemberElem(ByRef arr As Variant)
    Call Mutate(arr(0).Left)
End Sub
`})
	got := findingsByCode(findings, "VBA275")
	for _, finding := range got {
		switch finding.Procedure {
		case "ForwardElem", "ForwardElemCall", "ForwardMemberElem", "SwapElem":
			t.Fatalf("VBA275 reported element-forwarded parameter: %+v", finding)
		}
	}
	// Call Mutate((arr(0))) double-parenthesizes the element to a value: arr
	// is genuinely never written. SwapElem index reads inside v(i)/v(j) stay
	// possiblyWritten like reads inside any computed argument.
	want := map[string]bool{"DoubleParenElem.arr": false}
	for _, finding := range got {
		for param := range map[string]bool{"arr": true, "i": true, "j": true} {
			if parameterFindingMentions(finding, param) {
				want[finding.Procedure+"."+param] = true
			}
		}
	}
	for key, found := range want {
		if !found {
			t.Fatalf("VBA275 findings = %+v, want true positive %s", got, key)
		}
	}
}

func TestParameterPassingTreatsIndexedStatementOperandsAsWrites(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Main.bas": `Option Explicit
Private Type Pair
    Left As Long
End Type

Private Type Outer
    Items(0 To 4) As Pair
End Type

Private Sub WipeElem(ByRef arr As Variant)
    Erase arr(0)
End Sub

Private Sub ReadIntoElem(ByRef arr As Variant)
    Get #1, , arr(0)
End Sub

Private Sub WithElem(ByRef v As Variant)
    With v(0)
        .Left = 1
    End With
End Sub

Private Sub WithMemberElem(ByRef p As Outer)
    With p.Items(0)
        .Left = 1
    End With
End Sub
`})
	if got := findingsByCode(findings, "VBA275"); len(got) != 0 {
		t.Fatalf("VBA275 findings = %+v, want none for indexed statement operands", got)
	}
}

func TestParameterPassingTreatsBracketedAndShadowedTypeMemberWrites(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Main.bas": `Option Explicit
Private Type Pair
    Left As Long
End Type

Private Type Collection
    X As Long
End Type

Private Sub MutateBracket(ByRef [My Field] As Pair)
    [My Field].Left = 1
End Sub

Private Sub MutateShadow(ByRef p As Collection)
    p.X = 1
End Sub
`})
	if got := findingsByCode(findings, "VBA275"); len(got) != 0 {
		t.Fatalf("VBA275 findings = %+v, want none for bracketed/shadowed member writes", got)
	}
}

func TestParameterPassingTreatsShadowedUDTMemberByRefCallAsCallerVisibleMutation(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Main.bas": `Option Explicit
Private Type Collection
    X As Long
End Type

Private Sub Replace(ByRef value As Long)
    value = 1
End Sub

Private Sub Forward(ByRef p As Collection)
    Replace p.X
End Sub
`})
	// The UDT named Collection shadows the builtin object type, so p.X is a
	// caller-visible member write forwarded through Replace, not an object
	// member mutation that leaves the binding untouched.
	if got := findingsByCode(findings, "VBA275"); len(got) != 0 {
		t.Fatalf("VBA275 findings = %+v, want none for shadowed UDT member call", got)
	}
}

func TestParameterPassingKeepsSameNamedPropertyAccessorSummariesSeparate(t *testing.T) {
	findings := runParameterPassingAnalysis(t, parameterPassingConfig(), map[string]string{"Widget.cls": `Option Explicit
Private stored As Long

Public Property Get Item(ByRef index As Long) As Long
    Item = index
End Property

Public Property Let Item(ByRef index As Long, ByVal value As Long)
    index = 2
    value = index
End Property
`})
	// Both accessors share the qualified name Widget.Item; each must keep its
	// own mutation summary. The Let accessor writes index, so only the Get
	// accessor's index may be reported as can-be-ByVal.
	canByVal := findingsByCode(findings, "VBA275")
	if len(canByVal) != 1 || canByVal[0].Line >= 8 {
		t.Fatalf("VBA275 findings = %+v, want only the Property Get index", canByVal)
	}
	assigned := findingsByCode(findings, "VBA274")
	if len(assigned) != 1 || !strings.Contains(assigned[0].Message, "value") {
		t.Fatalf("VBA274 findings = %+v, want the Property Let value assignment", assigned)
	}
}

func TestParameterPassingSkipsConstrainedEventAndImplementsSignatures(t *testing.T) {
	cfg := config.Default()
	cfg.Analyze.DetectImplicitByRefParameters = true
	cfg.Analyze.DetectByRefParametersCanBeByVal = true
	result, err := (Analyzer{Config: cfg}).AnalyzeProject(t.Context(), sourceproject.SourceProject{Files: []sourceproject.SourceFile{
		{Path: "virtual/Sheet1.cls", ModuleKind: sourceproject.ModuleKindDocument, Source: []byte(`Option Explicit
Private Sub Worksheet_Change(Target As Range)
    Target.Calculate
End Sub
`)},
		{Path: "virtual/Worker.cls", ModuleKind: sourceproject.ModuleKindClass, Source: []byte(`Option Explicit
Implements IWorker
Private Sub IWorker_Run(value As Long)
    Dim observed As Long
    observed = value
End Sub
`)},
	}})
	if err != nil {
		t.Fatal(err)
	}
	if got := findingsByCode(result.Findings, "VBA273"); len(got) != 0 {
		t.Fatalf("VBA273 constrained findings = %+v, want none", got)
	}
	if got := findingsByCode(result.Findings, "VBA275"); len(got) != 0 {
		t.Fatalf("VBA275 constrained findings = %+v, want none", got)
	}
}

func TestParameterPassingDefaultsDisabled(t *testing.T) {
	findings := runParameterPassingAnalysis(t, config.Default(), map[string]string{"Main.bas": `Option Explicit
Public Sub Run(ByVal copy As Long, implicitValue As Long, ByRef explicitValue As Long)
    copy = 1
End Sub
`})
	for _, code := range []string{"VBA273", "VBA274", "VBA275", "VBA276", "VBA277"} {
		if got := findingsByCode(findings, code); len(got) != 0 {
			t.Fatalf("default %s findings = %+v, want none", code, got)
		}
	}
}

func TestParameterPassingBatchRealtimeAndInlineSuppression(t *testing.T) {
	dir := t.TempDir()
	source := []byte(`Option Explicit
Public Sub Run(ByVal copy As Long, ByRef inputValue As Long) ' xlflow:disable-line VBA274
    copy = 1
    Dim observed As Long
    observed = inputValue
End Sub
`)
	writeModule(t, dir, "Main.bas", string(source))
	cfg := parameterPassingConfig()
	batch, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	realtime, err := SourceRealtimeFindings(dir, filepath.Join(dir, "src", "modules", "Main.bas"), cfg, source)
	if err != nil {
		t.Fatal(err)
	}
	for _, findings := range [][]Finding{batch, realtime} {
		if got := findingsByCode(findings, "VBA274"); len(got) != 0 {
			t.Fatalf("suppressed VBA274 findings = %+v", got)
		}
		got := findingsByCode(findings, "VBA275")
		if len(got) != 1 || !strings.Contains(got[0].Message, "inputValue") {
			t.Fatalf("VBA275 findings = %+v", got)
		}
	}
}

func TestParameterPassingFilesystemAndInMemoryParity(t *testing.T) {
	dir := t.TempDir()
	source := []byte(`Option Explicit
Public Sub Run(ByVal copy As Long, ByRef inputValue As Long)
    copy = 1
    Dim observed As Long
    observed = inputValue
End Sub
`)
	writeModule(t, dir, "Main.bas", string(source))
	cfg := parameterPassingConfig()
	filesystem, err := (Analyzer{RootDir: dir, Config: cfg}).Run()
	if err != nil {
		t.Fatal(err)
	}
	inMemory, err := (Analyzer{Config: cfg}).AnalyzeProject(t.Context(), sourceproject.SourceProject{Files: []sourceproject.SourceFile{{
		Path:       "virtual/Main.bas",
		Source:     source,
		ModuleKind: sourceproject.ModuleKindStandard,
	}}})
	if err != nil {
		t.Fatal(err)
	}
	filesystemCodes := []string{}
	for _, finding := range filesystem {
		if strings.HasPrefix(finding.Code, "VBA27") {
			filesystemCodes = append(filesystemCodes, finding.Code+":"+finding.Procedure+":"+finding.Message)
		}
	}
	inMemoryCodes := []string{}
	for _, finding := range inMemory.Findings {
		if strings.HasPrefix(finding.Code, "VBA27") {
			inMemoryCodes = append(inMemoryCodes, finding.Code+":"+finding.Procedure+":"+finding.Message)
		}
	}
	if !reflect.DeepEqual(filesystemCodes, inMemoryCodes) {
		t.Fatalf("parameter findings differ:\nfilesystem: %#v\nin-memory: %#v", filesystemCodes, inMemoryCodes)
	}
}

func TestParameterMutationWorklistScalesWithLongCallChain(t *testing.T) {
	const procedureCount = 128
	var source strings.Builder
	source.WriteString("Option Explicit\n")
	for index := 0; index < procedureCount; index++ {
		fmt.Fprintf(&source, "Private Sub Step%03d(ByRef value As Long)\n", index)
		if index+1 < procedureCount {
			fmt.Fprintf(&source, "    Step%03d value\n", index+1)
		} else {
			source.WriteString("    value = 1\n")
		}
		source.WriteString("End Sub\n")
	}
	for index := 0; index < procedureCount; index++ {
		fmt.Fprintf(&source, "Private Sub Unrelated%03d(ByRef value As Long)\n    Dim observed As Long\n    observed = value\nEnd Sub\n", index)
	}
	file := parameterPassingParsedFile(t, source.String())
	summaries, work := buildParameterMutationSummariesWithWork([]parsedFile{file})
	if got := summaries["sub|main.step000"]["value"]; got != parameterWritten {
		t.Fatalf("Step000.value state = %v, want written", got)
	}
	if work.records != procedureCount*2 {
		t.Fatalf("records = %d, want %d", work.records, procedureCount*2)
	}
	if work.evaluations > work.records*2 {
		t.Fatalf("worklist evaluations = %d for %d records, want at most two per record", work.evaluations, work.records)
	}
}

func parameterPassingParsedFile(t *testing.T, source string) parsedFile {
	t.Helper()
	path := filepath.Join(t.TempDir(), "Main.bas")
	doc, err := vbaast.ParseDocument(path, []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	ir, err := procedureir.BuildParsed(procedureir.BuildOptions{Path: path, ModuleKind: "standard"}, doc)
	if err != nil {
		t.Fatal(err)
	}
	symbols := make([]procedureir.ResolverSymbol, 0, len(ir.Procedures))
	for _, procedure := range ir.Procedures {
		symbols = append(symbols, procedureir.ResolverSymbol{
			Name: procedure.Symbol.Name, Module: ir.ModuleName, ModuleKind: ir.ModuleKind,
			Kind: string(procedure.Symbol.Kind), Visibility: procedure.Symbol.Visibility,
			File: path, Line: procedure.Symbol.DeclarationRange.StartLine,
		})
	}
	ir = procedureir.Resolve(ir, procedureir.NewResolver(symbols))
	flow := vbacfg.BuildDocument(ir)
	return parsedFile{
		Path: path, Module: ir.ModuleName, ModuleKind: "standard", Source: []byte(source),
		Lines: normalizedSourceLines(source), IR: ir, CFG: flow,
		Procedures: sourceProceduresFromIRRef(&ir, flow),
	}
}
