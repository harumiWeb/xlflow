package analyze

import (
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestModuleAnalysisFactsUsesIRDeclarationsAndIndexedProcedureOwnership(t *testing.T) {
	lines := []string{
		"Dim moduleObject As Object",
		"Public Sub First()",
		"    Dim localValue As Long",
		"End Sub",
		"Private Const moduleConst As Long = 1",
	}
	procedures := []sourceProcedure{{StartLine: 2, EndLine: 4, StartByte: 10}}
	document := procedureir.DocumentIR{
		Declarations: []procedureir.Declaration{{
			Name: "moduleObject", Type: "Object", IsObject: true,
			Scope: procedureir.ScopeModule, Kind: "variable",
			Range: vbaast.Range{StartLine: 1, EndLine: 1},
		}, {
			Name: "newWorksheet", Type: " New Worksheet ", IsObject: true,
			Scope: procedureir.ScopeModule, Kind: "variable",
			Range: vbaast.Range{StartLine: 1, EndLine: 1},
		}, {
			Name: "table", Type: "ListObject",
			Scope: procedureir.ScopeModule, Kind: "variable",
			Range: vbaast.Range{StartLine: 1, EndLine: 1},
		}, {
			Name: "control", Type: "MSForms.Label", IsObject: true,
			Scope: procedureir.ScopeModule, Kind: "variable",
			Range: vbaast.Range{StartLine: 1, EndLine: 1},
		}, {
			Name: "customObject", Type: "TThis", IsObject: true,
			Scope: procedureir.ScopeModule, Kind: "variable",
			Range: vbaast.Range{StartLine: 1, EndLine: 1},
		}, {
			Name: "externalObject", Type: "mscorlib.AppDomain", IsObject: true,
			Scope: procedureir.ScopeModule, Kind: "variable",
			Range: vbaast.Range{StartLine: 1, EndLine: 1},
		}},
	}

	facts := buildModuleAnalysisFacts(lines, document, procedures)
	if _, ok := moduleDeclarations(lines, procedures)["moduleobject"]; !ok {
		t.Fatalf("compatibility module declarations omitted moduleObject")
	}
	if !lineInAnyProcedure(2, procedures) || lineInAnyProcedure(1, procedures) {
		t.Fatalf("compatibility procedure ownership helper disagrees with indexed facts")
	}
	if _, ok := facts.moduleDeclarations["moduleobject"]; !ok {
		t.Fatalf("module declarations = %#v, want moduleObject", facts.moduleDeclarations)
	}
	if declaration := facts.moduleDeclarations["newworksheet"]; declaration.Type != "Worksheet" || !declaration.NewExpression {
		t.Fatalf("normalized As New declaration = %#v, want Worksheet/NewExpression", declaration)
	}
	if declaration := facts.moduleDeclarations["table"]; !declaration.Object {
		t.Fatalf("IR ListObject declaration = %#v, want object declaration", declaration)
	}
	// MSForms controls are runtime-initialized and excluded, while other
	// IR-marked external objects such as mscorlib.AppDomain remain objects.
	if declaration := facts.moduleDeclarations["control"]; declaration.Object {
		t.Fatalf("dotted form-control declaration = %#v, want non-object declaration", declaration)
	}
	if declaration := facts.moduleDeclarations["customobject"]; !declaration.Object {
		t.Fatalf("unqualified custom declaration = %#v, want object declaration", declaration)
	}
	if declaration := facts.moduleDeclarations["externalobject"]; !declaration.Object {
		t.Fatalf("qualified external declaration = %#v, want object declaration", declaration)
	}
	if _, ok := facts.moduleDeclarations["localvalue"]; ok {
		t.Fatalf("procedure-local declaration leaked into module declarations: %#v", facts.moduleDeclarations)
	}
	if facts.lineInProcedure(1) || !facts.lineInProcedure(2) || !facts.lineInProcedure(4) || facts.lineInProcedure(5) {
		t.Fatalf("procedure line ownership = %#v, want only lines 2-4", facts.procedureLineOwners)
	}
	if got := facts.procedureDeclarations(procedures[0]); got["localvalue"].Name != "localValue" {
		t.Fatalf("cached procedure declarations = %#v, want localValue", got)
	}
	if got := facts.procedureFactsFor(procedures[0]); got == nil || got != facts.procedureFactsFor(procedures[0]) {
		t.Fatalf("procedure facts start-byte index = %p, want one stable facts pointer", got)
	}
}

func TestModuleAnalysisFactsFallsBackToSourceForIncompleteIR(t *testing.T) {
	lines := []string{
		"Private obj As Object",
		"Private Sub Run()",
		"    Dim localValue As Long",
		"End Sub",
	}
	procedures := []sourceProcedure{{StartLine: 2, EndLine: 4, StartByte: 10}}
	facts := buildModuleAnalysisFacts(lines, procedureir.DocumentIR{Parse: procedureir.ParseSummary{HasMissing: true}}, procedures)
	if declaration := facts.moduleDeclarations["obj"]; !declaration.Object {
		t.Fatalf("source fallback declaration = %#v, want object declaration", declaration)
	}
	if _, ok := facts.moduleDeclarations["localvalue"]; ok {
		t.Fatalf("source fallback included procedure-local declaration: %#v", facts.moduleDeclarations)
	}
}

func TestModuleAnalysisFactsIndexesOrderedConstantsAndLocalProcedures(t *testing.T) {
	lines := []string{
		"Private Const RootPath As String = ThisWorkbook.Path",
		"Public Sub First()",
		"    Const LocalPath As String = RootPath & \"\\data\"",
		"End Sub",
		"Private Function Second() As String",
		"    Second = LocalPath",
		"End Function",
	}
	procedures := []sourceProcedure{
		{Name: "First", StartLine: 2, EndLine: 4, StartByte: 101},
		{Name: "Second", StartLine: 5, EndLine: 7, StartByte: 201},
	}
	facts := buildModuleAnalysisFacts(lines, procedureir.DocumentIR{}, procedures)

	var constants []moduleConstantFact
	facts.forEachConstant(func(constant moduleConstantFact) { constants = append(constants, constant) })
	if len(constants) != 2 {
		t.Fatalf("constant facts = %#v, want two source declarations", constants)
	}
	if constants[0].Name != "RootPath" || constants[0].Line != 1 || !constants[0].Module {
		t.Fatalf("first constant fact = %#v, want module RootPath on line 1", constants[0])
	}
	if constants[1].Name != "LocalPath" || constants[1].Line != 3 || constants[1].Module {
		t.Fatalf("second constant fact = %#v, want procedure-local LocalPath on line 3", constants[1])
	}
	if !facts.hasConstant("rootpath") || !facts.hasConstant("LOCALPATH") || facts.hasConstant("Missing") {
		t.Fatalf("constant name index does not provide case-insensitive read-only lookup")
	}
	if !facts.hasProcedure("first") || !facts.hasProcedure("SECOND") || facts.hasProcedure("Missing") {
		t.Fatalf("procedure name index does not provide case-insensitive same-module lookup")
	}
}

func TestModuleAnalysisFactsConstantLookupRespectsProcedureScope(t *testing.T) {
	lines := []string{
		"Private Const ModuleValue As Long = 1",
		"Public Sub First()",
		"End Sub",
		"Public Sub Second()",
		"    Const LocalValue As Long = 2",
		"End Sub",
	}
	procedures := []sourceProcedure{
		{Name: "First", StartLine: 2, EndLine: 3, StartByte: 101},
		{Name: "Second", StartLine: 4, EndLine: 6, StartByte: 201},
	}
	facts := buildModuleAnalysisFacts(lines, procedureir.DocumentIR{}, procedures)
	if !facts.hasConstantForProcedure("ModuleValue", procedures[0]) {
		t.Fatalf("module constant should be visible in First")
	}
	if facts.hasConstantForProcedure("LocalValue", procedures[0]) {
		t.Fatalf("Second's local constant leaked into First")
	}
	if !facts.hasConstantForProcedure("LocalValue", procedures[1]) {
		t.Fatalf("Second's local constant should be visible in Second")
	}
	var visible []string
	facts.forEachConstantForProcedure(procedures[1], func(constant moduleConstantFact) {
		visible = append(visible, constant.Name)
	})
	if len(visible) != 2 || visible[0] != "ModuleValue" || visible[1] != "LocalValue" {
		t.Fatalf("visible constants = %#v, want module then current local", visible)
	}
}

func TestDeclarationScopeLayersWithoutModuleMapCopy(t *testing.T) {
	module := map[string]sourceDeclaration{
		"value":      {Name: "value", Type: "Object", Object: true},
		"moduleonly": {Name: "moduleOnly", Type: "Long"},
	}
	proc := sourceProcedure{
		StartLine: 1,
		EndLine:   1,
		Params:    newReadOnlySpan([]parameterInfo{{Name: "value", Type: "String"}}),
	}
	file := parsedFile{ModuleDeclarations: module, Lines: []string{""}}
	scope := newDeclarationScope(file, proc)
	if len(scope.module) == 0 || scope.module["value"].Type != "Object" {
		t.Fatalf("scope module layer = %#v, want original module map", scope.module)
	}
	if declaration, ok := scope.lookup("VALUE"); !ok || declaration.Type != "String" || !declaration.Parameter {
		t.Fatalf("parameter overlay = %#v, %v", declaration, ok)
	}
	module["moduleonly"] = sourceDeclaration{Name: "moduleOnly", Type: "String"}
	if declaration, ok := scope.lookup("moduleOnly"); !ok || declaration.Type != "String" {
		t.Fatalf("scope must retain module map reference: %#v, %v", declaration, ok)
	}
	seen := map[string]sourceDeclaration{}
	scope.forEach(func(key string, declaration sourceDeclaration) { seen[key] = declaration })
	if len(seen) != 2 || seen["value"].Type != "String" || seen["moduleonly"].Type != "String" {
		t.Fatalf("effective declaration scope = %#v, want parameter shadowing without duplication", seen)
	}
	arrayScope := newDeclarationScope(file, sourceProcedure{
		StartLine: 1,
		EndLine:   1,
		Params:    newReadOnlySpan([]parameterInfo{{Name: "moduleOnly", Type: "Long()"}}),
	})
	if !arrayScope.shadowsModule("moduleOnly") {
		t.Fatalf("parameter declaration should shadow module array names")
	}
}

func TestVBA212DeclarationScopeRetainsIRProcedureDeclarations(t *testing.T) {
	module := map[string]sourceDeclaration{"service": {Name: "service", Type: "ModuleService"}}
	proc := sourceProcedure{
		StartLine: 1,
		EndLine:   1,
		StartByte: 1,
		Declarations: newReadOnlySpan([]procedureir.Declaration{{
			Name: "service", Type: "LocalService", Scope: procedureir.ScopeLocal,
		}}),
	}
	scope := declarationScope{module: module}
	addVBA212ProcedureIRDeclarations(&scope, proc)
	if declaration, ok := scope.lookup("service"); !ok || declaration.Type != "LocalService" {
		t.Fatalf("IR procedure declaration = %#v, %v; want local shadowing module", declaration, ok)
	}
	legacy := vba212DeclarationIndexes([]string{"Sub Run()"}, []sourceProcedure{proc})
	if declaration, ok := legacy[proc.StartByte]["service"]; !ok || declaration.Type != "LocalService" {
		t.Fatalf("legacy VBA212 declaration index = %#v, %v; want local shadowing module", declaration, ok)
	}
}

var benchmarkModuleFactsSink *moduleAnalysisFacts

func BenchmarkModuleAnalysisFacts(b *testing.B) {
	for _, workload := range []struct {
		name         string
		procedures   int
		declarations int
	}{
		{name: "small", procedures: 8, declarations: 4},
		{name: "hundreds", procedures: 500, declarations: 32},
		{name: "thousands", procedures: 2000, declarations: 2000},
	} {
		workload := workload
		b.Run(workload.name, func(b *testing.B) {
			lines := make([]string, workload.procedures*3)
			procedures := make([]sourceProcedure, 0, workload.procedures)
			for index := 0; index < workload.procedures; index++ {
				start := index*3 + 1
				lines[start-1] = "Private Sub Run" + strconvItoa(index) + "()"
				lines[start] = "    Dim localValue As Long"
				lines[start+1] = "End Sub"
				procedures = append(procedures, sourceProcedure{StartLine: start, EndLine: start + 2, StartByte: index + 1})
			}
			declarations := make([]procedureir.Declaration, 0, workload.declarations)
			for index := 0; index < workload.declarations; index++ {
				declarations = append(declarations, procedureir.Declaration{
					Name: "ModuleValue" + strconvItoa(index), Type: "Long", Scope: procedureir.ScopeModule,
					Kind: "variable", Range: vbaast.Range{StartLine: index + 1, EndLine: index + 1},
				})
			}
			document := procedureir.DocumentIR{Declarations: declarations}
			b.ReportAllocs()
			b.ResetTimer()
			for index := 0; index < b.N; index++ {
				benchmarkModuleFactsSink = buildModuleAnalysisFacts(lines, document, procedures)
			}
			b.ReportMetric(float64(workload.procedures), "procedures/op")
			b.ReportMetric(float64(workload.declarations), "declarations/op")
		})
	}
}
