package analyze

import (
	"context"
	"reflect"
	"sort"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestBuildArrayParticipantSetExcludesUnrelatedProcedures(t *testing.T) {
	matched := func(callee, caller string) procedureir.CallSite {
		return procedureir.CallSite{
			Caller: procedureir.ProcedureRef{QualifiedName: caller},
			Callee: procedureir.Callee{BaseName: callee, Text: callee},
			Resolution: procedureir.CallResolution{
				Status:     procedureir.ResolutionMatched,
				Candidates: []procedureir.Candidate{{QualifiedName: "M." + callee}},
			},
		}
	}
	worker := sourceProcedure{Module: "M", Name: "ArrayWorker", Features: procedureFeatureSet{present: featureArray}}
	caller := sourceProcedure{Module: "M", Name: "Caller", Calls: newReadOnlySpan([]procedureir.CallSite{matched("ArrayWorker", "M.Caller")})}
	independent := sourceProcedure{Module: "M", Name: "Independent"}
	other := sourceProcedure{Module: "M", Name: "Other", Calls: newReadOnlySpan([]procedureir.CallSite{matched("Independent", "M.Other")})}
	file := parsedFile{Path: "M.bas", Module: "M", ModuleDeclarations: map[string]sourceDeclaration{}, Procedures: []sourceProcedure{worker, caller, independent, other}}
	got := buildArrayParticipantSet([]parsedFile{file}, analysisContext{})
	want := map[string]bool{"m.arrayworker": true, "m.caller": true}
	if len(got) != len(want) {
		t.Fatalf("participants = %#v, want %#v", got, want)
	}
	for key := range want {
		if !got[key] {
			t.Errorf("participant %q missing from %#v", key, got)
		}
	}
	for key := range got {
		if !want[key] {
			t.Errorf("unrelated participant %q in %#v", key, got)
		}
	}
}

func TestBuildArrayParticipantSetIncludesTransitiveResolvedCallers(t *testing.T) {
	matched := func(callee, caller string) procedureir.CallSite {
		return procedureir.CallSite{
			Caller: procedureir.ProcedureRef{QualifiedName: caller},
			Callee: procedureir.Callee{BaseName: callee, Text: callee},
			Resolution: procedureir.CallResolution{
				Status:     procedureir.ResolutionMatched,
				Candidates: []procedureir.Candidate{{QualifiedName: "M." + callee}},
			},
		}
	}
	worker := sourceProcedure{Module: "M", Name: "ArrayWorker", Features: procedureFeatureSet{present: featureArray}, Calls: newReadOnlySpan([]procedureir.CallSite{matched("ScalarHelper", "M.ArrayWorker")})}
	wrapper := sourceProcedure{Module: "M", Name: "Wrapper", Calls: newReadOnlySpan([]procedureir.CallSite{matched("ArrayWorker", "M.Wrapper")})}
	top := sourceProcedure{Module: "M", Name: "Top", Calls: newReadOnlySpan([]procedureir.CallSite{matched("Wrapper", "M.Top")})}
	scalarHelper := sourceProcedure{Module: "M", Name: "ScalarHelper"}
	file := parsedFile{Path: "M.bas", Module: "M", ModuleDeclarations: map[string]sourceDeclaration{}, Procedures: []sourceProcedure{worker, wrapper, top, scalarHelper}}

	got := buildArrayParticipantSet([]parsedFile{file}, analysisContext{})
	for _, key := range []string{"m.arrayworker", "m.wrapper", "m.top"} {
		if !got[key] {
			t.Errorf("transitive participant %q missing from %#v", key, got)
		}
	}
	if got["m.scalarhelper"] {
		t.Fatalf("scalar callee entered participant closure: %#v", got)
	}
}

func TestBuildArrayParticipantSetIncludesRecursiveSCCAndModuleFallback(t *testing.T) {
	matched := func(callee, caller string) procedureir.CallSite {
		return procedureir.CallSite{
			Caller:     procedureir.ProcedureRef{QualifiedName: caller},
			Callee:     procedureir.Callee{BaseName: callee, Text: callee},
			Resolution: procedureir.CallResolution{Status: procedureir.ResolutionMatched, Candidates: []procedureir.Candidate{{QualifiedName: "M." + callee}}},
		}
	}
	a := sourceProcedure{Module: "M", Name: "A", Features: procedureFeatureSet{present: featureArray}}
	b := sourceProcedure{Module: "M", Name: "B", Calls: newReadOnlySpan([]procedureir.CallSite{matched("A", "M.B")})}
	a.Calls = newReadOnlySpan([]procedureir.CallSite{matched("B", "M.A")})
	unknown := sourceProcedure{Module: "M", Name: "Unknown", Features: procedureFeatureSet{present: featureArray}, Calls: newReadOnlySpan([]procedureir.CallSite{{
		Caller:     procedureir.ProcedureRef{QualifiedName: "M.Unknown"},
		Callee:     procedureir.Callee{BaseName: "Run", Text: "Application.Run"},
		Resolution: procedureir.CallResolution{Status: procedureir.ResolutionDynamic},
	}})}
	independent := sourceProcedure{Module: "M", Name: "Independent"}
	file := parsedFile{Path: "M.bas", Module: "M", ModuleDeclarations: map[string]sourceDeclaration{}, Procedures: []sourceProcedure{a, b, unknown, independent}}
	got := buildArrayParticipantSet([]parsedFile{file}, analysisContext{})
	for _, key := range []string{"m.a", "m.b", "m.unknown", "m.independent"} {
		if !got[key] {
			t.Errorf("participant %q missing from recursive/fallback closure %#v", key, got)
		}
	}
}

func TestArrayParticipantWorklistOrderIsDeterministic(t *testing.T) {
	procedures := []sourceProcedure{
		{Module: "M", Name: "B", StartLine: 30},
		{Module: "M", Name: "A", StartLine: 20},
		{Module: "M", Name: "C", StartLine: 10},
	}
	inputs := [][]sourceProcedure{
		procedures,
		{procedures[2], procedures[0], procedures[1]},
		{procedures[1], procedures[2], procedures[0]},
	}
	var want []string
	for index, input := range inputs {
		ordered := append([]sourceProcedure(nil), input...)
		sort.SliceStable(ordered, func(i, j int) bool {
			return arrayProcedureLess(ordered[i], ordered[j])
		})
		got := make([]string, 0, len(ordered))
		for _, procedure := range ordered {
			got = append(got, arrayProcedureKey(procedure))
		}
		if index == 0 {
			want = got
			continue
		}
		if !reflect.DeepEqual(got, want) {
			t.Fatalf("worklist order changed with input order: got %v want %v", got, want)
		}
	}
}

func TestArrayReturnSummaryDuplicateNamesRemainUnknown(t *testing.T) {
	got := arrayReturnSummaryDuplicateNames([]sourceProcedure{
		{Module: "First", Name: "BuildValues"},
		{Module: "Second", Name: "buildvalues"},
		{Module: "First", Name: "Unique"},
	})
	if !got["buildvalues"] {
		t.Fatalf("duplicate return name was not marked ambiguous: %#v", got)
	}
	if got["unique"] {
		t.Fatalf("unique return name was marked ambiguous: %#v", got)
	}
}

func TestArrayReturnValueCompatibleIgnoresShape(t *testing.T) {
	left := arrayValue{
		kind:       arrayAllocated,
		knownArray: true,
		origin:     arrayOriginLocal,
		dimensions: []arrayDimension{{}},
	}
	right := arrayValue{
		kind:       arrayAllocated,
		knownArray: true,
		origin:     arrayOriginLocal,
		dimensions: []arrayDimension{{lower: arrayBound{known: true, value: 1}, upper: arrayBound{known: true, value: 2}}},
	}
	if !arrayReturnValueCompatible(sourceProcedure{ProcedureKind: procedureir.ProcedurePropertyGet}, left, right) {
		t.Fatalf("array return values with different shapes should retain allocation compatibility: left=%#v right=%#v", left, right)
	}
}

func TestInlineArrayReturnAssignmentTextUsesArraySummary(t *testing.T) {
	text, ok := inlineArrayReturnAssignmentText("Dim values() As Variant: values = source.arr", map[string]arrayValue{
		"arr": {kind: arrayAllocated, knownArray: true, origin: arrayOriginLocal},
	})
	if !ok || text != "values = source.arr" {
		t.Fatalf("known array-return member assignment was not extracted: text=%q ok=%v", text, ok)
	}
}

func TestInlineArrayQualifiedReturnAssignmentUsesTypeNameCase(t *testing.T) {
	file := parsedFile{Lines: []string{
		"Select Case TypeName(vFibers)",
		"    Case \"stdEnumerator\"",
		"        Dim queue() As Object: queue = vFibers.AsArray(vbObject)",
	}}
	proc := sourceProcedure{Module: "M", Name: "Run", StartLine: 1, EndLine: 3}
	text, ok := inlineArrayQualifiedReturnAssignmentText(file, proc, 3, file.Lines[2], map[string]arrayValue{
		"stdenumerator.asarray": {kind: arrayAllocated, knownArray: true, origin: arrayOriginLocal},
	})
	if !ok || text != "queue = Array()" {
		t.Fatalf("typed array-return member assignment was not normalized: text=%q ok=%v", text, ok)
	}
}

func TestInferDocumentedArrayReturnSummariesRequiresArrayImplementation(t *testing.T) {
	proc := sourceProcedure{Module: "M", Name: "AsArray", ProcedureKind: procedureir.ProcedureFunction, StartLine: 2, EndLine: 5}
	file := parsedFile{
		Lines: []string{
			"' @returns Array<T>",
			"Public Function AsArray() As Variant",
			"    ReDim values(1 To 2)",
			"    AsArray = values",
			"End Function",
		},
		Procedures: []sourceProcedure{proc},
	}
	if got := inferDocumentedArrayReturnSummaries([]parsedFile{file}); !got["m.asarray"].knownArray {
		t.Fatalf("documented allocated array return was not recognized: %#v", got)
	}
	proc.StartLine = 1
	file.Lines[0] = "Public Function AsArray() As Variant"
	if got := inferDocumentedArrayReturnSummaries([]parsedFile{file}); len(got) != 0 {
		t.Fatalf("undocumented or non-array return was recognized: %#v", got)
	}
}

func TestArrayCandidateKeyUsesProcedureKindForLineFallback(t *testing.T) {
	procedure := sourceProcedure{
		Module:        "M",
		Name:          "Value",
		ProcedureKind: procedureir.ProcedurePropertyGet,
		StartLine:     12,
	}
	key := arrayProcedureKey(procedure)
	all := map[string]sourceProcedure{key: procedure}
	index := buildArrayCandidateIndex(all)

	got := arrayCandidateKey(procedureir.Candidate{
		QualifiedName: "Missing.Value",
		Kind:          string(procedureir.ProcedurePropertyGet),
		Line:          12,
	}, all, index)
	if got != key {
		t.Fatalf("line/kind candidate = %q, want %q", got, key)
	}
	if got := arrayCandidateKey(procedureir.Candidate{
		QualifiedName: "Missing.Value",
		Kind:          string(procedureir.ProcedureFunction),
		Line:          12,
	}, all, index); got != "" {
		t.Fatalf("different candidate kind matched property getter: %q", got)
	}
}

func TestInferArrayReturnSummariesKeepsUnselectedDuplicateUnknown(t *testing.T) {
	makeFile := func(module, path string) parsedFile {
		source := []byte("Public Function BuildValues() As Variant()\n" +
			"    ReDim BuildValues(0 To 1)\n" +
			"End Function\n")
		ir, err := procedureir.BuildSource(procedureir.BuildOptions{
			Path: path, ModuleName: module, ModuleKind: "standard",
		}, source)
		if err != nil {
			t.Fatalf("build %s: %v", module, err)
		}
		flow := cfg.BuildDocument(ir)
		file := parsedFile{
			Path: path, Module: module, Source: source,
			Lines: normalizedSourceLines(string(source)), IR: ir, CFG: flow,
		}
		file.Procedures = sourceProceduresFromIRRef(&file.IR, flow)
		return file
	}
	files := []parsedFile{
		makeFile("First", "First.bas"),
		makeFile("Second", "Second.bas"),
	}
	ctx := analysisContext{arrayInterproceduralParticipants: map[string]bool{"first.buildvalues": true}}
	if summaries := inferArrayReturnSummaries(files, nil, ctx); summaries["buildvalues"].knownArray {
		t.Fatalf("duplicate bare-name summary became definite after participant filtering: %#v", summaries)
	}
}

func TestArrayParticipantGraphRetainsSameNamedPropertyAccessors(t *testing.T) {
	getter := sourceProcedure{
		Module:        "M",
		Name:          "Value",
		ProcedureKind: procedureir.ProcedurePropertyGet,
		Features:      procedureFeatureSet{present: featureArray},
	}
	setter := sourceProcedure{
		Module:        "M",
		Name:          "Value",
		ProcedureKind: procedureir.ProcedurePropertyLet,
	}
	file := parsedFile{
		Path: "M.cls", Module: "M", ModuleDeclarations: map[string]sourceDeclaration{},
		Procedures: []sourceProcedure{getter, setter},
	}
	participants, _, keys := buildArrayParticipantSets([]parsedFile{file}, analysisContext{})
	getterKey := keys[arrayParticipantProcedureIdentity(getter)]
	setterKey := keys[arrayParticipantProcedureIdentity(setter)]
	if getterKey == "" || setterKey == "" || getterKey == setterKey {
		t.Fatalf("same-named accessors did not receive stable distinct keys: getter=%q setter=%q keys=%v", getterKey, setterKey, keys)
	}
	if !participants[getterKey] {
		t.Fatalf("array property getter was excluded: participants=%v", participants)
	}
	if participants[setterKey] {
		t.Fatalf("unrelated property setter entered participant closure: participants=%v", participants)
	}
}

func TestArrayParticipantPlanExcludesModuleOnlyScalarProcedure(t *testing.T) {
	moduleDecls := map[string]sourceDeclaration{"values": {Name: "values", Array: true}}
	proc := sourceProcedure{Module: "M", Name: "Unrelated", Facts: &procedureAnalysisFacts{}}
	proc.ArrayParticipantReady = true
	proc.ArrayParticipant = false
	plan := buildProcedureAnalysisPlanWithModuleFeatures(analyzeConfigForRules("VBA227"), proc, moduleDeclarationFeatures(moduleDecls))
	if plan.runs(procedureDomainArray) {
		t.Fatalf("module-only scalar procedure retained array plan: %#v", plan)
	}
	proc.ArrayParticipant = true
	plan = buildProcedureAnalysisPlanWithModuleFeatures(analyzeConfigForRules("VBA227"), proc, moduleDeclarationFeatures(moduleDecls))
	if !plan.runs(procedureDomainArray) {
		t.Fatalf("participant procedure lost array plan: %#v", plan)
	}
}

func TestArrayParticipantPlanRestrictionPreservesExistingPlan(t *testing.T) {
	proc := sourceProcedure{Module: "M", Name: "Unrelated", PlanReady: true, Plan: procedureAnalysisPlan{
		enabled:            procedureDomainBit(procedureDomainArray),
		planned:            procedureDomainBit(procedureDomainArray),
		plannedKernels:     procedureKernelBit(procedureKernelArray),
		plannedProjections: procedureProjectionBit(procedureProjectionArrayLifecycle),
	}}
	got := restrictArrayProcedurePlan(proc.Plan, false)
	if got.enabled != proc.Plan.enabled {
		t.Fatalf("restriction changed enabled domain mask: got %#v want %#v", got, proc.Plan)
	}
	if got.runs(procedureDomainArray) || got.runsKernel(procedureKernelArray) || got.runsProjection(procedureProjectionArrayLifecycle) {
		t.Fatalf("unrelated procedure retained array work: %#v", got)
	}
}

func TestBuildArrayParticipantSetSeedsModuleArrayAccess(t *testing.T) {
	worker := sourceProcedure{Module: "M", Name: "ReadModuleArray", Accesses: newReadOnlySpan([]procedureir.VariableAccess{{
		Name: "values", Scope: procedureir.ScopeModule, Mode: procedureir.AccessRead,
	}})}
	independent := sourceProcedure{Module: "M", Name: "Scalar"}
	file := parsedFile{
		Path: "M.bas", Module: "M",
		ModuleDeclarations: map[string]sourceDeclaration{"values": {Name: "values", Array: true}},
		Procedures:         []sourceProcedure{worker, independent},
	}
	got := buildArrayParticipantSet([]parsedFile{file}, analysisContext{})
	if !got["m.readmodulearray"] || got["m.scalar"] {
		t.Fatalf("module-array access seeds = %#v, want only reader", got)
	}
}

func TestBuildArrayParticipantSetFailsOpenForUnknownOnlyProcedure(t *testing.T) {
	recovered := sourceProcedure{
		Module:   "M",
		Name:     "Recovered",
		Features: procedureFeatureSet{unknown: featureArray},
	}
	completeUnknown := sourceProcedure{
		Module:   "M",
		Name:     "CompleteUnknown",
		Document: &procedureir.DocumentIR{},
		IR:       &procedureir.ProcedureIR{},
		Graph:    &cfg.Graph{},
		Features: procedureFeatureSet{unknown: featureArray},
	}
	scalar := sourceProcedure{Module: "M", Name: "Scalar"}
	file := parsedFile{
		Path: "M.bas", Module: "M",
		ModuleDeclarations: map[string]sourceDeclaration{
			"values": {Name: "values", Array: true},
		},
		Procedures: []sourceProcedure{recovered, completeUnknown, scalar},
	}
	got := buildArrayParticipantSet([]parsedFile{file}, analysisContext{})
	if !got["m.recovered"] {
		t.Fatalf("unknown-only procedure was excluded from fail-open closure: %#v", got)
	}
	if !got["m.completeunknown"] {
		t.Fatalf("complete IR with unknown array capability was excluded from fail-open closure: %#v", got)
	}
	if got["m.scalar"] {
		t.Fatalf("module-only scalar procedure entered unknown-only closure: %#v", got)
	}
}

func TestArrayInterproceduralBoundaryAdmitsConnectedUnknownArrayOperation(t *testing.T) {
	matched := func(callee, caller string) procedureir.CallSite {
		return procedureir.CallSite{
			Caller: procedureir.ProcedureRef{QualifiedName: caller},
			Callee: procedureir.Callee{BaseName: callee, Text: callee},
			Resolution: procedureir.CallResolution{
				Status:     procedureir.ResolutionMatched,
				Candidates: []procedureir.Candidate{{QualifiedName: "M." + callee}},
			},
		}
	}
	worker := sourceProcedure{
		Module:   "M",
		Name:     "ArrayWorker",
		Features: procedureFeatureSet{present: featureArray},
		Calls:    newReadOnlySpan([]procedureir.CallSite{matched("UnknownArray", "M.ArrayWorker")}),
	}
	unknownArray := sourceProcedure{
		Module:     "M",
		Name:       "UnknownArray",
		Document:   &procedureir.DocumentIR{},
		IR:         &procedureir.ProcedureIR{},
		Graph:      &cfg.Graph{},
		Features:   procedureFeatureSet{unknown: featureArray},
		Statements: newReadOnlySpan([]procedureir.Statement{{Text: "values(1) = 0"}}),
	}
	unrelatedUnknown := sourceProcedure{
		Module:   "M",
		Name:     "UnrelatedUnknown",
		Document: &procedureir.DocumentIR{},
		IR:       &procedureir.ProcedureIR{},
		Graph:    &cfg.Graph{},
		Features: procedureFeatureSet{unknown: featureArray},
	}
	file := parsedFile{
		Path: "M.bas", Module: "M",
		ModuleDeclarations: map[string]sourceDeclaration{"values": {Name: "values", Array: true}},
		Procedures:         []sourceProcedure{worker, unknownArray, unrelatedUnknown},
	}
	participants := buildArrayParticipantSet([]parsedFile{file}, analysisContext{})
	interprocedural := buildArrayInterproceduralParticipantSet([]parsedFile{file}, analysisContext{}, participants)
	if !participants["m.unknownarray"] || !participants["m.unrelatedunknown"] {
		t.Fatalf("unknown local participants = %#v", participants)
	}
	if !interprocedural["m.unknownarray"] {
		t.Fatalf("connected unknown array operation was excluded from fixed-point boundary: %#v", interprocedural)
	}
	if interprocedural["m.unrelatedunknown"] {
		t.Fatalf("unrelated unknown procedure widened fixed-point boundary: %#v", interprocedural)
	}
}

func TestCloneArrayNameSetPreservesUnknownEntryState(t *testing.T) {
	input := map[string]bool{"allocated": true, "unknown": false}
	got := cloneArrayNameSet(input)
	if !reflect.DeepEqual(got, input) {
		t.Fatalf("cloned module entry state changed evidence: got %#v want %#v", got, input)
	}
}

func TestArrayModuleEntryStateDoesNotApplyUnallocatedEvidence(t *testing.T) {
	file := parsedFile{Path: "M.bas", Module: "M", Lines: []string{"Sub Helper", "End Sub"}}
	proc := sourceProcedure{Module: "M", Name: "Helper", StartLine: 1, EndLine: 2}
	variables := map[string]arrayVariable{
		"values": {name: "values", isArray: true},
	}
	initial := arrayFlowState{"values": {kind: arrayUnknown, knownArray: true}}
	got := applyArrayModuleEntryState(initial, file, proc, variables, map[string]sourceDeclaration{
		"values": {Name: "values", Array: true},
	}, arrayModuleEntryStates{
		"m.helper": {"values": false},
	})
	if got["values"].kind == arrayAllocated {
		t.Fatalf("unallocated module-entry evidence was applied: %#v", got["values"])
	}
}

func TestRealtimeProcedureProjectionPreservesEmptyIRFallback(t *testing.T) {
	fallback := []sourceProcedure{{StartLine: 1, EndLine: 4, StartByte: 0, EndByte: 20}}
	if got := rebindRealtimeProcedureProjection(fallback, nil); len(got) != 1 || got[0].StartLine != fallback[0].StartLine || got[0].EndLine != fallback[0].EndLine {
		t.Fatalf("empty materialized realtime projection replaced fallback: got %#v want %#v", got, fallback)
	}
	materialized := []sourceProcedure{{Name: "ArrayWorker", StartLine: 2, EndLine: 3}}
	if got := rebindRealtimeProcedureProjection(fallback, materialized); len(got) != 1 || got[0].Name != materialized[0].Name || got[0].StartLine != materialized[0].StartLine {
		t.Fatalf("materialized realtime projection was not rebound: got %#v want %#v", got, materialized)
	}
}

func TestBuildArrayParticipantSetBoundsUncertainty(t *testing.T) {
	matched := func(name string) procedureir.Candidate {
		return procedureir.Candidate{QualifiedName: "M." + name}
	}
	arraySeed := sourceProcedure{Module: "M", Name: "ArraySeed", Features: procedureFeatureSet{present: featureArray}}
	ambiguous := sourceProcedure{Module: "M", Name: "Ambiguous", Calls: newReadOnlySpan([]procedureir.CallSite{{
		Callee:     procedureir.Callee{BaseName: "Pick", Text: "Pick"},
		Resolution: procedureir.CallResolution{Status: procedureir.ResolutionAmbiguous, Candidates: []procedureir.Candidate{matched("ArraySeed"), matched("Candidate")}},
	}})}
	candidate := sourceProcedure{Module: "M", Name: "Candidate"}
	unresolved := sourceProcedure{Module: "M", Name: "Unresolved", Features: procedureFeatureSet{present: featureArray}, Calls: newReadOnlySpan([]procedureir.CallSite{{
		Callee:     procedureir.Callee{BaseName: "Run", Text: "Application.Run"},
		Resolution: procedureir.CallResolution{Status: procedureir.ResolutionDynamic},
	}})}
	unrelated := sourceProcedure{Module: "M", Name: "Unrelated"}
	external := sourceProcedure{Module: "N", Name: "ExternalOnly", Features: procedureFeatureSet{present: featureArray}, Calls: newReadOnlySpan([]procedureir.CallSite{{
		Callee:     procedureir.Callee{BaseName: "Open", Text: "obj.Open"},
		Resolution: procedureir.CallResolution{Status: procedureir.ResolutionMemberCall},
	}})}
	externalUnrelated := sourceProcedure{Module: "N", Name: "Unrelated"}
	files := []parsedFile{
		{Path: "M.bas", Module: "M", ModuleDeclarations: map[string]sourceDeclaration{}, Procedures: []sourceProcedure{arraySeed, ambiguous, candidate, unresolved, unrelated}},
		{Path: "N.bas", Module: "N", ModuleDeclarations: map[string]sourceDeclaration{}, Procedures: []sourceProcedure{external, externalUnrelated}},
	}
	got := buildArrayParticipantSet(files, analysisContext{})
	for _, key := range []string{"m.arrayseed", "m.ambiguous", "m.candidate", "m.unresolved", "m.unrelated", "n.externalonly"} {
		if !got[key] {
			t.Errorf("bounded participant %q missing from %#v", key, got)
		}
	}
	if got["n.unrelated"] {
		t.Fatalf("external/member call expanded unrelated project procedure: %#v", got)
	}
}

func TestArrayParticipantSyntheticTelemetryExcludesScalarProcedures(t *testing.T) {
	root := t.TempDir()
	fixture := writeSingleModuleBenchmarkProject(t, root, singleModuleBenchmarkWorkload{shape: "array-chain", size: 2000})
	t.Setenv(typedb.EnvDir, t.TempDir())
	recorder := analysisstats.NewRecorder()
	_, err := (Analyzer{RootDir: root, Config: config.Default()}).RunResultContext(analysisstats.WithRecorder(context.Background(), recorder))
	if err != nil {
		t.Fatal(err)
	}
	_, counters := recorder.Totals()
	values := make(map[string]uint64, len(counters))
	for _, counter := range counters {
		values[counter.Name] = counter.Value
	}
	for _, name := range []string{
		analysisstats.ArrayParticipantProceduresCounter,
		analysisstats.ArrayCandidateProceduresCounter,
		analysisstats.ArrayInterproceduralCFGWalksCounter,
	} {
		if _, ok := values[name]; !ok {
			t.Fatalf("telemetry counter %q is missing: %v", name, values)
		}
	}
	if values[analysisstats.ArrayParticipantProceduresCounter] == 0 {
		t.Fatalf("array participant counter is zero for array fixture: %v", values)
	}
	if values[analysisstats.ArrayParticipantProceduresCounter] >= uint64(fixture.procedures/10) {
		t.Fatalf("array participants = %d for %d-procedure fixture, want a small dependency closure; counters=%v", values[analysisstats.ArrayParticipantProceduresCounter], fixture.procedures, values)
	}
	if values[analysisstats.ArrayCandidateProceduresCounter] >= uint64(fixture.procedures/10) {
		t.Fatalf("array candidates = %d for %d-procedure fixture, want unrelated scalar procedures excluded; counters=%v", values[analysisstats.ArrayCandidateProceduresCounter], fixture.procedures, values)
	}
	if values[analysisstats.ArrayInterproceduralCFGWalksCounter] >= uint64(fixture.procedures/10) {
		t.Fatalf("interprocedural CFG walks = %d for %d-procedure fixture, want bounded participant work; counters=%v", values[analysisstats.ArrayInterproceduralCFGWalksCounter], fixture.procedures, values)
	}
}
