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

func TestCloneArrayNameSetPreservesUnknownEntryState(t *testing.T) {
	input := map[string]bool{"allocated": true, "unknown": false}
	got := cloneArrayNameSet(input)
	if !reflect.DeepEqual(got, input) {
		t.Fatalf("cloned module entry state changed evidence: got %#v want %#v", got, input)
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
