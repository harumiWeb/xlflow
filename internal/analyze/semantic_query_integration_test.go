package analyze

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/analyze/semanticquery"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
	"github.com/harumiWeb/xlflow/internal/vba/effects"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestSemanticQueryReusesProcedureKernelsAcrossBatchRevisions(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, "Main.bas", `Option Explicit
Public Const ModuleConstant As Long = 1
Public Sub Run()
    Dim values() As Long
    ReDim values(1 To 2)
    values(1) = 1
End Sub

Public Sub Untouched()
    Dim other() As Long
    ReDim other(1 To 2)
    other(1) = 1
End Sub
`)
	store := semanticquery.New(semanticquery.Options{})
	run := func() (Result, map[string]uint64, error) {
		recorder := analysisstats.NewRecorder()
		ctx := analysisstats.WithRecorder(context.Background(), recorder)
		ctx = semanticquery.WithContext(ctx, semanticquery.Context{Store: store})
		result, err := (Analyzer{RootDir: dir, Config: config.Default()}).RunResultContext(ctx)
		_, counters := recorder.Snapshot()
		return result, counters, err
	}
	first, firstCounters, err := run()
	if err != nil {
		t.Fatal(err)
	}
	second, secondCounters, err := run()
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(first.Findings, second.Findings) {
		t.Fatalf("warm analysis changed findings: first=%#v second=%#v", first.Findings, second.Findings)
	}
	if secondCounters[analysisstats.SemanticQueryHitsCounter] == 0 {
		t.Fatalf("warm analysis recorded no semantic query hits: first=%v second=%v", firstCounters, secondCounters)
	}
	if secondCounters[analysisstats.SemanticQueryRecomputedKernelsCounter] >= firstCounters[analysisstats.SemanticQueryRecomputedKernelsCounter] {
		t.Fatalf("warm analysis recomputed at least as many kernels: first=%v second=%v", firstCounters, secondCounters)
	}
	sourcePath := filepath.Join(dir, "src", "modules", "Main.bas")
	source, err := os.ReadFile(sourcePath)
	if err != nil {
		t.Fatal(err)
	}
	moduleUpdated := strings.Replace(string(source), "ModuleConstant As Long = 1", "ModuleConstant As Long = 2", 1)
	if moduleUpdated == string(source) {
		t.Fatal("module constant replacement did not change source")
	}
	if err := os.WriteFile(sourcePath, []byte(moduleUpdated), 0o644); err != nil {
		t.Fatal(err)
	}
	_, moduleCounters, err := run()
	if err != nil {
		t.Fatal(err)
	}
	if moduleCounters[analysisstats.SemanticQueryMissesCounter] == 0 {
		t.Fatalf("module input edit reused all old kernels: counters=%v", moduleCounters)
	}
	source, err = os.ReadFile(sourcePath)
	if err != nil {
		t.Fatal(err)
	}
	updated := strings.Replace(string(source), "other(1) = 1", "other(1) = 2", 1)
	if err := os.WriteFile(sourcePath, []byte(updated), 0o644); err != nil {
		t.Fatal(err)
	}
	_, changedCounters, err := run()
	if err != nil {
		t.Fatal(err)
	}
	if changedCounters[analysisstats.SemanticQueryHitsCounter] == 0 {
		t.Fatalf("local edit did not reuse the unchanged procedure: counters=%v", changedCounters)
	}
	if changedCounters[analysisstats.SemanticQueryRecomputedKernelsCounter] >= firstCounters[analysisstats.SemanticQueryRecomputedKernelsCounter] {
		t.Fatalf("local edit recomputed every kernel: first=%v changed=%v", firstCounters, changedCounters)
	}
}

func TestSemanticQueryWarmMatchesFreshAcrossRevisionChanges(t *testing.T) {
	dir := t.TempDir()
	callee := `Option Explicit
Public Function Compute(value As Long) As Long
    Compute = value
End Function
`
	caller := `Option Explicit
Public Sub Run()
    Dim result As Long
    result = Compute(1)
End Sub
`
	other := `Option Explicit
Public Function ComputeOther(value As Long) As Long
    ComputeOther = value
End Function
`
	writeCase := func(calleeSource, callerSource string) {
		writeModule(t, dir, "Callee.bas", calleeSource)
		writeModule(t, dir, "Caller.bas", callerSource)
		writeModule(t, dir, "Other.bas", other)
	}
	baseConfig := config.Default()
	writeCase(callee, caller)
	warmStore := semanticquery.New(semanticquery.Options{})
	run := func(store *semanticquery.Store, cfg config.Config) (Result, error) {
		ctx := semanticquery.WithContext(context.Background(), semanticquery.Context{Store: store})
		return (Analyzer{RootDir: dir, Config: cfg}).RunResultContext(ctx)
	}
	if _, err := run(warmStore, baseConfig); err != nil {
		t.Fatal(err)
	}
	cases := []struct {
		name      string
		callee    string
		caller    string
		configure func(config.Config) config.Config
	}{
		{
			name:   "body",
			callee: strings.Replace(callee, "Compute = value", "Compute = value + 1", 1),
			caller: caller,
		},
		{
			name:   "signature",
			callee: strings.Replace(callee, "value As Long", "value As Variant", 1),
			caller: caller,
		},
		{
			name:   "resolution-redirect",
			callee: callee,
			caller: strings.Replace(caller, "Compute(1)", "ComputeOther(1)", 1),
		},
		{
			name:   "call-add",
			callee: callee,
			caller: strings.Replace(caller, "End Sub", "    Call ComputeOther(2)\nEnd Sub", 1),
		},
		{
			name:   "call-remove",
			callee: callee,
			caller: strings.Replace(caller, "    result = Compute(1)\n", "", 1),
		},
		{
			name:   "effect",
			callee: strings.Replace(callee, "    Compute = value", "    Range(\"A1\").Value = value\n    Compute = value", 1),
			caller: caller,
		},
		{
			name:   "module-context",
			callee: strings.Replace(callee, "Option Explicit\n", "Option Explicit\nPublic Const ContextMarker As Long = 1\n", 1),
			caller: caller,
		},
		{
			name:   "configuration",
			callee: callee,
			caller: caller,
			configure: func(cfg config.Config) config.Config {
				cfg.Analyze.DetectRangeValueArrayShape = !cfg.Analyze.DetectRangeValueArrayShape
				return cfg
			},
		},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			writeCase(tc.callee, tc.caller)
			cfg := baseConfig
			if tc.configure != nil {
				cfg = tc.configure(cfg)
			}
			warm, err := run(warmStore, cfg)
			if err != nil {
				t.Fatal(err)
			}
			fresh, err := run(semanticquery.New(semanticquery.Options{}), cfg)
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(warm.Findings, fresh.Findings) {
				t.Fatalf("warm findings differ from fresh recomputation: warm=%#v fresh=%#v", warm.Findings, fresh.Findings)
			}
			if !reflect.DeepEqual(warm.Warnings, fresh.Warnings) || !reflect.DeepEqual(warm.PreflightFindings, fresh.PreflightFindings) {
				t.Fatalf("warm auxiliary results differ from fresh recomputation: warm warnings=%#v preflight=%#v fresh warnings=%#v preflight=%#v", warm.Warnings, warm.PreflightFindings, fresh.Warnings, fresh.PreflightFindings)
			}
			warmJSON, err := json.Marshal(warm)
			if err != nil {
				t.Fatal(err)
			}
			freshJSON, err := json.Marshal(fresh)
			if err != nil {
				t.Fatal(err)
			}
			if string(warmJSON) != string(freshJSON) {
				t.Fatalf("warm JSON differs from fresh recomputation:\nwarm=%s\nfresh=%s", warmJSON, freshJSON)
			}
		})
	}
}

func TestSemanticArrayModuleFingerprintIncludesDeclarations(t *testing.T) {
	facts := &moduleAnalysisFacts{
		moduleDeclarations: map[string]sourceDeclaration{"value": {}},
	}
	file := parsedFile{ModuleFacts: facts}
	before := semanticArrayModuleFingerprint(file)
	facts.moduleDeclarations["other"] = sourceDeclaration{}
	after := semanticArrayModuleFingerprint(file)
	if before == after {
		t.Fatal("array module fingerprint ignored a module declaration change")
	}
}

func TestPrepareSemanticQueryFactsPublishesPreparedArrayIndex(t *testing.T) {
	ctx := analysisContext{}
	prepareSemanticQueryFacts(Analyzer{}, nil, &ctx)
	if ctx.arrayCapabilityIndex == nil {
		t.Fatal("prepared array capability index was not published to the caller context")
	}
}

func TestSemanticProcedureIdentityUsesCanonicalPath(t *testing.T) {
	file := parsedFile{Path: filepath.Join("workspace", "nested", "..", "Main.bas"), Module: "Main"}
	proc := sourceProcedure{Name: "Run", ProcedureKind: procedureir.ProcedureSub}
	want := semanticCanonicalPath(file.Path) + "::Main::Run::sub"
	if got := semanticProcedureIdentity(file, proc); got != want {
		t.Fatalf("procedure identity = %q, want %q", got, want)
	}
}

func TestSemanticProcedureFingerprintIncludesNearbyEvidence(t *testing.T) {
	file := parsedFile{
		Path:   "Main.bas",
		Lines:  []string{"Option Explicit", "", "Public Sub Run()", "    Range(\"A1\").Value = 1", "End Sub", "", "' evidence"},
		Source: []byte("Option Explicit\n\nPublic Sub Run()\n    Range(\"A1\").Value = 1\nEnd Sub\n\n' evidence\n"),
	}
	proc := sourceProcedure{Name: "Run", Kind: "Sub", ProcedureKind: procedureir.ProcedureSub, StartLine: 3, EndLine: 5, StartByte: 20, EndByte: 68}
	before := semanticProcedureBaseFingerprint(Analyzer{RootDir: "workspace"}, file, proc)
	file.Lines[6] = "' changed evidence"
	after := semanticProcedureBaseFingerprint(Analyzer{RootDir: "workspace"}, file, proc)
	if before == after {
		t.Fatal("procedure fingerprint ignored nearby evidence source")
	}
}

func TestSemanticProcedureFingerprintIncludesSourcePosition(t *testing.T) {
	file := parsedFile{
		Path:   "Main.bas",
		Source: []byte("Public Sub Run()\nEnd Sub\n"),
	}
	proc := sourceProcedure{
		Name:          "Run",
		Kind:          "Sub",
		ProcedureKind: procedureir.ProcedureSub,
		StartLine:     1,
		EndLine:       2,
		StartByte:     0,
		EndByte:       len(file.Source) - 1,
	}
	moved := proc
	moved.StartLine++
	moved.EndLine++
	if semanticProcedureBaseFingerprint(Analyzer{RootDir: "workspace"}, file, proc) == semanticProcedureBaseFingerprint(Analyzer{RootDir: "workspace"}, file, moved) {
		t.Fatal("procedure fingerprint ignored a source position change")
	}
}

func TestSemanticPreparedPlanFingerprintIncludesPrivateBits(t *testing.T) {
	base := procedureAnalysisPlan{enabled: 1, planned: 1, enabledKernels: 1, plannedKernels: 1}
	changed := base
	changed.plannedDataflowLanes = 1
	if semanticProcedurePlanFingerprint(base) == semanticProcedurePlanFingerprint(changed) {
		t.Fatal("prepared plan fingerprint ignored private plan bits")
	}
	features := procedureFeatureSet{present: featureArray}
	unknown := procedureFeatureSet{present: featureArray, unknown: featureDataflow}
	if semanticProcedureFeatureFingerprint(features) == semanticProcedureFeatureFingerprint(unknown) {
		t.Fatal("procedure feature fingerprint ignored unknown bits")
	}
}

func TestSemanticEffectFingerprintIncludesEffectLeaves(t *testing.T) {
	without := &effects.ProcedureSummary{}
	with := &effects.ProcedureSummary{Direct: []effects.Evidence{{Effect: effects.WritesCells}}}
	if semanticProcedureEffectFingerprint(without) == semanticProcedureEffectFingerprint(with) {
		t.Fatal("effect fingerprint ignored a changed effect leaf")
	}
}
