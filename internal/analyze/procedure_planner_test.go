package analyze

import (
	"context"
	"reflect"
	"sync"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestProcedureFeatureSetClassifiesOwnedIRInOneProjection(t *testing.T) {
	facts := newProcedureAnalysisFactsWithDeclarations(
		[]procedureir.Declaration{
			{ID: 1, Name: "values", IsArray: true},
			{ID: 2, Name: "book", IsObject: true, Type: "Workbook"},
			{ID: 3, Name: "items", Type: "Scripting.Dictionary"},
		},
		[]procedureir.Statement{
			{ID: 10, Kind: procedureir.StatementReDim, Text: "ReDim values(10)"},
			{ID: 11, Kind: procedureir.StatementFor, Text: "For i = 1 To 10"},
			{ID: 12, Kind: procedureir.StatementOnError, Text: "On Error GoTo Handler"},
			{ID: 13, Kind: procedureir.StatementAssignment, Text: "Application.ScreenUpdating = False"},
			{ID: 14, Kind: procedureir.StatementAssignment, Text: "Shell(commandLine)"},
			{ID: 15, Kind: procedureir.StatementAssignment, Text: "ADODB.Command.Execute"},
			{ID: 16, Kind: procedureir.StatementAssignment, Text: "WinHttpRequest.Send"},
			{ID: 17, Kind: procedureir.StatementAssignment, Text: "Open path For Input As #1"},
			{ID: 18, Kind: procedureir.StatementAssignment, Text: "Workbooks.Open path"},
			{ID: 19, Kind: procedureir.StatementAssignment, Text: "Close #1"},
		},
		[]procedureir.Expression{
			{ID: 20, Kind: procedureir.ExpressionBinary, Text: "total / count"},
			{ID: 21, Kind: procedureir.ExpressionMember, Text: "Application.Range.Value2"},
			{ID: 22, Kind: procedureir.ExpressionNew, Text: "New Collection"},
		},
		[]procedureir.CallSite{{
			ID:         30,
			Callee:     procedureir.Callee{Text: "DoWork", BaseName: "DoWork"},
			Resolution: procedureir.CallResolution{Status: procedureir.ResolutionBuiltinLike},
		}},
		nil,
	)

	if facts == nil {
		t.Fatal("facts = nil")
	}
	for _, feature := range []procedureFeature{
		featureArray, featureReDim, featureLoop, featureRangeArray,
		featureRuntimeExpression, featureObject, featureDictionaryCollection,
		featureOnError, featureDataflow, featureProcessLaunch, featureSQL,
		featureHTTP, featureFileIO, featureResourceAcquire, featureResourceRelease,
		featureExcel, featureExcelOperation, featureApplicationState,
		featureCalls, featureByRefCalls, featureMemberAccess,
	} {
		if facts.features.present&feature == 0 {
			t.Errorf("feature %v was not classified as present (set=%#v)", feature, facts.features)
		}
		if facts.features.unknown&feature != 0 {
			t.Errorf("feature %v was unexpectedly unknown (set=%#v)", feature, facts.features)
		}
	}
	if facts.features.present&featureEventHandler != 0 {
		t.Error("event handler feature was inferred without a procedure symbol")
	}
}

func TestProcedureFeatureSetProvesScalarProcedureAbsent(t *testing.T) {
	facts := newProcedureAnalysisFactsWithDeclarations(
		nil,
		[]procedureir.Statement{{ID: 1, Kind: procedureir.StatementAssignment, Text: "value = 1"}},
		[]procedureir.Expression{{ID: 2, Kind: procedureir.ExpressionLiteral, Text: "1"}},
		nil,
		nil,
	)
	if facts.features.present != 0 || facts.features.unknown != 0 {
		t.Fatalf("scalar feature set = %#v, want no present or unknown features", facts.features)
	}
	for _, feature := range []procedureFeature{
		featureArray, featureReDim, featureLoop, featureDataflow,
		featureExcel, featureApplicationState, featureCalls,
	} {
		if facts.features.mayHave(feature) {
			t.Errorf("scalar procedure may have %v", feature)
		}
	}
}

func TestProcedureFeatureSetRecognizesQualifiedCollectionAndFileRename(t *testing.T) {
	for _, typeName := range []string{"VBA.Collection", "Scripting.Dictionary"} {
		facts := newProcedureAnalysisFactsWithDeclarations(
			[]procedureir.Declaration{{ID: 1, Name: "items", Type: typeName}},
			nil,
			nil,
			nil,
			nil,
		)
		if facts.features.present&featureDictionaryCollection == 0 {
			t.Errorf("qualified %s declaration was not classified as a collection feature: %#v", typeName, facts.features)
		}
	}
	nameAssignment := newProcedureAnalysisFactsWithDeclarations(
		nil,
		[]procedureir.Statement{{ID: 1, Kind: procedureir.StatementAssignment, Text: "name = value"}},
		nil,
		nil,
		nil,
	)
	if nameAssignment.features.mayHave(featureFileIO) {
		t.Fatalf("ordinary name assignment was classified as file I/O: %#v", nameAssignment.features)
	}
	rename := newProcedureAnalysisFactsWithDeclarations(
		nil,
		[]procedureir.Statement{{ID: 2, Kind: procedureir.StatementAssignment, Text: "Name oldPath As newPath"}},
		nil,
		nil,
		nil,
	)
	if !rename.features.mayHave(featureFileIO) {
		t.Fatalf("Name ... As ... was not classified as file I/O: %#v", rename.features)
	}
}

func TestProcedureFeatureSetFailsOpenForRecoveredAndDynamicIR(t *testing.T) {
	recovered := newProcedureAnalysisFactsWithDeclarations(
		nil,
		[]procedureir.Statement{{ID: 1, Kind: procedureir.StatementRecovered, Recovered: true}},
		nil,
		nil,
		nil,
	)
	if recovered.features.unknown != allProcedureFeatures {
		t.Fatalf("recovered features = %#v, want all features unknown", recovered.features)
	}

	dynamic := newProcedureAnalysisFactsWithDeclarations(
		nil,
		nil,
		nil,
		[]procedureir.CallSite{{
			ID:         2,
			Callee:     procedureir.Callee{Text: "Application.Run", BaseName: "Run"},
			Resolution: procedureir.CallResolution{Status: procedureir.ResolutionDynamic},
		}},
		nil,
	)
	if dynamic.features.present|dynamic.features.unknown != allProcedureFeatures {
		t.Fatalf("dynamic-call features = %#v, want every feature present or unknown", dynamic.features)
	}

	missingGraph := sourceProceduresFromIR(procedureir.DocumentIR{Procedures: []procedureir.ProcedureIR{{
		Symbol: procedureir.ProcedureSymbol{Name: "Recovered", Kind: procedureir.ProcedureSub},
	}}})
	if len(missingGraph) != 1 || missingGraph[0].Features.present|missingGraph[0].Features.unknown != allProcedureFeatures {
		t.Fatalf("missing-graph features = %#v, want every feature potentially applicable", missingGraph)
	}
}

func TestFinalizeProcedureFeaturesIncludesProcedureAndProjectUncertainty(t *testing.T) {
	document := procedureir.DocumentIR{ModuleName: "Sheet1"}
	procedure := procedureir.ProcedureIR{Symbol: procedureir.ProcedureSymbol{
		Name: "Workbook_Open", Kind: procedureir.ProcedureSub, IsEventHandler: true,
		ReturnType: "Workbook", Parameters: []procedureir.Parameter{{Name: "items", Type: "Collection"}},
	}}
	features := finalizeProcedureFeatures(procedureFeatureSet{}, document, procedure, true, false)
	for _, feature := range []procedureFeature{featureEventHandler, featureObject, featureDictionaryCollection} {
		if !features.mayHave(feature) {
			t.Errorf("finalized feature set does not contain %v: %#v", feature, features)
		}
	}

	document.Parse.HasMissing = true
	features = finalizeProcedureFeatures(procedureFeatureSet{}, document, procedure, true, false)
	if features.present|features.unknown != allProcedureFeatures {
		t.Fatalf("missing-node features = %#v, want every feature present or unknown", features)
	}
}

func TestBuildProcedureAnalysisPlanEnabledDisabledAndAnyAllRequirements(t *testing.T) {
	moduleDecls := map[string]sourceDeclaration{}

	cfg := analyzeConfigForRules("VBA208", "VBA241")
	arrayOnly := sourceProcedureWithFeatureSet(procedureFeatureSet{present: featureArray})
	plan := buildProcedureAnalysisPlan(cfg, arrayOnly, moduleDecls)
	if !plan.enabledDomain(procedureDomainArray) || !plan.runs(procedureDomainArray) {
		t.Fatalf("array-only plan = %#v, want enabled and planned for any-of VBA208", plan)
	}

	redimWithoutLoop := sourceProcedureWithFeatureSet(procedureFeatureSet{present: featureReDim})
	plan = buildProcedureAnalysisPlan(cfg, redimWithoutLoop, moduleDecls)
	if !plan.enabledDomain(procedureDomainArray) || !plan.runs(procedureDomainArray) {
		t.Fatalf("ReDim-only plan = %#v, want planned through VBA208 any-of", plan)
	}

	loopWithoutRedim := sourceProcedureWithFeatureSet(procedureFeatureSet{present: featureLoop})
	plan = buildProcedureAnalysisPlan(cfg, loopWithoutRedim, moduleDecls)
	if plan.runs(procedureDomainArray) {
		t.Fatalf("loop-only plan = %#v, VBA241 all-of should be absent", plan)
	}

	unknownLoop := sourceProcedureWithFeatureSet(procedureFeatureSet{
		present: featureReDim,
		unknown: featureLoop,
	})
	plan = buildProcedureAnalysisPlan(cfg, unknownLoop, moduleDecls)
	if !plan.runs(procedureDomainArray) {
		t.Fatalf("unknown-loop plan = %#v, all-of requirement must fail open", plan)
	}

	disabled := analyzeConfigForRules()
	plan = buildProcedureAnalysisPlan(cfg, sourceProcedureWithFeatureSet(procedureFeatureSet{present: featureArray}), moduleDecls)
	if !plan.enabledDomain(procedureDomainArray) {
		t.Fatal("enabled array requirement unexpectedly disabled")
	}
	plan = buildProcedureAnalysisPlan(disabled, arrayOnly, moduleDecls)
	if plan.enabled != procedureDomainBit(procedureDomainArray) || plan.planned != 0 {
		t.Fatalf("disabled plan = %#v, want only the always-enabled compatibility array domain and no run", plan)
	}
}

func TestProcedurePlannerKeepsCompatibilityAndIndirectPrerequisitesApplicable(t *testing.T) {
	moduleDecls := map[string]sourceDeclaration{}

	loopPlan := buildProcedureAnalysisPlan(
		analyzeConfigForRules("VBA227"),
		sourceProcedureWithFeatureSet(procedureFeatureSet{present: featureLoop}),
		moduleDecls,
	)
	if !loopPlan.runs(procedureDomainArray) {
		t.Fatalf("scalar For Each plan = %#v, want VBA227 array domain", loopPlan)
	}

	objectArrayPlan := buildProcedureAnalysisPlan(
		analyzeConfigForRules(),
		sourceProcedureWithFeatureSet(procedureFeatureSet{present: featureArray | featureObject}),
		moduleDecls,
	)
	if !objectArrayPlan.runs(procedureDomainArray) {
		t.Fatalf("object-array compatibility plan = %#v, want VBA101/VBA102 array domain", objectArrayPlan)
	}

	helperLoopPlan := buildProcedureAnalysisPlan(
		analyzeConfigForRules("VBA225"),
		sourceProcedureWithFeatureSet(procedureFeatureSet{present: featureLoop | featureCalls}),
		moduleDecls,
	)
	if !helperLoopPlan.runs(procedureDomainExcel) {
		t.Fatalf("loop helper plan = %#v, want Excel domain", helperLoopPlan)
	}
	memberLoopPlan := buildProcedureAnalysisPlan(
		analyzeConfigForRules("VBA225"),
		sourceProcedureWithFeatureSet(procedureFeatureSet{present: featureLoop | featureMemberAccess}),
		moduleDecls,
	)
	if !memberLoopPlan.runs(procedureDomainExcel) {
		t.Fatalf("parenthesis-free loop member plan = %#v, want Excel domain", memberLoopPlan)
	}

	var fullRangeFeatures procedureFeatureSet
	fullRangeFeatures.observeText(`ws.UsedRange.Formula = "=1"`)
	fullRangePlan := buildProcedureAnalysisPlan(
		analyzeConfigForRules("VBA242"),
		sourceProcedureWithFeatureSet(fullRangeFeatures),
		moduleDecls,
	)
	if !fullRangePlan.runs(procedureDomainExcel) {
		t.Fatalf("UsedRange plan = %#v features=%#v, want Excel domain", fullRangePlan, fullRangeFeatures)
	}
}

func TestBuildProcedureAnalysisPlanUnknownFactsRemainPlanned(t *testing.T) {
	cfg := analyzeConfigForRules(
		"VBA249", "VBA202", "VBA207", "VBA204", "VBA224", "VBA219", "VBA225", "VBA203",
	)
	proc := sourceProcedureWithFeatureSet(procedureFeatureSet{unknown: allProcedureFeatures})
	plan := buildProcedureAnalysisPlan(cfg, proc, map[string]sourceDeclaration{})
	for _, domain := range []analysisstats.Domain{
		procedureDomainRuntime, procedureDomainArray, procedureDomainObject,
		procedureDomainDictionary, procedureDomainError, procedureDomainDataflow,
		procedureDomainResource, procedureDomainExcel, procedureDomainApplicationState,
	} {
		if plan.enabledDomain(domain) && !plan.runs(domain) {
			t.Errorf("unknown facts disabled enabled domain %v: %#v", domain, plan)
		}
	}
}

func TestProcedurePlannerDecisionsRecordPlannedAndSkippedOnce(t *testing.T) {
	recorder := analysisstats.NewRecorder()
	profile := newProcedureDomainProfile(analysisstats.WithRecorder(context.Background(), recorder))
	planned := procedureAnalysisPlan{
		enabled: procedureDomainBit(procedureDomainArray) | procedureDomainBit(procedureDomainDataflow),
		planned: procedureDomainBit(procedureDomainArray),
	}
	profile.plannerDecision(planned, procedureDomainArray)
	profile.plannerDecision(planned, procedureDomainDataflow)
	profile.flush()
	_, counters := recorder.Totals()
	values := map[string]uint64{}
	for _, counter := range counters {
		values[counter.Name] = counter.Value
	}
	if values[analysisstats.ArrayPlannedRunsCounter] != 1 || values[analysisstats.ArraySkippedRunsCounter] != 0 {
		t.Fatalf("array planner counters = %#v", values)
	}
	if values[analysisstats.DataflowSkippedRunsCounter] != 1 || values[analysisstats.DataflowPlannedRunsCounter] != 0 {
		t.Fatalf("dataflow planner counters = %#v", values)
	}
}

func TestPlannerCountersCoverEveryGatedDomain(t *testing.T) {
	for _, domain := range gatedProcedureDomains {
		if _, _, ok := plannerCounters(domain); !ok {
			t.Errorf("gated domain %s has no planner counters", domain)
		}
	}
}

func TestProcedureRuleRequirementsCoverEveryGatedRule(t *testing.T) {
	// This list is intentionally kept beside the planner. If a new expensive
	// domain is gated in analyzer.go without adding its rule prerequisite, this
	// test fails at compile/test time instead of silently changing semantics.
	wanted := map[string]bool{
		"VBA101/array": true, "VBA102/array": true,
		"VBA249/runtime": true, "VBA249/array": true,
		"VBA208/array": true, "VBA209/array": true, "VBA226/array": true, "VBA227/array": true, "VBA241/array": true,
		"VBA202/object":     true,
		"VBA207/dictionary": true, "VBA213/dictionary": true, "VBA230/dictionary": true, "VBA231/dictionary": true,
		"VBA232/dictionary": true, "VBA233/dictionary": true, "VBA234/dictionary": true, "VBA235/dictionary": true,
		"VBA204/error": true, "VBA214/error": true, "VBA237/error": true,
		"VBA224/dataflow": true, "VBA236/dataflow": true, "VBA239/dataflow": true,
		"VBA245/dataflow": true, "VBA246/dataflow": true, "VBA247/dataflow": true,
		"VBA219/resource": true,
		"VBA225/excel":    true, "VBA238/excel": true, "VBA242/excel": true, "VBA243/excel": true,
		"VBA203/application_state": true, "VBA220/application_state": true, "VBA221/application_state": true,
	}
	seen := map[string]bool{}
	always := map[string]bool{}
	for _, requirement := range procedureRuleRequirements {
		key := requirement.id + "/" + requirement.domain.String()[len("procedure_local/"):]
		_, ok := wanted[key]
		if !ok {
			t.Errorf("requirement %q is not in the intended gated rule list", key)
			continue
		}
		if requirement.any == 0 && requirement.all == 0 {
			t.Errorf("requirement %q has no applicability prerequisite", requirement.id)
		}
		seen[key] = true
		always[key] = requirement.always
	}
	for key := range wanted {
		if !seen[key] {
			t.Errorf("gated rule/domain %q has no planner requirement", key)
		}
		id := key[:len("VBA000")]
		if enabled, known := config.AnalyzeRuleEnabled(config.Default().Analyze, id); !known && !always[key] {
			t.Errorf("gated rule %q is not known by the config registry (enabled=%v)", id, enabled)
		}
	}
}

func TestSourceProceduresFromIRFeatureSummaryIsOwnedAndSafeForConcurrentReads(t *testing.T) {
	document := procedureir.DocumentIR{
		Path: "module.bas", ModuleName: "Sheet1",
		Procedures: []procedureir.ProcedureIR{{
			Symbol: procedureir.ProcedureSymbol{
				Name: "Workbook_Open", Kind: procedureir.ProcedureSub,
				IsEventHandler: true, DeclarationRange: vbaRange(1, 1),
			},
			Declarations: []procedureir.Declaration{{ID: 1, IsArray: true}},
			Statements:   []procedureir.Statement{{ID: 2, Kind: procedureir.StatementFor, Text: "For i = 1 To 2"}},
		}},
	}
	flow := cfg.Document{Graphs: []cfg.Graph{{}}}
	procedures := sourceProceduresFromIR(document, flow)
	if len(procedures) != 1 {
		t.Fatalf("procedures = %d, want 1", len(procedures))
	}
	features := procedures[0].Features
	for _, feature := range []procedureFeature{featureArray, featureLoop, featureEventHandler} {
		if !features.mayHave(feature) || features.unknown&feature != 0 {
			t.Fatalf("feature %v = %#v, want present and known", feature, features)
		}
	}

	// sourceProceduresFromIR copies the procedure-owned IR projection before
	// constructing facts; changing the caller's document cannot change it.
	document.Procedures[0].Statements[0].Kind = procedureir.StatementRecovered
	document.Procedures[0].Declarations[0].IsArray = false
	if got := procedures[0].Statements[0].Kind; got != procedureir.StatementFor {
		t.Fatalf("owned statement changed after source mutation: %v", got)
	}
	if !procedures[0].Features.mayHave(featureArray) || procedures[0].Features.unknown&featureArray != 0 {
		t.Fatalf("owned feature changed after source mutation: %#v", procedures[0].Features)
	}

	var wait sync.WaitGroup
	for worker := 0; worker < 8; worker++ {
		wait.Add(1)
		go func() {
			defer wait.Done()
			for i := 0; i < 1000; i++ {
				got := procedures[0].Features
				if !got.mayHave(featureArray) || !got.mayHave(featureLoop) || !got.mayHave(featureEventHandler) {
					t.Errorf("concurrent feature read = %#v", got)
					return
				}
			}
		}()
	}
	wait.Wait()
}

func sourceProcedureWithFeatureSet(features procedureFeatureSet) sourceProcedure {
	return sourceProcedure{Features: features, Facts: &procedureAnalysisFacts{features: features}}
}

func analyzeConfigForRules(ids ...string) config.AnalyzeConfig {
	var cfg config.AnalyzeConfig
	value := reflect.ValueOf(&cfg).Elem()
	for i := 0; i < value.NumField(); i++ {
		if value.Field(i).Kind() == reflect.Bool && value.Field(i).CanSet() {
			value.Field(i).SetBool(false)
		}
	}
	for _, id := range ids {
		switch id {
		case "VBA202":
			cfg.DetectObjectUseBeforeSet = true
		case "VBA203":
			cfg.DetectApplicationStateRestore = true
		case "VBA204":
			cfg.DetectErrorHandlerFallthrough = true
		case "VBA207":
			cfg.DetectDictionaryCollectionGuard = true
		case "VBA208":
			cfg.DetectRedimPreserveDimension = true
		case "VBA209":
			cfg.DetectObjectArrayComparison = true
		case "VBA213":
			cfg.DetectDictionaryIterationValueUsage = true
		case "VBA214":
			cfg.DetectLeakedOnErrorResumeNextScopes = true
		case "VBA219":
			cfg.DetectResourceLeaks = true
		case "VBA220":
			cfg.DetectEventHandlerReentry = true
		case "VBA221":
			cfg.DetectApplicationStateCallEffects = true
		case "VBA224":
			cfg.DetectUntrustedDataFlow = true
		case "VBA225":
			cfg.DetectExcelCellAccessInLoops = true
		case "VBA226":
			cfg.DetectRangeValueArrayShape = true
		case "VBA227":
			cfg.DetectArrayLifecycleSafety = true
		case "VBA230":
			cfg.DetectDictionaryCompareModeOrder = true
		case "VBA231":
			cfg.DetectDictionaryLoopMaterialization = true
		case "VBA232":
			cfg.DetectDictionaryKeyNormalization = true
		case "VBA233":
			cfg.DetectLateBoundDictionaryConstants = true
		case "VBA234":
			cfg.DetectCollectionIterationMutation = true
		case "VBA235":
			cfg.DetectCollectionIndexOrigin = true
		case "VBA236":
			cfg.DetectUnsafeCommandConstruction = true
		case "VBA237":
			cfg.DetectErrorSuppressionPropagation = true
		case "VBA238":
			cfg.DetectLoopInvariantExcelObjectResolution = true
		case "VBA239":
			cfg.DetectUnsafeSQLConstruction = true
		case "VBA241":
			cfg.DetectRedimPreserveInLoops = true
		case "VBA242":
			cfg.DetectExpensiveFullRangeOperations = true
		case "VBA243":
			cfg.DetectValue2PerformanceOpportunities = true
		case "VBA245":
			cfg.DetectUnsafeFilePath = true
		case "VBA246":
			cfg.DetectUnsafeHTTPConfiguration = true
		case "VBA247":
			cfg.DetectMissingHTTPTimeout = true
		case "VBA249":
			cfg.DetectDeterministicRuntimeErrors = true
		default:
			panic("test helper has no field for " + id)
		}
	}
	return cfg
}

func vbaRange(startLine, endLine int) vbaast.Range {
	return vbaast.Range{StartLine: startLine, EndLine: endLine}
}
