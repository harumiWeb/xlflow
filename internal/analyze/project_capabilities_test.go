package analyze

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func plannedProcedure(features procedureFeature, domains ...analysisstats.Domain) sourceProcedure {
	var planned uint16
	for _, domain := range domains {
		planned |= procedureDomainBit(domain)
	}
	return sourceProcedure{
		Features:  procedureFeatureSet{present: features},
		Plan:      procedureAnalysisPlan{planned: planned},
		PlanReady: true,
	}
}

func TestProjectCapabilityDependencyClosure(t *testing.T) {
	plan := projectCapabilityPlan{required: projectCapabilityArrayInterprocedural}
	plan = closeProjectCapabilityDependencies(plan)
	for _, capability := range []projectCapability{
		projectCapabilityArrayInterprocedural,
		projectCapabilityProjectConstants,
		projectCapabilityResolution,
		projectCapabilityTypeDB,
	} {
		if !plan.requires(capability) {
			t.Fatalf("capability %v missing from closure: %#x", capability, plan.required)
		}
	}
}

func TestProjectCapabilityRequirementRowsUseRegisteredRules(t *testing.T) {
	cfg := config.Default().Analyze
	for _, requirement := range procedureRuleRequirements {
		if requirement.capabilities == 0 {
			continue
		}
		if _, known := config.AnalyzeRuleEnabled(cfg, requirement.id); !known && !requirement.always {
			t.Errorf("capability requirement %q is not registered in the analyze rule table", requirement.id)
		}
		if projectCapabilityName(requirement.capabilities&-requirement.capabilities) == "" {
			t.Errorf("capability requirement %q has an unknown direct capability mask %#x", requirement.id, requirement.capabilities)
		}
	}
}

func TestProjectCapabilityPlanSkipsUnneededDomains(t *testing.T) {
	cfg := config.Default().Analyze
	cfg.DetectObjectUseBeforeSet = false
	cfg.DetectApplicationStateRestore = false
	cfg.DetectEventHandlerReentry = false
	cfg.DetectApplicationStateCallEffects = false
	cfg.DetectPublicAPITypeSafety = false
	cfg.DetectErrorSuppressionPropagation = false
	cfg.DetectProcedureCallCycles = false
	files := []parsedFile{{Procedures: []sourceProcedure{plannedProcedure(0)}}}
	plan := buildProjectCapabilityPlan(cfg, files)
	for _, capability := range []projectCapability{
		projectCapabilityEffects,
		projectCapabilityObjectFlow,
		projectCapabilityArrayInterprocedural,
		projectCapabilityDataFlowInputs,
		projectCapabilityApplicationState,
		projectCapabilityEventReentry,
	} {
		if plan.requires(capability) {
			t.Fatalf("unneeded capability %v was planned: %#x", capability, plan.required)
		}
	}
	for _, capability := range []projectCapability{
		projectCapabilityTypeDB,
		projectCapabilityResolution,
		projectCapabilityProjectConstants,
		projectCapabilityByRefSymbols,
	} {
		if !plan.requires(capability) {
			t.Fatalf("baseline capability %v was skipped: %#x", capability, plan.required)
		}
	}
}

func TestProjectCapabilityPlanRetainsTransitiveResolutionForArray(t *testing.T) {
	cfg := config.Default().Analyze
	cfg.DetectObjectUseBeforeSet = false
	cfg.DetectApplicationStateRestore = false
	cfg.DetectEventHandlerReentry = false
	cfg.DetectApplicationStateCallEffects = false
	cfg.DetectPublicAPITypeSafety = false
	cfg.DetectErrorSuppressionPropagation = false
	cfg.DetectProcedureCallCycles = false
	files := []parsedFile{{Procedures: []sourceProcedure{
		plannedProcedure(featureArray, procedureDomainArray),
	}}}
	plan := buildProjectCapabilityPlan(cfg, files)
	if !plan.requires(projectCapabilityArrayInterprocedural) || !plan.requires(projectCapabilityResolution) {
		t.Fatalf("array plan = %#x, want array and resolution", plan.required)
	}
}

func TestProjectEffectsCapabilityRequiresApplicableConsumer(t *testing.T) {
	cfg := config.Default().Analyze
	cfg.DetectApplicationStateRestore = true
	if buildProjectCapabilityPlan(cfg, []parsedFile{{Procedures: []sourceProcedure{plannedProcedure(0)}}}).requires(projectCapabilityEffects) {
		t.Fatal("scalar-only project should not require Effects")
	}
	if !buildProjectCapabilityPlan(cfg, []parsedFile{{Procedures: []sourceProcedure{
		plannedProcedure(featureApplicationState, procedureDomainApplicationState),
	}}}).requires(projectCapabilityEffects) {
		t.Fatal("application-state procedure should require Effects")
	}
}

func TestProjectEffectsCapabilityRetainsCrossModuleEagerContainerDefinitions(t *testing.T) {
	cfg := config.Default().Analyze
	cfg.DetectNonShortCircuitObjectGuard = true
	files := []parsedFile{{Procedures: []sourceProcedure{plannedProcedure(0)}, IR: procedureir.DocumentIR{Procedures: []procedureir.ProcedureIR{{
		Symbol: procedureir.ProcedureSymbol{Name: "IIf"},
	}}}}}
	if !buildProjectCapabilityPlan(cfg, files).requires(projectCapabilityEffects) {
		t.Fatal("project-defined IIf should retain Effects for cross-module VBA212 resolution")
	}
}

func TestProjectCapabilityGetterDetectionUsesProcedureIRWhitespace(t *testing.T) {
	t.Parallel()
	cfg := config.Default().Analyze
	cfg.DetectNonShortCircuitObjectGuard = true
	for _, test := range []struct {
		name   string
		source string
	}{
		{name: "spaces", source: "Public Property  Get Value() As Long\nValue = 1\nEnd Property\n"},
		{name: "tab", source: "Public Property\tGet Value() As Long\nValue = 1\nEnd Property\n"},
	} {
		t.Run(test.name, func(t *testing.T) {
			doc, err := vbaast.ParseDocument("Widget.cls", []byte(test.source))
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()
			ir, err := procedureir.BuildParsed(procedureir.BuildOptions{Path: "Widget.cls", ModuleKind: "class"}, doc)
			if err != nil {
				t.Fatal(err)
			}
			file := parsedFile{Lines: normalizedSourceLines(test.source), IR: ir}
			if vba212SourceMayHaveGetter(file) {
				t.Fatal("source-only getter matcher unexpectedly accepted noncanonical whitespace")
			}
			if !projectHasGetterOrEagerContainer([]parsedFile{file}) {
				t.Fatal("IR getter detection missed Property Get with noncanonical whitespace")
			}
			if !buildProjectCapabilityPlan(cfg, []parsedFile{file}).requires(projectCapabilityEffects) {
				t.Fatal("Property Get should retain Effects for VBA212")
			}
		})
	}
}
