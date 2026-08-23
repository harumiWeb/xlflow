package analyze

import (
	"context"
	"path/filepath"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/effects"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// projectCapability identifies one reusable project-level analysis product.
// The bitset is intentionally internal: capability names and their dependency
// edges are implementation details, not a CLI or LSP contract.
type projectCapability uint16

const (
	projectCapabilityTypeDB projectCapability = 1 << iota
	projectCapabilityResolution
	projectCapabilityProjectConstants
	projectCapabilityByRefSymbols
	projectCapabilityEffects
	projectCapabilityObjectFlow
	projectCapabilityArrayInterprocedural
	projectCapabilityDataFlowInputs
	projectCapabilityDictionaryCollection
	projectCapabilityApplicationState
	projectCapabilityEventReentry
	projectCapabilityPublicAPITypeIndex
	projectCapabilityExcelLoopSymbols
	projectCapabilityExcelAPIHelpers
	projectCapabilityModuleState
)

type projectCapabilityPlan struct {
	required projectCapability
}

// ProjectCapabilityDocument is the protocol-neutral input used by batch and
// LSP planning. LSP snapshots already own IR/CFG, so exposing this small value
// avoids making the internal parsedFile representation part of the public API.
type ProjectCapabilityDocument struct {
	IR     procedureir.DocumentIR
	CFG    vbacfg.Document
	Source string
}

// ProjectCapabilityRequirements is the public, read-only view needed by the
// LSP Full-diagnostics boundary. Capability names remain internal to the
// analyzer; callers only decide which shared project products to request.
type ProjectCapabilityRequirements struct {
	TypeDB               bool
	Resolution           bool
	ProjectConstants     bool
	ByRefSymbols         bool
	Effects              bool
	ObjectFlow           bool
	ArrayInterprocedural bool
	DataFlowInputs       bool
	DictionaryCollection bool
	ApplicationState     bool
	EventReentry         bool
	PublicAPITypeIndex   bool
	ExcelLoopSymbols     bool
	ExcelAPIHelpers      bool
	ModuleState          bool
}

// PlanProjectCapabilities applies the same rule/feature/dependency planner to
// a protocol-neutral project snapshot. It is intentionally pure: no resolver,
// effect summary, or index is constructed while planning.
func PlanProjectCapabilities(cfg config.AnalyzeConfig, documents []ProjectCapabilityDocument) ProjectCapabilityRequirements {
	files := make([]parsedFile, 0, len(documents))
	for _, document := range documents {
		module := strings.TrimSpace(document.IR.ModuleName)
		if module == "" {
			module = strings.TrimSuffix(filepath.Base(document.IR.Path), filepath.Ext(document.IR.Path))
		}
		procedures := sourceProceduresFromIRRef(&document.IR, document.CFG)
		file := parsedFile{
			Path:       document.IR.Path,
			Lines:      normalizedSourceLines(document.Source),
			Module:     module,
			ModuleKind: document.IR.ModuleKind,
			Source:     []byte(document.Source),
			IR:         document.IR,
			CFG:        document.CFG,
			Procedures: procedures,
		}
		file.ModuleFacts = buildModuleAnalysisFacts(file.Lines, file.IR, procedures)
		file.ModuleDeclarations = file.ModuleFacts.moduleDeclarations
		materializeProcedureAnalysisPlans(&file, effects.ProjectSummary{}, cfg)
		files = append(files, file)
	}
	plan := buildProjectCapabilityPlan(cfg, files)
	return ProjectCapabilityRequirements{
		TypeDB:               plan.requires(projectCapabilityTypeDB),
		Resolution:           plan.requires(projectCapabilityResolution),
		ProjectConstants:     plan.requires(projectCapabilityProjectConstants),
		ByRefSymbols:         plan.requires(projectCapabilityByRefSymbols),
		Effects:              plan.requires(projectCapabilityEffects),
		ObjectFlow:           plan.requires(projectCapabilityObjectFlow),
		ArrayInterprocedural: plan.requires(projectCapabilityArrayInterprocedural),
		DataFlowInputs:       plan.requires(projectCapabilityDataFlowInputs),
		DictionaryCollection: plan.requires(projectCapabilityDictionaryCollection),
		ApplicationState:     plan.requires(projectCapabilityApplicationState),
		EventReentry:         plan.requires(projectCapabilityEventReentry),
		PublicAPITypeIndex:   plan.requires(projectCapabilityPublicAPITypeIndex),
		ExcelLoopSymbols:     plan.requires(projectCapabilityExcelLoopSymbols),
		ExcelAPIHelpers:      plan.requires(projectCapabilityExcelAPIHelpers),
		ModuleState:          plan.requires(projectCapabilityModuleState),
	}
}

func (plan projectCapabilityPlan) requires(capability projectCapability) bool {
	return plan.required&capability != 0
}

func (plan *projectCapabilityPlan) require(capability projectCapability) {
	if plan == nil {
		return
	}
	plan.required |= capability
}

// projectCapabilityDependencies is the only transitive dependency table. A
// capability builder must consume its dependencies from the planned bundle;
// it must not construct them privately.
var projectCapabilityDependencies = map[projectCapability]projectCapability{
	projectCapabilityResolution:           projectCapabilityTypeDB,
	projectCapabilityProjectConstants:     projectCapabilityTypeDB | projectCapabilityResolution,
	projectCapabilityByRefSymbols:         projectCapabilityResolution,
	projectCapabilityEffects:              projectCapabilityResolution,
	projectCapabilityObjectFlow:           projectCapabilityResolution,
	projectCapabilityArrayInterprocedural: projectCapabilityResolution | projectCapabilityProjectConstants,
	projectCapabilityDataFlowInputs:       projectCapabilityTypeDB,
	projectCapabilityDictionaryCollection: projectCapabilityResolution,
	projectCapabilityApplicationState:     projectCapabilityEffects,
	projectCapabilityEventReentry:         projectCapabilityEffects,
	projectCapabilityPublicAPITypeIndex:   projectCapabilityTypeDB | projectCapabilityResolution,
	projectCapabilityExcelLoopSymbols:     projectCapabilityTypeDB | projectCapabilityResolution,
	projectCapabilityExcelAPIHelpers:      projectCapabilityResolution,
	projectCapabilityModuleState:          0,
}

func closeProjectCapabilityDependencies(plan projectCapabilityPlan) projectCapabilityPlan {
	for {
		before := plan.required
		for capability, dependencies := range projectCapabilityDependencies {
			if plan.required&capability != 0 {
				plan.required |= dependencies
			}
		}
		if before == plan.required {
			return plan
		}
	}
}

func projectRuleEnabled(cfg config.AnalyzeConfig, id string) bool {
	enabled, known := config.AnalyzeRuleEnabled(cfg, id)
	return known && enabled
}

func dataFlowInputsEnabled(cfg config.AnalyzeConfig) bool {
	return cfg.DetectUntrustedDataFlow || cfg.DetectUnsafeCommandConstruction || cfg.DetectUnsafeSQLConstruction ||
		cfg.DetectUnsafeFilePath || cfg.DetectUnsafeHTTPConfiguration || cfg.DetectMissingHTTPTimeout
}

func projectHasFeature(files []parsedFile, feature procedureFeature) bool {
	for _, file := range files {
		for _, procedure := range file.procedureProjection() {
			if procedure.Features.mayHave(feature) {
				return true
			}
		}
	}
	return false
}

// buildProjectCapabilityPlan combines the enabled-rule plan with the already
// materialized procedure plans. It is intentionally conservative: a missing
// procedure plan keeps the corresponding domain required rather than proving
// that a capability is unnecessary.
func buildProjectCapabilityPlan(cfg config.AnalyzeConfig, files []parsedFile) projectCapabilityPlan {
	var plan projectCapabilityPlan

	// Compile-equivalent and compatibility paths are baseline consumers. The
	// current analyzer always provides these checks, so their preparation must
	// remain unconditional even when optional runtime rules are disabled.
	plan.require(projectCapabilityTypeDB | projectCapabilityResolution | projectCapabilityProjectConstants | projectCapabilityByRefSymbols)

	// The procedure requirement table owns the rule-to-domain and direct
	// capability edges. Project-only rows are evaluated by their rule switch;
	// procedure rows additionally need at least one planned procedure in the
	// declared domain. This keeps capability declarations and applicability
	// prerequisites in one reviewable table.
	emptyEffects := emptyProjectEffects()
	for _, requirement := range procedureRuleRequirements {
		if requirement.capabilities == 0 {
			continue
		}
		if requirement.projectOnly {
			if !projectRuleEnabled(cfg, requirement.id) {
				continue
			}
			if requirement.any != 0 && !projectHasFeature(files, requirement.any) {
				continue
			}
			if requirement.getterSource && !projectHasGetterOrEagerContainer(files) {
				continue
			}
			plan.require(requirement.capabilities)
			continue
		}
		enabled := requirement.always || projectRuleEnabled(cfg, requirement.id)
		if enabled && projectPlansDomain(cfg, files, emptyEffects, requirement.domain) {
			plan.require(requirement.capabilities)
		}
	}

	return closeProjectCapabilityDependencies(plan)
}

func projectHasGetterOrEagerContainer(files []parsedFile) bool {
	for _, file := range files {
		if vba212SourceMayHaveGetter(file) {
			return true
		}
	}
	return projectHasEagerContainerDefinition(files)
}

func projectHasEagerContainerDefinition(files []parsedFile) bool {
	for _, file := range files {
		for _, procedure := range file.IR.Procedures {
			switch strings.ToLower(strings.TrimSpace(procedure.Symbol.Name)) {
			case "iif", "choose", "switch":
				return true
			}
		}
	}
	return false
}

// emptyProjectEffects makes the first applicability pass explicit. A nil
// summary is reserved for standalone callers; batch planning uses an empty,
// immutable value so every procedure gets the same conservative fallback.
func emptyProjectEffects() (out effects.ProjectSummary) { return out }

func projectCapabilityName(capability projectCapability) string {
	switch capability {
	case projectCapabilityTypeDB:
		return "typedb"
	case projectCapabilityResolution:
		return "resolution"
	case projectCapabilityProjectConstants:
		return "project_constants"
	case projectCapabilityByRefSymbols:
		return "byref_symbols"
	case projectCapabilityEffects:
		return "effects"
	case projectCapabilityObjectFlow:
		return "object"
	case projectCapabilityArrayInterprocedural:
		return "array"
	case projectCapabilityDataFlowInputs:
		return "dataflow"
	case projectCapabilityDictionaryCollection:
		return "dictionary"
	case projectCapabilityApplicationState:
		return "application_state"
	case projectCapabilityEventReentry:
		return "event_reentry"
	case projectCapabilityPublicAPITypeIndex:
		return "public_api_type_index"
	case projectCapabilityExcelLoopSymbols:
		return "excel_loop_symbols"
	case projectCapabilityExcelAPIHelpers:
		return "excel_api_helpers"
	case projectCapabilityModuleState:
		return "module_state"
	default:
		return ""
	}
}

// beginProjectCapabilityBuild records entry exactly once and returns the
// matching elapsed-stage finisher. Capability stages intentionally use the
// same stable names as their counters so performance-log consumers can join
// construction counts with elapsed time.
func beginProjectCapabilityBuild(ctx context.Context, capability projectCapability) func(error) {
	name := projectCapabilityName(capability)
	if name == "" {
		return func(error) {}
	}
	return analysisstats.MeasureCapabilityBuild(ctx, "capability_"+name+"_builds")
}

func initializeProjectCapabilityTelemetry(ctx context.Context) {
	if recorder := analysisstats.FromContext(ctx); recorder != nil {
		for _, name := range analysisstats.CapabilityBuildCounters {
			recorder.AddSum(name, 0)
		}
	}
}
