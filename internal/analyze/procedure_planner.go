package analyze

import (
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
	"github.com/harumiWeb/xlflow/internal/vba/effects"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// procedureFeature identifies cheap, immutable evidence that a procedure may
// require one of the semantic analysis domains. A feature is deliberately not
// a diagnostic conclusion: it only proves applicability (or the absence of
// applicability when the corresponding unknown bit is also clear).
type procedureFeature uint64

const (
	featureArray procedureFeature = 1 << iota
	featureReDim
	featureLoop
	featureRangeArray
	featureRuntimeExpression
	featureObject
	featureDictionaryCollection
	featureOnError
	featureDataflow
	featureProcessLaunch
	featureSQL
	featureHTTP
	featureFileIO
	featureResourceAcquire
	featureResourceRelease
	featureExcel
	featureExcelOperation
	featureApplicationState
	featureEventHandler
	featureCalls
	featureByRefCalls
	featureMemberAccess
	procedureFeatureLimit
)

const allProcedureFeatures = procedureFeatureLimit - 1

// procedureFeatureSet is a compact three-state summary. A bit absent from both
// masks is proven absent; present wins over unknown when both are supplied by
// independent facts.
type procedureFeatureSet struct {
	present procedureFeature
	unknown procedureFeature
}

func (features *procedureFeatureSet) add(feature procedureFeature) {
	features.present |= feature
	features.unknown &^= feature
}

func (features *procedureFeatureSet) addUnknown(feature procedureFeature) {
	features.unknown |= feature &^ features.present
}

func (features procedureFeatureSet) mayHave(feature procedureFeature) bool {
	return features.present&feature != 0 || features.unknown&feature != 0
}

func (features procedureFeatureSet) mayHaveAll(required procedureFeature) bool {
	return required == 0 || (features.present|features.unknown)&required == required
}

func (features *procedureFeatureSet) observeDeclaration(declaration procedureir.Declaration) {
	if declaration.Recovered || len(declaration.ConditionalBranches) > 0 {
		features.addUnknown(allProcedureFeatures)
		return
	}
	if declaration.IsArray || declaration.ValueShape == procedureir.ValueShapeFixedArray || declaration.ValueShape == procedureir.ValueShapeDynamicArray {
		features.add(featureArray)
	}
	if declaration.IsObject || isObjectType(declaration.Type) {
		features.add(featureObject)
	}
	if dictionaryCollectionType(declaration.Type) {
		features.add(featureDictionaryCollection)
	}
}

func (features *procedureFeatureSet) observeStatement(statement procedureir.Statement) {
	if statement.Recovered || statement.Kind == procedureir.StatementRecovered || statement.Kind == procedureir.StatementUnknown || len(statement.ConditionalBranches) > 0 {
		features.addUnknown(allProcedureFeatures)
		return
	}
	switch statement.Kind {
	case procedureir.StatementReDim:
		features.add(featureArray | featureReDim)
	case procedureir.StatementFor, procedureir.StatementForEach, procedureir.StatementDo, procedureir.StatementWhile:
		features.add(featureLoop)
	case procedureir.StatementOnError, procedureir.StatementResume:
		features.add(featureOnError)
	case procedureir.StatementCall:
		features.add(featureCalls | featureByRefCalls)
	}
	features.observeText(statement.Text)
}

func (features *procedureFeatureSet) observeExpression(expression procedureir.Expression) {
	if expression.Recovered || expression.Kind == procedureir.ExpressionUnknown {
		features.addUnknown(allProcedureFeatures)
		return
	}
	switch expression.Kind {
	case procedureir.ExpressionBinary, procedureir.ExpressionUnary:
		features.add(featureRuntimeExpression)
	case procedureir.ExpressionMember:
		features.add(featureMemberAccess)
	case procedureir.ExpressionCall:
		features.add(featureCalls | featureByRefCalls)
	case procedureir.ExpressionNew:
		features.add(featureObject)
	}
	features.observeText(expression.Text)
}

func (features *procedureFeatureSet) observeCall(call procedureir.CallSite) {
	features.add(featureCalls | featureByRefCalls)
	features.observeText(call.Callee.Text + " " + call.Callee.BaseName + " " + call.Callee.Member)
	switch call.Resolution.Status {
	case procedureir.ResolutionMatched, procedureir.ResolutionBuiltinLike:
		// These are the only statuses that prove a stable target. A matched call
		// can still mutate ByRef/module array or object state, which is why the
		// call features above remain present.
	case procedureir.ResolutionExternal, procedureir.ResolutionMemberCall:
		// The syntax is owned and can classify known APIs, but the external
		// implementation can still supply values or mutations to flow domains.
		features.addUnknown(allProcedureFeatures)
	default:
		features.addUnknown(allProcedureFeatures)
	}
}

func (features *procedureFeatureSet) observeText(text string) {
	lower := strings.ToLower(text)
	if lower == "" {
		return
	}
	if containsAny(lower, "redim ", "erase ", "ubound(", "lbound(", "isarray(", "array(", "split(") {
		features.add(featureArray)
	}
	if strings.Contains(lower, "redim ") {
		features.add(featureReDim)
	}
	if containsAny(lower, "range(", ".value", ".value2") {
		features.add(featureRangeArray)
	}
	if containsAny(lower, "cbyte(", "cint(", "clng(", "clnglng(", "csng(", "cdbl(", "ccur(", "cdec(", " / ", " mod ", " ^ ") {
		features.add(featureRuntimeExpression)
	}
	if containsAny(lower, "scripting.dictionary", "createobject(\"scripting.dictionary\")", "new collection", " as collection", ".comparemode", ".exists(", ".keys", ".items") {
		features.add(featureDictionaryCollection | featureObject)
	}
	if containsAny(lower, " on error ", "on error ", " resume ", "resume next") {
		features.add(featureOnError)
	}
	if containsAny(lower, "shell(", "shell ", ".run(", ".exec(", "shellexecute", "cmd.exe", "powershell") {
		features.add(featureDataflow | featureProcessLaunch)
	}
	if containsAny(lower, ".execute", ".commandtext", "adodb.command", "adodb.connection", "recordset.open") {
		features.add(featureDataflow | featureSQL)
	}
	if containsAny(lower, "xmlhttp", "winhttprequest", ".open(\"get", ".open(\"post", "setrequestheader", ".send") {
		features.add(featureDataflow | featureHTTP)
	}
	if containsAny(lower, "kill ", "rmdir ", "filecopy ", "name ", " open ", "open ", "saveas", "deletefile", "copyfile", "movefile", "opentextfile") {
		features.add(featureDataflow | featureFileIO)
	}
	if containsAny(lower, "workbooks.open", "open ") {
		features.add(featureResourceAcquire)
	}
	if containsAny(lower, ".close", "close #", "close ") {
		features.add(featureResourceRelease)
	}
	if containsAny(lower, "application.", "workbook", "worksheet", "worksheets", "sheets", "range(", ".range", "cells(", ".cells", ".rows", ".columns") {
		features.add(featureExcel)
	}
	if containsAny(lower, ".value", ".value2", ".rows", ".columns", ".resize", ".find(") {
		features.add(featureExcelOperation)
	}
	// Keep VBA242's prerequisite aligned with its owned range-shape scanner.
	// UsedRange and EntireRow/EntireColumn need not include Range(...) syntax.
	if vba242LooksLikeRangeOperation(text) {
		features.add(featureExcel | featureExcelOperation)
	}
	if containsAny(lower, "screenupdating", "enableevents", "displayalerts", "application.calculation", "statusbar") {
		features.add(featureApplicationState)
	}
}

func containsAny(value string, needles ...string) bool {
	for _, needle := range needles {
		if strings.Contains(value, needle) {
			return true
		}
	}
	return false
}

func dictionaryCollectionType(typeName string) bool {
	lower := strings.ToLower(strings.TrimSpace(typeName))
	return lower == "collection" || lower == "dictionary" || strings.HasSuffix(lower, ".dictionary")
}

func finalizeProcedureFeatures(features procedureFeatureSet, document procedureir.DocumentIR, procedure procedureir.ProcedureIR, graphPresent bool, graphUnknown bool) procedureFeatureSet {
	if document.Parse.HasError || document.Parse.HasMissing || procedure.Symbol.Recovered || len(procedure.Symbol.ConditionalBranches) > 0 {
		features.addUnknown(allProcedureFeatures)
	}
	if procedure.Symbol.IsEventHandler {
		features.add(featureEventHandler)
	}
	if procedure.Symbol.IsArray || procedure.Symbol.ValueShape == procedureir.ValueShapeFixedArray || procedure.Symbol.ValueShape == procedureir.ValueShapeDynamicArray {
		features.add(featureArray)
	}
	if isObjectType(procedure.Symbol.ReturnType) {
		features.add(featureObject)
	}
	for _, parameter := range procedure.Symbol.Parameters {
		if parameter.Recovered {
			features.addUnknown(allProcedureFeatures)
			continue
		}
		if parameter.IsArray || parameter.ParamArray || parameter.ValueShape == procedureir.ValueShapeFixedArray || parameter.ValueShape == procedureir.ValueShapeDynamicArray {
			features.add(featureArray)
		}
		if isObjectType(parameter.Type) {
			features.add(featureObject)
		}
		if dictionaryCollectionType(parameter.Type) {
			features.add(featureDictionaryCollection)
		}
	}
	if !graphPresent || graphUnknown {
		features.addUnknown(allProcedureFeatures)
	}
	return features
}

type procedureRuleRequirement struct {
	id     string
	domain analysisstats.Domain
	any    procedureFeature
	all    procedureFeature
	always bool
}

var procedureRuleRequirements = [...]procedureRuleRequirement{
	{id: "VBA249", domain: analysisstats.DomainRuntime, any: featureRuntimeExpression | featureArray | featureCalls},
	{id: "VBA208", domain: analysisstats.DomainArray, any: featureArray | featureReDim | featureCalls},
	{id: "VBA209", domain: analysisstats.DomainArray, any: featureArray | featureRangeArray | featureCalls},
	{id: "VBA226", domain: analysisstats.DomainArray, any: featureRangeArray | featureExcel | featureCalls},
	{id: "VBA227", domain: analysisstats.DomainArray, any: featureArray | featureLoop | featureCalls},
	{id: "VBA241", domain: analysisstats.DomainArray, all: featureReDim | featureLoop},
	{id: "VBA249", domain: analysisstats.DomainArray, any: featureArray | featureCalls},
	{id: "VBA101", domain: analysisstats.DomainArray, all: featureArray | featureObject, always: true},
	{id: "VBA102", domain: analysisstats.DomainArray, all: featureArray | featureObject, always: true},
	{id: "VBA202", domain: analysisstats.DomainObject, any: featureObject | featureMemberAccess | featureCalls},
	{id: "VBA207", domain: analysisstats.DomainDictionary, any: featureDictionaryCollection | featureCalls},
	{id: "VBA213", domain: analysisstats.DomainDictionary, any: featureDictionaryCollection | featureLoop | featureCalls},
	{id: "VBA230", domain: analysisstats.DomainDictionary, any: featureDictionaryCollection | featureCalls},
	{id: "VBA231", domain: analysisstats.DomainDictionary, any: featureDictionaryCollection | featureLoop | featureCalls},
	{id: "VBA232", domain: analysisstats.DomainDictionary, any: featureDictionaryCollection | featureCalls},
	{id: "VBA233", domain: analysisstats.DomainDictionary, any: featureDictionaryCollection | featureCalls},
	{id: "VBA234", domain: analysisstats.DomainDictionary, any: featureDictionaryCollection | featureLoop | featureCalls},
	{id: "VBA235", domain: analysisstats.DomainDictionary, any: featureDictionaryCollection | featureCalls},
	{id: "VBA204", domain: analysisstats.DomainError, any: featureOnError},
	{id: "VBA214", domain: analysisstats.DomainError, any: featureOnError},
	{id: "VBA237", domain: analysisstats.DomainError, any: featureOnError | featureCalls},
	{id: "VBA224", domain: analysisstats.DomainDataflow, any: featureDataflow | featureCalls},
	{id: "VBA236", domain: analysisstats.DomainDataflow, any: featureProcessLaunch | featureCalls},
	{id: "VBA239", domain: analysisstats.DomainDataflow, any: featureSQL | featureCalls},
	{id: "VBA245", domain: analysisstats.DomainDataflow, any: featureFileIO | featureCalls},
	{id: "VBA246", domain: analysisstats.DomainDataflow, any: featureHTTP | featureCalls},
	{id: "VBA247", domain: analysisstats.DomainDataflow, any: featureHTTP | featureCalls},
	{id: "VBA219", domain: analysisstats.DomainResource, any: featureResourceAcquire | featureCalls},
	{id: "VBA225", domain: analysisstats.DomainExcel, any: featureExcel | featureCalls | featureMemberAccess, all: featureLoop},
	{id: "VBA238", domain: analysisstats.DomainExcel, any: featureExcel | featureCalls | featureMemberAccess, all: featureLoop},
	{id: "VBA242", domain: analysisstats.DomainExcel, any: featureExcel | featureExcelOperation},
	{id: "VBA243", domain: analysisstats.DomainExcel, any: featureExcel | featureExcelOperation},
	{id: "VBA203", domain: analysisstats.DomainApplicationState, any: featureApplicationState | featureCalls},
	{id: "VBA220", domain: analysisstats.DomainApplicationState, any: featureEventHandler | featureCalls | featureApplicationState},
	{id: "VBA221", domain: analysisstats.DomainApplicationState, any: featureApplicationState | featureCalls},
}

type procedureAnalysisPlan struct {
	enabled uint16
	planned uint16
}

func procedureDomainBit(domain analysisstats.Domain) uint16 { return uint16(1) << uint(domain) }

func (plan procedureAnalysisPlan) enabledDomain(domain analysisstats.Domain) bool {
	return plan.enabled&procedureDomainBit(domain) != 0
}

func (plan procedureAnalysisPlan) runs(domain analysisstats.Domain) bool {
	return plan.planned&procedureDomainBit(domain) != 0
}

func buildProcedureAnalysisPlan(cfg config.AnalyzeConfig, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) procedureAnalysisPlan {
	features := proc.Features
	if proc.Facts == nil {
		features.addUnknown(allProcedureFeatures)
	}
	if moduleDecls == nil {
		features.addUnknown(featureArray | featureObject | featureDictionaryCollection)
	} else {
		for _, declaration := range moduleDecls {
			if declaration.Array {
				features.add(featureArray)
			}
			if declaration.Object {
				features.add(featureObject)
			}
			if dictionaryCollectionType(declaration.Type) {
				features.add(featureDictionaryCollection)
			}
		}
	}
	if proc.Effects != nil {
		if len(proc.Effects.DirectUncertainty) > 0 || len(proc.Effects.PropagatedUncertainty) > 0 {
			features.addUnknown(allProcedureFeatures)
		}
		if proc.Effects.Error.HasErrorHandler || proc.Effects.Error.UsesResumeNext || proc.Effects.Error.SuppressesErrors || proc.Effects.Error.MayRaise || len(proc.Effects.Error.Direct) > 0 || len(proc.Effects.Error.Propagated) > 0 {
			features.add(featureOnError)
		}
		if proc.Effects.Has(effects.ChangesApplicationState) || proc.Effects.Has(effects.RestoresApplicationState) ||
			proc.Effects.Has(effects.DisablesEvents) || proc.Effects.Has(effects.RestoresEvents) || proc.Effects.Has(effects.ChangesCalculation) {
			features.add(featureApplicationState)
		}
		if proc.Effects.Has(effects.LaunchesProcess) {
			features.add(featureDataflow | featureProcessLaunch)
		}
		if proc.Effects.Has(effects.WritesCells) || proc.Effects.Has(effects.ChangesWorkbook) ||
			proc.Effects.Has(effects.OpensWorkbook) || proc.Effects.Has(effects.Recalculates) ||
			proc.Effects.Has(effects.ChangesSelection) {
			features.add(featureExcel)
		}
		if proc.Effects.Has(effects.OpensWorkbook) {
			features.add(featureResourceAcquire)
		}
	} else if features.mayHave(featureCalls) {
		features.addUnknown(featureApplicationState | featureOnError | featureDataflow | featureArray | featureObject)
	}

	var plan procedureAnalysisPlan
	for _, requirement := range procedureRuleRequirements {
		enabled := requirement.always
		if !enabled {
			var known bool
			enabled, known = config.AnalyzeRuleEnabled(cfg, requirement.id)
			if !known || !enabled {
				continue
			}
		}
		bit := procedureDomainBit(requirement.domain)
		plan.enabled |= bit
		if (requirement.any == 0 || features.mayHave(requirement.any)) && features.mayHaveAll(requirement.all) {
			plan.planned |= bit
		}
	}
	return plan
}

func (proc sourceProcedure) analysisPlan(cfg config.AnalyzeConfig, moduleDecls map[string]sourceDeclaration) procedureAnalysisPlan {
	if proc.PlanReady {
		return proc.Plan
	}
	return buildProcedureAnalysisPlan(cfg, proc, moduleDecls)
}

func materializeProcedureAnalysisPlans(file *parsedFile, projectEffects effects.ProjectSummary, cfg config.AnalyzeConfig) {
	if file == nil {
		return
	}
	procedures := file.Procedures
	if procedures == nil {
		procedures = sourceProceduresFromIR(file.IR, file.CFG)
	}
	moduleDecls := file.moduleDecls()
	for i := range procedures {
		planningProcedure := procedures[i]
		if i < len(file.IR.Procedures) {
			id := procedureEffectIdentity(file.IR, file.IR.Procedures[i].Symbol)
			if summary, ok := projectEffects.Lookup(id); ok {
				planningProcedure.Effects = &summary
			}
			if summary, ok := projectEffects.LookupDirect(id); ok {
				procedures[i].Effects = &summary
			}
		}
		procedures[i].Plan = buildProcedureAnalysisPlan(cfg, planningProcedure, moduleDecls)
		procedures[i].PlanReady = true
	}
	file.Procedures = procedures
}

func projectPlansDomain(cfg config.AnalyzeConfig, files []parsedFile, projectEffects effects.ProjectSummary, domain analysisstats.Domain) bool {
	for _, file := range files {
		moduleDecls := file.moduleDecls()
		for _, proc := range sourceProceduresWithEffects(file, projectEffects) {
			if proc.analysisPlan(cfg, moduleDecls).runs(domain) {
				return true
			}
		}
	}
	return false
}
