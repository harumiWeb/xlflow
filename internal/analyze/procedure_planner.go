package analyze

import (
	"strings"
	"unicode"

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
	features.observeText(call.Callee.Text)
	features.observeText(call.Callee.BaseName)
	features.observeText(call.Callee.Member)
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
	if containsAny(lower, "shell(", "shell ", ".run(", ".run ", ".exec(", ".exec ", "wscript.shell.run", "wscript.shell.exec", "shellexecute", "cmd.exe", "powershell") {
		features.add(featureDataflow | featureProcessLaunch)
	}
	if containsAny(lower, ".execute", ".commandtext", "adodb.command", "adodb.connection", "recordset.open") {
		features.add(featureDataflow | featureSQL)
	}
	if containsAny(lower, "xmlhttp", "winhttprequest", ".open(\"get", ".open(\"post", "open \"get", "open \"post", "open \"put", "open \"delete", "open \"patch", "setrequestheader", ".send") {
		features.add(featureDataflow | featureHTTP)
	}
	compactMemberText := lower
	if strings.Contains(lower, "workbooks") && strings.Contains(lower, ".") {
		compactMemberText = strings.Map(func(r rune) rune {
			if unicode.IsSpace(r) {
				return -1
			}
			return r
		}, lower)
	}
	trimmedLower := strings.TrimSpace(lower)
	openStatement := strings.HasPrefix(trimmedLower, "open ")
	openMember := strings.Contains(lower, ".open") || strings.Contains(compactMemberText, ".open")
	httpOpen := containsAny(lower, ".open(\"get", ".open(\"post", "open \"get", "open \"post", "open \"put", "open \"delete", "open \"patch")
	fileOpen := openStatement || (openMember && !httpOpen)
	if containsAny(lower, "kill ", "kill(", "rmdir ", "rmdir(", "filecopy ", "filecopy(", "saveas", "deletefile", "copyfile", "movefile", "opentextfile") || fileOpen || looksLikeFileRename(lower) {
		features.add(featureDataflow | featureFileIO)
	}
	if containsAny(compactMemberText, "workbooks.open") || fileOpen {
		features.add(featureResourceAcquire)
	}
	if containsAny(compactMemberText, "workbooks.open", "saveas") {
		features.add(featureDataflow)
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
	if containsAny(lower,
		"screenupdating", "enableevents", "displayalerts", "application.calculation", "statusbar",
		"cursor", "interactive", "asktoupdatelinks", "automationsecurity", "cutcopymode",
	) {
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

func looksLikeFileRename(lower string) bool {
	trimmed := strings.TrimSpace(lower)
	return strings.HasPrefix(trimmed, "name ") && strings.Contains(trimmed, " as ")
}

func dictionaryCollectionType(typeName string) bool {
	lower := strings.ToLower(strings.TrimSpace(typeName))
	return lower == "collection" || lower == "dictionary" || strings.HasSuffix(lower, ".collection") || strings.HasSuffix(lower, ".dictionary")
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
	id           string
	domain       analysisstats.Domain
	any          procedureFeature
	all          procedureFeature
	always       bool
	capabilities projectCapability
	projectOnly  bool
	getterSource bool
}

var procedureRuleRequirements = [...]procedureRuleRequirement{
	{id: "VBA249", domain: analysisstats.DomainRuntime, any: featureRuntimeExpression | featureArray | featureCalls},
	{id: "VBA212", domain: analysisstats.DomainOther, capabilities: projectCapabilityEffects, projectOnly: true, getterSource: true},
	{id: "VBA208", domain: analysisstats.DomainArray, any: featureArray | featureReDim | featureCalls, capabilities: projectCapabilityArrayInterprocedural},
	{id: "VBA209", domain: analysisstats.DomainArray, any: featureArray | featureRangeArray | featureObject | featureCalls, capabilities: projectCapabilityArrayInterprocedural},
	{id: "VBA226", domain: analysisstats.DomainArray, any: featureRangeArray | featureExcel | featureCalls, capabilities: projectCapabilityArrayInterprocedural},
	{id: "VBA227", domain: analysisstats.DomainArray, any: featureArray | featureLoop | featureCalls, capabilities: projectCapabilityArrayInterprocedural},
	{id: "VBA241", domain: analysisstats.DomainArray, all: featureReDim | featureLoop, capabilities: projectCapabilityArrayInterprocedural},
	{id: "VBA249", domain: analysisstats.DomainArray, any: featureArray | featureCalls, capabilities: projectCapabilityArrayInterprocedural},
	{id: "VBA101", domain: analysisstats.DomainArray, all: featureArray | featureObject, always: true, capabilities: projectCapabilityArrayInterprocedural},
	{id: "VBA102", domain: analysisstats.DomainArray, all: featureArray | featureObject, always: true, capabilities: projectCapabilityArrayInterprocedural},
	{id: "VBA202", domain: analysisstats.DomainObject, any: featureObject | featureMemberAccess, capabilities: projectCapabilityObjectFlow},
	{id: "VBA207", domain: analysisstats.DomainDictionary, any: featureDictionaryCollection | featureCalls, capabilities: projectCapabilityDictionaryCollection},
	{id: "VBA213", domain: analysisstats.DomainDictionary, any: featureDictionaryCollection | featureLoop | featureCalls, capabilities: projectCapabilityDictionaryCollection},
	{id: "VBA230", domain: analysisstats.DomainDictionary, any: featureDictionaryCollection | featureCalls, capabilities: projectCapabilityDictionaryCollection},
	{id: "VBA231", domain: analysisstats.DomainDictionary, any: featureDictionaryCollection | featureLoop | featureCalls, capabilities: projectCapabilityDictionaryCollection},
	{id: "VBA232", domain: analysisstats.DomainDictionary, any: featureDictionaryCollection | featureCalls, capabilities: projectCapabilityDictionaryCollection},
	{id: "VBA233", domain: analysisstats.DomainDictionary, any: featureDictionaryCollection | featureCalls, capabilities: projectCapabilityDictionaryCollection},
	{id: "VBA234", domain: analysisstats.DomainDictionary, any: featureDictionaryCollection | featureLoop | featureCalls, capabilities: projectCapabilityDictionaryCollection},
	{id: "VBA235", domain: analysisstats.DomainDictionary, any: featureDictionaryCollection | featureCalls, capabilities: projectCapabilityDictionaryCollection},
	{id: "VBA204", domain: analysisstats.DomainError, any: featureOnError},
	{id: "VBA214", domain: analysisstats.DomainError, any: featureOnError},
	{id: "VBA237", domain: analysisstats.DomainError, any: featureOnError, capabilities: projectCapabilityEffects},
	{id: "VBA224", domain: analysisstats.DomainDataflow, any: featureDataflow, capabilities: projectCapabilityDataFlowInputs},
	{id: "VBA236", domain: analysisstats.DomainDataflow, any: featureProcessLaunch, capabilities: projectCapabilityDataFlowInputs},
	{id: "VBA239", domain: analysisstats.DomainDataflow, any: featureSQL, capabilities: projectCapabilityDataFlowInputs},
	{id: "VBA245", domain: analysisstats.DomainDataflow, any: featureFileIO, capabilities: projectCapabilityDataFlowInputs},
	{id: "VBA246", domain: analysisstats.DomainDataflow, any: featureHTTP, capabilities: projectCapabilityDataFlowInputs},
	{id: "VBA247", domain: analysisstats.DomainDataflow, any: featureHTTP, capabilities: projectCapabilityDataFlowInputs},
	{id: "VBA219", domain: analysisstats.DomainResource, any: featureResourceAcquire | featureCalls},
	{id: "VBA225", domain: analysisstats.DomainExcel, any: featureExcel | featureCalls | featureMemberAccess, all: featureLoop, capabilities: projectCapabilityExcelLoopSymbols},
	{id: "VBA238", domain: analysisstats.DomainExcel, any: featureExcel | featureCalls | featureMemberAccess, all: featureLoop, capabilities: projectCapabilityExcelLoopSymbols},
	{id: "VBA242", domain: analysisstats.DomainExcel, any: featureExcel | featureExcelOperation},
	{id: "VBA243", domain: analysisstats.DomainExcel, any: featureExcel | featureExcelOperation},
	{id: "VBA203", domain: analysisstats.DomainApplicationState, any: featureApplicationState, capabilities: projectCapabilityApplicationState},
	{id: "VBA220", domain: analysisstats.DomainApplicationState, any: featureEventHandler | featureApplicationState, capabilities: projectCapabilityEventReentry},
	{id: "VBA221", domain: analysisstats.DomainApplicationState, any: featureApplicationState, capabilities: projectCapabilityApplicationState},
	{id: "VBA218", domain: analysisstats.DomainOther, capabilities: projectCapabilityExcelAPIHelpers, projectOnly: true},
	{id: "VBA222", domain: analysisstats.DomainOther, capabilities: projectCapabilityPublicAPITypeIndex, projectOnly: true},
	{id: "VBA240", domain: analysisstats.DomainOther, capabilities: projectCapabilityModuleState, projectOnly: true},
	{id: "VBA244", domain: analysisstats.DomainOther, any: featureCalls, capabilities: projectCapabilityEffects, projectOnly: true},
}

type procedureAnalysisPlan struct {
	enabled            uint16
	planned            uint16
	enabledKernels     uint16
	plannedKernels     uint16
	enabledProjections uint64
	plannedProjections uint64
}

// procedureKernel identifies a semantic preparation pass.  Kernel values are
// intentionally compact and are used as indexes into the static canonical
// order below; a procedure plan therefore does not allocate slices or maps.
type procedureKernel uint8

const (
	procedureKernelRuntime procedureKernel = iota
	procedureKernelArray
	procedureKernelObject
	procedureKernelDictionary
	procedureKernelError
	procedureKernelDataflow
	procedureKernelResource
	procedureKernelExcel
	procedureKernelApplicationState
	procedureKernelLimit
)

const procedureKernelMaskWidth = 16

var _ [procedureKernelMaskWidth - int(procedureKernelLimit)]struct{}

var canonicalProcedureKernelOrder = [...]procedureKernel{
	procedureKernelRuntime,
	procedureKernelArray,
	procedureKernelObject,
	procedureKernelDictionary,
	procedureKernelError,
	procedureKernelDataflow,
	procedureKernelResource,
	procedureKernelExcel,
	procedureKernelApplicationState,
}

func procedureKernelBit(kernel procedureKernel) uint16 { return uint16(1) << kernel }

// procedureProjection is the stable, diagnostic-facing part of a kernel. A
// projection only turns an already materialized semantic result into findings;
// it must never rebuild the result. The executor preserves its established
// static append sequence while consulting these bits at every migrated step.
type procedureProjection uint8

const (
	procedureProjectionRuntime procedureProjection = iota
	procedureProjectionArrayRedim
	procedureProjectionArrayComparison
	procedureProjectionArrayRangeShape
	procedureProjectionArrayLifecycle
	procedureProjectionArrayRedimLoop
	procedureProjectionObject
	procedureProjectionDictionaryGuard
	procedureProjectionDictionaryIteration
	procedureProjectionDictionaryCompareMode
	procedureProjectionDictionaryLoop
	procedureProjectionDictionaryKeyNormalization
	procedureProjectionDictionaryLateBound
	procedureProjectionDictionaryCollectionMutation
	procedureProjectionDictionaryIndexOrigin
	procedureProjectionErrorHandler
	procedureProjectionErrorResume
	procedureProjectionErrorSuppression
	procedureProjectionDataflowUntrusted
	procedureProjectionDataflowCommand
	procedureProjectionDataflowSQL
	procedureProjectionDataflowFile
	procedureProjectionDataflowHTTP
	procedureProjectionResource
	procedureProjectionExcelLoop
	procedureProjectionExcelInvariant
	procedureProjectionExcelRange
	procedureProjectionExcelValue2
	procedureProjectionApplicationRestore
	procedureProjectionApplicationEffects
	procedureProjectionApplicationReentry
	procedureProjectionLimit
)

const procedureProjectionMaskWidth = 64

var _ [procedureProjectionMaskWidth - int(procedureProjectionLimit)]struct{}

func procedureProjectionBit(projection procedureProjection) uint64 { return uint64(1) << projection }

func procedureKernelForDomain(domain analysisstats.Domain) (procedureKernel, bool) {
	switch domain {
	case procedureDomainRuntime:
		return procedureKernelRuntime, true
	case procedureDomainArray:
		return procedureKernelArray, true
	case procedureDomainObject:
		return procedureKernelObject, true
	case procedureDomainDictionary:
		return procedureKernelDictionary, true
	case procedureDomainError:
		return procedureKernelError, true
	case procedureDomainDataflow:
		return procedureKernelDataflow, true
	case procedureDomainResource:
		return procedureKernelResource, true
	case procedureDomainExcel:
		return procedureKernelExcel, true
	case procedureDomainApplicationState:
		return procedureKernelApplicationState, true
	default:
		return 0, false
	}
}

func procedureProjectionForRequirement(requirement procedureRuleRequirement) procedureProjection {
	// VBA249 has two independent projections: runtime expressions and the
	// array transfer. Keep this split explicit so both can consume one array
	// result without making the runtime kernel own array diagnostics.
	if requirement.id == "VBA249" && requirement.domain == procedureDomainArray {
		return procedureProjectionArrayRedim
	}
	switch requirement.id {
	case "VBA249":
		return procedureProjectionRuntime
	case "VBA208":
		return procedureProjectionArrayRedim
	case "VBA209":
		return procedureProjectionArrayComparison
	case "VBA226":
		return procedureProjectionArrayRangeShape
	case "VBA227":
		return procedureProjectionArrayLifecycle
	case "VBA241":
		return procedureProjectionArrayRedimLoop
	case "VBA101", "VBA102":
		return procedureProjectionArrayLifecycle
	case "VBA202":
		return procedureProjectionObject
	case "VBA207":
		return procedureProjectionDictionaryGuard
	case "VBA213":
		return procedureProjectionDictionaryIteration
	case "VBA230":
		return procedureProjectionDictionaryCompareMode
	case "VBA231":
		return procedureProjectionDictionaryLoop
	case "VBA232":
		return procedureProjectionDictionaryKeyNormalization
	case "VBA233":
		return procedureProjectionDictionaryLateBound
	case "VBA234":
		return procedureProjectionDictionaryCollectionMutation
	case "VBA235":
		return procedureProjectionDictionaryIndexOrigin
	case "VBA204":
		return procedureProjectionErrorHandler
	case "VBA214":
		return procedureProjectionErrorResume
	case "VBA237":
		return procedureProjectionErrorSuppression
	case "VBA224":
		return procedureProjectionDataflowUntrusted
	case "VBA236":
		return procedureProjectionDataflowCommand
	case "VBA239":
		return procedureProjectionDataflowSQL
	case "VBA245":
		return procedureProjectionDataflowFile
	case "VBA246", "VBA247":
		return procedureProjectionDataflowHTTP
	case "VBA219":
		return procedureProjectionResource
	case "VBA225":
		return procedureProjectionExcelLoop
	case "VBA238":
		return procedureProjectionExcelInvariant
	case "VBA242":
		return procedureProjectionExcelRange
	case "VBA243":
		return procedureProjectionExcelValue2
	case "VBA203":
		return procedureProjectionApplicationRestore
	case "VBA221":
		return procedureProjectionApplicationEffects
	case "VBA220":
		return procedureProjectionApplicationReentry
	}
	// A future requirement must still be represented in a plan. Falling back
	// to the first projection in its domain is conservative and keeps unknown
	// facts fail-open until a dedicated projection is registered.
	switch requirement.domain {
	case procedureDomainArray:
		return procedureProjectionArrayLifecycle
	case procedureDomainObject:
		return procedureProjectionObject
	case procedureDomainDictionary:
		return procedureProjectionDictionaryGuard
	case procedureDomainError:
		return procedureProjectionErrorHandler
	case procedureDomainDataflow:
		return procedureProjectionDataflowUntrusted
	case procedureDomainResource:
		return procedureProjectionResource
	case procedureDomainExcel:
		return procedureProjectionExcelLoop
	case procedureDomainApplicationState:
		return procedureProjectionApplicationRestore
	default:
		return procedureProjectionRuntime
	}
}

func procedureDomainBit(domain analysisstats.Domain) uint16 { return uint16(1) << uint(domain) }

// Procedure plans use a compact uint16 mask. Keep this compile-time guard in
// sync with the DomainOther sentinel so a future domain cannot silently make
// procedureDomainBit overflow the mask.
const procedureDomainMaskWidth = 16

var _ [procedureDomainMaskWidth - int(procedureDomainOther) - 1]struct{}

func (plan procedureAnalysisPlan) enabledDomain(domain analysisstats.Domain) bool {
	return plan.enabled&procedureDomainBit(domain) != 0
}

func (plan procedureAnalysisPlan) runs(domain analysisstats.Domain) bool {
	return plan.planned&procedureDomainBit(domain) != 0
}

func (plan procedureAnalysisPlan) enabledKernel(kernel procedureKernel) bool {
	if plan.enabledKernels != 0 {
		return plan.enabledKernels&procedureKernelBit(kernel) != 0
	}
	// Plans constructed by older callers (and small unit fixtures) only carry
	// domain masks. Deriving the kernel closure keeps that representation
	// compatible while all materialized plans use explicit kernel bits.
	for domain := analysisstats.DomainRuntime; domain < procedureDomainOther; domain++ {
		mapped, ok := procedureKernelForDomain(domain)
		if ok && mapped == kernel && plan.enabledDomain(domain) {
			return true
		}
	}
	return false
}

func (plan procedureAnalysisPlan) runsKernel(kernel procedureKernel) bool {
	if plan.plannedKernels != 0 {
		return plan.plannedKernels&procedureKernelBit(kernel) != 0
	}
	for domain := analysisstats.DomainRuntime; domain < procedureDomainOther; domain++ {
		mapped, ok := procedureKernelForDomain(domain)
		if ok && mapped == kernel && plan.runs(domain) {
			return true
		}
	}
	return false
}

func (plan procedureAnalysisPlan) enabledProjection(projection procedureProjection) bool {
	if plan.enabledProjections == 0 {
		if kernel, ok := procedureKernelForProjection(projection); ok {
			return plan.enabledKernel(kernel)
		}
	}
	return plan.enabledProjections&procedureProjectionBit(projection) != 0
}

func (plan procedureAnalysisPlan) runsProjection(projection procedureProjection) bool {
	if plan.plannedProjections == 0 {
		if kernel, ok := procedureKernelForProjection(projection); ok {
			return plan.runsKernel(kernel)
		}
	}
	return plan.plannedProjections&procedureProjectionBit(projection) != 0
}

func (plan procedureAnalysisPlan) runsAnyProjection(projections ...procedureProjection) bool {
	for _, projection := range projections {
		if plan.runsProjection(projection) {
			return true
		}
	}
	return false
}

func procedureKernelForProjection(projection procedureProjection) (procedureKernel, bool) {
	switch projection {
	case procedureProjectionRuntime:
		return procedureKernelRuntime, true
	case procedureProjectionArrayRedim, procedureProjectionArrayComparison,
		procedureProjectionArrayRangeShape, procedureProjectionArrayLifecycle,
		procedureProjectionArrayRedimLoop:
		return procedureKernelArray, true
	case procedureProjectionObject:
		return procedureKernelObject, true
	case procedureProjectionDictionaryGuard, procedureProjectionDictionaryIteration,
		procedureProjectionDictionaryCompareMode, procedureProjectionDictionaryLoop,
		procedureProjectionDictionaryKeyNormalization, procedureProjectionDictionaryLateBound,
		procedureProjectionDictionaryCollectionMutation, procedureProjectionDictionaryIndexOrigin:
		return procedureKernelDictionary, true
	case procedureProjectionErrorHandler, procedureProjectionErrorResume, procedureProjectionErrorSuppression:
		return procedureKernelError, true
	case procedureProjectionDataflowUntrusted, procedureProjectionDataflowCommand,
		procedureProjectionDataflowSQL, procedureProjectionDataflowFile, procedureProjectionDataflowHTTP:
		return procedureKernelDataflow, true
	case procedureProjectionResource:
		return procedureKernelResource, true
	case procedureProjectionExcelLoop, procedureProjectionExcelInvariant,
		procedureProjectionExcelRange, procedureProjectionExcelValue2:
		return procedureKernelExcel, true
	case procedureProjectionApplicationRestore, procedureProjectionApplicationEffects,
		procedureProjectionApplicationReentry:
		return procedureKernelApplicationState, true
	default:
		return 0, false
	}
}

func moduleDeclarationFeatures(moduleDecls map[string]sourceDeclaration) procedureFeatureSet {
	var features procedureFeatureSet
	if moduleDecls == nil {
		features.addUnknown(featureArray | featureObject | featureDictionaryCollection)
		return features
	}
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
	return features
}

func buildProcedureAnalysisPlan(cfg config.AnalyzeConfig, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) procedureAnalysisPlan {
	return buildProcedureAnalysisPlanWithModuleFeatures(cfg, proc, moduleDeclarationFeatures(moduleDecls))
}

func buildProcedureAnalysisPlanWithModuleFeatures(cfg config.AnalyzeConfig, proc sourceProcedure, moduleFeatures procedureFeatureSet) procedureAnalysisPlan {
	features := proc.Features
	if proc.Facts == nil {
		features.addUnknown(allProcedureFeatures)
	}
	features.addUnknown(moduleFeatures.unknown)
	features.add(moduleFeatures.present)
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
	}

	var plan procedureAnalysisPlan
	for _, requirement := range procedureRuleRequirements {
		if requirement.projectOnly {
			continue
		}
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
		applicable := (requirement.any == 0 || features.mayHave(requirement.any)) && features.mayHaveAll(requirement.all)
		if kernel, ok := procedureKernelForDomain(requirement.domain); ok {
			plan.enabledKernels |= procedureKernelBit(kernel)
		}
		projection := procedureProjectionForRequirement(requirement)
		plan.enabledProjections |= procedureProjectionBit(projection)
		if applicable {
			plan.planned |= bit
			if kernel, ok := procedureKernelForDomain(requirement.domain); ok {
				plan.plannedKernels |= procedureKernelBit(kernel)
			}
			plan.plannedProjections |= procedureProjectionBit(projection)
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
		procedures = sourceProceduresFromIRRef(&file.IR, file.CFG)
	}
	moduleFeatures := moduleDeclarationFeatures(file.moduleDecls())
	for i := range procedures {
		planningProcedure := procedures[i]
		if i < len(file.IR.Procedures) {
			id := procedureEffectIdentity(file.IR, file.IR.Procedures[i].Symbol)
			// Planning consumes the propagated summary so indirect project effects
			// fail open; rule execution retains the direct summary expected by the
			// existing semantic domains.
			if summary, ok := projectEffects.Lookup(id); ok {
				planningProcedure.Effects = &summary
			}
			if summary, ok := projectEffects.LookupDirect(id); ok {
				procedures[i].Effects = &summary
			}
		}
		procedures[i].Plan = buildProcedureAnalysisPlanWithModuleFeatures(cfg, planningProcedure, moduleFeatures)
		procedures[i].PlanReady = true
	}
	file.Procedures = procedures
}

func projectPlansDomain(cfg config.AnalyzeConfig, files []parsedFile, projectEffects effects.ProjectSummary, domain analysisstats.Domain) bool {
	for _, file := range files {
		procedures := file.procedureView()
		materialized := file.Procedures != nil
		if materialized {
			for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
				proc := procedures.valueAt(procedureIndex)
				if !proc.PlanReady {
					materialized = false
					break
				}
			}
		}
		if !materialized {
			procedures = newReadOnlySpan(sourceProceduresWithEffects(file, projectEffects))
		}
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			plan := proc.Plan
			if !proc.PlanReady {
				plan = proc.analysisPlan(cfg, file.moduleDecls())
			}
			if plan.runs(domain) {
				return true
			}
		}
	}
	return false
}
