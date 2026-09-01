package analyze

import (
	"context"
	"encoding/json"
	"fmt"
	"path/filepath"
	"runtime"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/analyze/semanticquery"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
	"github.com/harumiWeb/xlflow/internal/vba/effects"
)

// withSemanticQueryContext makes the process-local store and the request's
// recorder available to every kernel evaluation without changing the public
// analyzer method signatures. A caller-owned store (for example, an LSP
// workspace) wins; ordinary batch calls use the process store.
func withSemanticQueryContext(ctx context.Context) (context.Context, semanticquery.Context) {
	if ctx == nil {
		ctx = context.Background()
	}
	value := semanticquery.FromContext(ctx)
	if value.Store == nil {
		value.Store = semanticquery.DefaultStore()
	}
	if value.Metrics == nil {
		value.Metrics = analysisstats.FromContext(ctx)
	}
	return semanticquery.WithContext(ctx, value), value
}

// semanticQueryRevisionFacts are immutable inputs shared by every file and
// every semantic lane in one analyzer revision.  The values deliberately stay
// opaque strings at the semanticquery boundary, while the analyzer keeps the
// expensive projections (configuration, type database, source and call
// resolution) alive for the lifetime of the revision.
type semanticQueryRevisionFacts struct {
	config       string
	configDigest string
	analyzer     string
	typeDB       string
	arrayIndex   *semanticArrayCapabilityIndex
}

// semanticArrayCapabilityIndex is a normalized, immutable projection of the
// project-wide array summaries. The old query path rebuilt and formatted every
// source map for each procedure. Values are formatted once at revision setup;
// procedure capabilities then perform direct leaf lookups only.
type semanticArrayCapabilityIndex struct {
	returns           map[string]string
	allocationGuards  map[string]string
	safeArrayLengths  map[string]string
	byRefAllocations  map[string]string
	conditional       map[string]string
	length            map[string]string
	moduleAllocations map[string]string
	moduleEntryStates map[string]string
	byRefEntryStates  map[string]string
	byRefConditions   map[string]string
	participants      map[string]string
}

type semanticFileQueryFacts struct {
	revision   *semanticQueryRevisionFacts
	module     string
	array      string
	dataflow   string
	procedures map[int]*semanticProcedureQueryFacts
}

type semanticProcedureQueryFacts struct {
	identity          string
	base              string
	plan              string
	effects           string
	capabilities      map[string]string
	capabilityDigests map[string]string
	callDependencies  []semanticquery.Key
}

// prepareSemanticQueryFacts builds the revision/file/procedure projections
// once, after all project capabilities have been materialized.  It is called
// before procedure workers start and the resulting pointers are immutable.
func prepareSemanticQueryFacts(a Analyzer, files []parsedFile, ctx *analysisContext) {
	configFingerprint := semanticConfigFingerprint(a.Config)
	revision := &semanticQueryRevisionFacts{
		config:       configFingerprint,
		configDigest: semanticquery.Hash(configFingerprint),
		analyzer:     semanticAnalyzerCapability(a),
		typeDB:       semanticTypeDBFingerprint(a),
	}
	revision.arrayIndex = buildSemanticArrayCapabilityIndex(*ctx)
	ctx.arrayCapabilityIndex = revision.arrayIndex
	targetFacts := make(map[string]*semanticProcedureQueryFacts)
	for index := range files {
		file := &files[index]
		module := semanticModuleFactsFingerprint(*file)
		file.moduleFactsFingerprint = module
		facts := &semanticFileQueryFacts{
			revision:   revision,
			module:     module,
			array:      semanticArrayModuleFingerprint(*file),
			dataflow:   semanticDataflowModuleFingerprint(*file),
			procedures: make(map[int]*semanticProcedureQueryFacts, len(file.Procedures)),
		}
		for _, proc := range file.procedures() {
			capabilities := make(map[string]string, 3)
			// Only lanes that the materialized plan can execute need a prepared
			// capability. This keeps the revision setup from scanning the array
			// indexes for procedures that are known not to enter that kernel.
			if !proc.PlanReady || proc.Plan.runsKernel(procedureKernelArray) {
				capabilities["array"] = semanticAnalysisCapabilityUncached(*ctx, *file, proc, "array")
			}
			if !proc.PlanReady || proc.Plan.runsDataflowLane(procedureDataflowLaneGeneric) {
				capabilities["dataflow"] = semanticAnalysisCapabilityUncached(analysisContext{}, *file, proc, "dataflow")
			}
			if !proc.PlanReady || proc.Plan.runsDataflowLane(procedureDataflowLaneHTTP) {
				capabilities["http"] = semanticAnalysisCapabilityUncached(analysisContext{}, *file, proc, "http")
			}
			capabilityDigests := make(map[string]string, len(capabilities))
			for kernel, capability := range capabilities {
				capabilityDigests[kernel] = semanticquery.Hash(capability, revision.analyzer, revision.typeDB)
			}
			procedure := &semanticProcedureQueryFacts{
				identity:          semanticProcedureIdentity(*file, proc),
				base:              semanticProcedureBaseFingerprint(a, *file, proc),
				plan:              semanticProcedurePlanFingerprint(proc.Plan),
				effects:           semanticquery.Hash(semanticProcedureEffectFingerprint(proc.Effects)),
				capabilities:      capabilities,
				capabilityDigests: capabilityDigests,
			}
			facts.procedures[proc.Index] = procedure
			targetFacts[semanticCallTargetKey(file.Path, proc.Module+"."+proc.Name, string(proc.ProcedureKind), proc.StartLine)] = procedure
		}
		file.semanticQueryFacts = facts
	}
	for index := range files {
		file := &files[index]
		facts := file.semanticQueryFacts
		if facts == nil {
			continue
		}
		for _, proc := range file.procedures() {
			if procedure := facts.procedures[proc.Index]; procedure != nil {
				procedure.callDependencies = semanticCallDependencies(*file, proc, targetFacts)
			}
		}
	}
}

func buildSemanticArrayCapabilityIndex(ctx analysisContext) *semanticArrayCapabilityIndex {
	return &semanticArrayCapabilityIndex{
		returns:           normalizeSemanticCapabilityMap(ctx.arrayReturns),
		allocationGuards:  normalizeSemanticCapabilityMap(ctx.arrayAllocationGuards),
		safeArrayLengths:  normalizeSemanticCapabilityMap(ctx.arraySafeArrayLengthGuards),
		byRefAllocations:  normalizeSemanticCapabilityMap(ctx.arrayByRefAllocations),
		conditional:       normalizeSemanticCapabilityMap(ctx.arrayByRefConditionalAllocations),
		length:            normalizeSemanticCapabilityMap(ctx.arrayByRefLengthAllocations),
		moduleAllocations: normalizeSemanticCapabilityMap(ctx.arrayModuleAllocations),
		moduleEntryStates: normalizeSemanticCapabilityMap(ctx.arrayModuleEntryStates),
		byRefEntryStates:  normalizeSemanticCapabilityMap(ctx.arrayByRefEntryStates),
		byRefConditions:   normalizeSemanticCapabilityMap(ctx.arrayByRefEntryConditions),
		participants:      normalizeSemanticCapabilityMap(ctx.arrayParticipants),
	}
}

func normalizeSemanticCapabilityMap[T any, M ~map[string]T](values M) map[string]string {
	if len(values) == 0 {
		return nil
	}
	grouped := make(map[string][]string, len(values))
	for key, value := range values {
		normalized := strings.ToLower(strings.TrimSpace(key))
		if normalized == "" {
			continue
		}
		grouped[normalized] = append(grouped[normalized], fmt.Sprintf("%#v", value))
	}
	result := make(map[string]string, len(grouped))
	for key, values := range grouped {
		sort.Strings(values)
		unique := values[:0]
		for _, value := range values {
			if len(unique) == 0 || unique[len(unique)-1] != value {
				unique = append(unique, value)
			}
		}
		result[key] = strings.Join(unique, "\x1f")
	}
	return result
}

func semanticConfigFingerprint(cfg config.Config) string {
	return semanticStableJSON(cfg.Analyze)
}

func semanticDataflowModuleFingerprint(file parsedFile) string {
	if len(file.DataFlowModuleBindings) == 0 {
		return ""
	}
	return semanticStableJSON(file.DataFlowModuleBindings)
}

func semanticArrayModuleFingerprint(file parsedFile) string {
	facts := file.ModuleFacts
	if facts == nil {
		return ""
	}
	return semanticquery.Hash(
		semanticStableJSON(facts.moduleDeclarations),
		semanticStableJSON(facts.arrayOperationsByName),
		semanticStableJSON(facts.arrayOperationsByLine),
		semanticStableJSON(file.ArrayIntegerModuleConstants),
		semanticStableJSON(file.RangeValueModuleConstants),
		semanticStableJSON(file.ConstantValues),
		fmt.Sprintf("option-base:%d:%t", file.ArrayOptionBase, file.ArrayOptionBaseSet),
	)
}

func semanticProcedureBaseFingerprint(a Analyzer, file parsedFile, proc sourceProcedure) string {
	start, end := proc.StartByte, proc.EndByte
	if start < 0 {
		start = 0
	}
	if end < start {
		end = start
	}
	if start > len(file.Source) {
		start = len(file.Source)
	}
	if end > len(file.Source) {
		end = len(file.Source)
	}
	body := file.Source[start:end]
	return semanticquery.Hash(
		a.RootDir,
		semanticCanonicalPath(file.Path),
		file.Module,
		file.ModuleKind,
		proc.Name,
		proc.Kind,
		string(proc.ProcedureKind),
		proc.ReturnType,
		fmt.Sprintf("position:%d:%d", proc.StartLine, proc.EndLine),
		string(body),
		semanticProcedureNearbyFingerprint(file, proc),
		semanticProcedureFeatureFingerprint(proc.Features),
	)
}

// procedureAnalysisPlan and procedureFeatureSet intentionally keep their
// compact bit fields private. JSON cannot observe those fields, so the query
// boundary uses explicit scalar encodings instead of serializing a misleading
// empty object for every plan or feature set.
func semanticProcedurePlanFingerprint(plan procedureAnalysisPlan) string {
	return fmt.Sprintf("%d:%d:%d:%d:%d:%d:%d:%d",
		plan.enabled,
		plan.planned,
		plan.enabledKernels,
		plan.plannedKernels,
		plan.enabledProjections,
		plan.plannedProjections,
		plan.enabledDataflowLanes,
		plan.plannedDataflowLanes,
	)
}

func semanticProcedureFeatureFingerprint(features procedureFeatureSet) string {
	return fmt.Sprintf("%d:%d", uint64(features.present), uint64(features.unknown))
}

func semanticProcedureEffectFingerprint(summary *effects.ProcedureSummary) string {
	if summary == nil {
		return ""
	}
	parts := make([]string, 0, len(semanticEffectKinds)+8)
	for _, kind := range semanticEffectKinds {
		if summary.Has(kind) {
			parts = append(parts, "effect:"+string(kind))
		}
	}
	for _, uncertainty := range append(append([]effects.CallUncertainty(nil), summary.DirectUncertainty...), summary.PropagatedUncertainty...) {
		parts = append(parts, "uncertainty:"+string(uncertainty.Kind))
	}
	sort.Strings(parts)
	parts = append(parts,
		fmt.Sprintf("error-handler:%t", summary.Error.HasErrorHandler),
		fmt.Sprintf("error-resume-next:%t", summary.Error.UsesResumeNext),
		fmt.Sprintf("error-suppresses:%t", summary.Error.SuppressesErrors),
		fmt.Sprintf("error-rethrows:%t", summary.Error.RethrowsErrors),
		fmt.Sprintf("error-success:%t", summary.Error.ReturnsSuccessFlag),
		fmt.Sprintf("error-raises:%t", summary.Error.MayRaise),
		fmt.Sprintf("error-logs:%t", summary.Error.LogsAndContinues),
	)
	return semanticquery.Hash(parts...)
}

var semanticEffectKinds = [...]effects.EffectKind{
	effects.WritesCells,
	effects.ChangesWorkbook,
	effects.OpensWorkbook,
	effects.OpensFile,
	effects.ClosesWorkbook,
	effects.DisablesEvents,
	effects.RestoresEvents,
	effects.ChangesCalculation,
	effects.Recalculates,
	effects.ChangesSelection,
	effects.ChangesControls,
	effects.ShowsDialog,
	effects.LaunchesProcess,
	effects.SuppressesErrors,
	effects.RaisesError,
	effects.ChangesApplicationState,
	effects.RestoresApplicationState,
}

// semanticProcedureNearbyFingerprint covers the small source window used by
// finding evidence. The procedure body already covers all lines inside the
// declaration; two lines on either side are enough to keep NearbyCode fresh
// without making an edit elsewhere in the module invalidate every procedure.
func semanticProcedureNearbyFingerprint(file parsedFile, proc sourceProcedure) string {
	if len(file.Lines) == 0 {
		return ""
	}
	start := proc.StartLine - 3 // StartLine is one-based; include two prior lines.
	if start < 0 {
		start = 0
	}
	end := proc.EndLine + 2 // Include two following lines (exclusive index below).
	if end > len(file.Lines) {
		end = len(file.Lines)
	}
	if end <= start {
		return ""
	}
	return semanticStableJSON(file.Lines[start:end])
}

func semanticCallTargetKey(file, qualifiedName, kind string, line int) string {
	return strings.Join([]string{
		semanticCanonicalPath(file),
		strings.ToLower(strings.TrimSpace(qualifiedName)),
		strings.ToLower(strings.TrimSpace(kind)),
		fmt.Sprintf("%d", line),
	}, "\x00")
}

func semanticCallDependencies(file parsedFile, proc sourceProcedure, targetFacts map[string]*semanticProcedureQueryFacts) []semanticquery.Key {
	dependencies := make([]semanticquery.Key, 0, proc.Calls.Len())
	for index := 0; index < proc.Calls.Len(); index++ {
		call, ok := proc.Calls.At(index)
		if !ok {
			continue
		}
		resolutionFingerprint := semanticStableJSON(call.Resolution)
		if len(call.Resolution.Candidates) == 0 {
			dependencies = append(dependencies, semanticquery.Key{
				Procedure:   semanticCanonicalPath(file.Path),
				Fingerprint: semanticquery.Hash(fmt.Sprintf("call:%d", call.ID), resolutionFingerprint),
				Kernel:      "call-resolution",
			})
			continue
		}
		fingerprint := semanticquery.Hash(resolutionFingerprint)
		for _, candidate := range call.Resolution.Candidates {
			dependencies = append(dependencies, semanticquery.Key{
				Procedure:   fmt.Sprintf("%s::%s::%s::%d", semanticCanonicalPath(candidate.File), candidate.QualifiedName, candidate.Kind, candidate.Line),
				Fingerprint: fingerprint,
				Kernel:      "call-resolution",
			})
			if target := targetFacts[semanticCallTargetKey(candidate.File, candidate.QualifiedName, candidate.Kind, candidate.Line)]; target != nil {
				dependencies = append(dependencies,
					semanticquery.Key{Procedure: target.identity, Fingerprint: target.base, Kernel: "callee-source"},
					semanticquery.Key{Procedure: target.identity, Fingerprint: target.effects, Kernel: "callee-effect"},
				)
			}
		}
	}
	if len(dependencies) < 2 {
		return dependencies
	}
	sort.Slice(dependencies, func(i, j int) bool {
		if dependencies[i].Procedure != dependencies[j].Procedure {
			return dependencies[i].Procedure < dependencies[j].Procedure
		}
		if dependencies[i].Fingerprint != dependencies[j].Fingerprint {
			return dependencies[i].Fingerprint < dependencies[j].Fingerprint
		}
		if dependencies[i].Kernel != dependencies[j].Kernel {
			return dependencies[i].Kernel < dependencies[j].Kernel
		}
		if dependencies[i].Config != dependencies[j].Config {
			return dependencies[i].Config < dependencies[j].Config
		}
		return dependencies[i].Capability < dependencies[j].Capability
	})
	unique := dependencies[:0]
	for _, dependency := range dependencies {
		if len(unique) == 0 || unique[len(unique)-1] != dependency {
			unique = append(unique, dependency)
		}
	}
	return unique
}

func semanticQueryRevisionID(rootDir string, cfg config.Config, files []parsedFile) string {
	configBytes, err := json.Marshal(cfg.Analyze)
	if err != nil {
		configBytes = []byte(fmt.Sprintf("%#v", cfg.Analyze))
	}
	ordered := append([]parsedFile(nil), files...)
	sort.Slice(ordered, func(i, j int) bool { return ordered[i].Path < ordered[j].Path })
	parts := []string{"workspace", rootDir, string(configBytes)}
	for _, file := range ordered {
		parts = append(parts, file.Path, file.Module, file.ModuleKind, semanticquery.Hash(string(file.Source)))
	}
	return semanticquery.Hash(parts...)
}

func semanticProcedureIdentity(file parsedFile, proc sourceProcedure) string {
	return fmt.Sprintf("%s::%s::%s::%s", semanticCanonicalPath(file.Path), file.Module, proc.Name, proc.ProcedureKind)
}

// semanticCanonicalPath is the analyzer-side counterpart of the LSP
// symbolFileKey. Query identities and exact document invalidation must agree
// on cleaned, case-insensitive Windows paths even when one side originated in
// a URI and the other originated in a workspace snapshot.
func semanticCanonicalPath(path string) string {
	path = strings.TrimSpace(path)
	if path == "" {
		return ""
	}
	path = filepath.Clean(path)
	if runtime.GOOS == "windows" {
		path = strings.ToLower(path)
	}
	return path
}

func semanticProcedureQueryKey(a Analyzer, file parsedFile, proc sourceProcedure, kernel string, plan procedureAnalysisPlan, capability string) semanticquery.Key {
	if facts := file.semanticQueryFacts; facts != nil {
		if procedure := facts.procedures[proc.Index]; procedure != nil {
			if preparedCapability, ok := procedure.capabilities[kernel]; ok {
				capability = preparedCapability
			}
			planFingerprint := procedure.plan
			if planFingerprint == "" {
				planFingerprint = semanticProcedurePlanFingerprint(plan)
			}
			capabilityDigest := procedure.capabilityDigests[kernel]
			if capabilityDigest == "" {
				capabilityDigest = semanticquery.Hash(capability, facts.revision.analyzer, facts.revision.typeDB)
			}
			return semanticquery.Key{
				Procedure:   procedure.identity,
				Fingerprint: procedure.base,
				Kernel:      kernel + ":" + planFingerprint,
				Config:      facts.revision.configDigest,
				Capability:  capabilityDigest,
			}
		}
	}
	fingerprint := semanticProcedureBaseFingerprint(a, file, proc)
	return semanticquery.Key{
		Procedure:   semanticProcedureIdentity(file, proc),
		Fingerprint: fingerprint,
		Kernel:      kernel + ":" + semanticProcedurePlanFingerprint(plan),
		Config:      semanticquery.Hash(semanticConfigFingerprint(a.Config)),
		Capability:  semanticquery.Hash(capability, semanticAnalyzerCapability(a), semanticTypeDBFingerprint(a)),
	}
}

func semanticQueryDependencies(key semanticquery.Key, file parsedFile, proc sourceProcedure) []semanticquery.Key {
	if facts := file.semanticQueryFacts; facts != nil {
		if procedure := facts.procedures[proc.Index]; procedure != nil {
			kernel := key.Kernel
			if delimiter := strings.IndexByte(kernel, ':'); delimiter >= 0 {
				kernel = kernel[:delimiter]
			}
			dependencies := []semanticquery.Key{
				{Procedure: key.Procedure, Fingerprint: key.Fingerprint, Kernel: "source"},
				{Procedure: key.Procedure, Fingerprint: key.Fingerprint, Kernel: "config", Config: key.Config},
				{Procedure: key.Procedure, Fingerprint: key.Fingerprint, Kernel: "capability", Capability: key.Capability},
				{Procedure: key.Procedure, Fingerprint: procedure.effects, Kernel: "effect-summary"},
			}
			// Keep module inputs at the leaf granularity consumed by each lane.
			// The array leaf contains module array indexes/constants; dataflow and
			// HTTP use their module binding/declaration leaves. This prevents an
			// unrelated lane from being invalidated by a changed map.
			switch kernel {
			case "array":
				dependencies = append(dependencies, semanticquery.Key{Procedure: semanticCanonicalPath(file.Path), Fingerprint: facts.array, Kernel: "array-module"})
			case "dataflow":
				dependencies = append(dependencies, semanticquery.Key{Procedure: semanticCanonicalPath(file.Path), Fingerprint: facts.dataflow, Kernel: "dataflow-module"})
			case "http":
				dependencies = append(dependencies,
					semanticquery.Key{Procedure: semanticCanonicalPath(file.Path), Fingerprint: facts.module, Kernel: "module-context"},
					semanticquery.Key{Procedure: semanticCanonicalPath(file.Path), Fingerprint: facts.dataflow, Kernel: "dataflow-module"},
				)
			default:
				dependencies = append(dependencies, semanticquery.Key{Procedure: semanticCanonicalPath(file.Path), Fingerprint: facts.module, Kernel: "module-context"})
			}
			return append(dependencies, procedure.callDependencies...)
		}
	}
	dependencies := []semanticquery.Key{
		{Procedure: key.Procedure, Fingerprint: key.Fingerprint, Kernel: "source"},
		{Procedure: key.Procedure, Fingerprint: key.Fingerprint, Kernel: "config", Config: key.Config},
		{Procedure: key.Procedure, Fingerprint: key.Fingerprint, Kernel: "capability", Capability: key.Capability},
		{Procedure: semanticCanonicalPath(file.Path), Fingerprint: semanticquery.Hash(semanticModuleFactsFingerprint(file)), Kernel: "module"},
		{Procedure: key.Procedure, Fingerprint: semanticProcedureEffectFingerprint(proc.Effects), Kernel: "effect-summary"},
	}
	for index := 0; index < proc.Calls.Len(); index++ {
		call, ok := proc.Calls.At(index)
		if !ok {
			continue
		}
		resolutionFingerprint := semanticStableJSON(call.Resolution)
		if len(call.Resolution.Candidates) == 0 {
			dependencies = append(dependencies, semanticquery.Key{
				Procedure:   semanticCanonicalPath(file.Path),
				Fingerprint: semanticquery.Hash(fmt.Sprintf("call:%d", call.ID), resolutionFingerprint),
				Kernel:      "call-resolution",
			})
			continue
		}
		for _, candidate := range call.Resolution.Candidates {
			dependencies = append(dependencies, semanticquery.Key{
				Procedure:   fmt.Sprintf("%s::%s::%s::%d", semanticCanonicalPath(candidate.File), candidate.QualifiedName, candidate.Kind, candidate.Line),
				Fingerprint: semanticquery.Hash(resolutionFingerprint),
				Kernel:      "call-resolution",
			})
		}
	}
	return dependencies
}

func semanticAnalyzerCapability(a Analyzer) string {
	return semanticquery.Hash(semanticStableJSON(a.visibleConstants), semanticStableJSON(a.visibleConstantValues))
}

func semanticTypeDBFingerprint(a Analyzer) string {
	if a.typeDB == nil {
		return ""
	}
	// The DB is immutable for an analyzer revision. Serialize its compact
	// canonical maps directly instead of materializing and retaining a large
	// AllConstantsList for every short-lived DB pointer created by tests or
	// realtime requests. encoding/json orders map keys, while the candidate
	// buckets retain their deterministic load order.
	return semanticquery.Hash(
		semanticStableJSON(a.typeDB.Constants),
		semanticStableJSON(a.typeDB.ConstantCandidates),
	)
}

func semanticModuleFactsFingerprint(file parsedFile) string {
	if file.moduleFactsFingerprint != "" {
		return file.moduleFactsFingerprint
	}
	facts := file.ModuleFacts
	if facts == nil {
		return ""
	}
	return semanticquery.Hash(
		semanticStableJSON(facts.moduleDeclarations),
		semanticStableJSON(facts.constants),
		semanticStableJSON(facts.moduleConstantNames),
		semanticStableJSON(facts.procedureNames),
		semanticStableJSON(facts.privateModule),
	)
}

func semanticAnalysisCapability(ctx analysisContext, file parsedFile, proc sourceProcedure, kernel string) string {
	if facts := file.semanticQueryFacts; facts != nil {
		if procedure := facts.procedures[proc.Index]; procedure != nil {
			if capability, ok := procedure.capabilities[kernel]; ok {
				return capability
			}
		}
	}
	return semanticAnalysisCapabilityUncached(ctx, file, proc, kernel)
}

func semanticAnalysisCapabilityUncached(ctx analysisContext, file parsedFile, proc sourceProcedure, kernel string) string {
	// These maps are immutable for one analysis context. Including the relevant
	// capability projection in the key prevents a stale value when resolution,
	// effect summaries, or array participant closure changes.
	switch kernel {
	case "array":
		parts := []string{
			fmt.Sprintf("%t", proc.ArrayParticipant),
			fmt.Sprintf("%t", proc.ArrayParticipantReady),
			semanticProcedureEffectFingerprint(proc.Effects),
			fmt.Sprintf("%#v", ctx.arrayModuleConfigurations[file.Path]),
		}
		keys := []string{arrayProcedureKey(proc), strings.ToLower(proc.Name)}
		if participantKey := ctx.arrayParticipantKeys[arrayProcedureKey(proc)]; participantKey != "" {
			keys = append(keys, participantKey)
		}
		for index := 0; index < proc.Calls.Len(); index++ {
			if call, ok := proc.Calls.At(index); ok {
				keys = append(keys, strings.ToLower(call.Callee.BaseName))
				for _, candidate := range call.Resolution.Candidates {
					keys = append(keys, strings.ToLower(candidate.QualifiedName))
				}
			}
		}
		var (
			returns, allocationGuards, safeArrayLengths, byRefAllocations map[string]string
			conditional, length, moduleAllocations                        map[string]string
			moduleEntryStates, byRefEntryStates                           map[string]string
			byRefConditions, participants                                 map[string]string
		)
		if index := ctx.arrayCapabilityIndex; index != nil {
			returns = index.returns
			allocationGuards = index.allocationGuards
			safeArrayLengths = index.safeArrayLengths
			byRefAllocations = index.byRefAllocations
			conditional = index.conditional
			length = index.length
			moduleAllocations = index.moduleAllocations
			moduleEntryStates = index.moduleEntryStates
			byRefEntryStates = index.byRefEntryStates
			byRefConditions = index.byRefConditions
			participants = index.participants
		}
		parts = append(parts,
			semanticArrayCapabilitySubset(returns, ctx.arrayReturns, keys),
			semanticArrayCapabilitySubset(allocationGuards, ctx.arrayAllocationGuards, keys),
			semanticArrayCapabilitySubset(safeArrayLengths, ctx.arraySafeArrayLengthGuards, keys),
			semanticArrayCapabilitySubset(byRefAllocations, ctx.arrayByRefAllocations, keys),
			semanticArrayCapabilitySubset(conditional, ctx.arrayByRefConditionalAllocations, keys),
			semanticArrayCapabilitySubset(length, ctx.arrayByRefLengthAllocations, keys),
			semanticArrayCapabilitySubset(moduleAllocations, ctx.arrayModuleAllocations, keys),
			semanticArrayCapabilitySubset(moduleEntryStates, ctx.arrayModuleEntryStates, keys),
			semanticArrayCapabilitySubset(byRefEntryStates, ctx.arrayByRefEntryStates, keys),
			semanticArrayCapabilitySubset(byRefConditions, ctx.arrayByRefEntryConditions, keys),
			semanticArrayCapabilitySubset(participants, ctx.arrayParticipants, keys),
		)
		return semanticquery.Hash(parts...)
	case "dataflow", "http":
		return semanticquery.Hash(semanticDataflowModuleFingerprint(file), semanticProcedureEffectFingerprint(proc.Effects))
	default:
		return semanticquery.Hash(semanticProcedureEffectFingerprint(proc.Effects))
	}
}

func semanticArrayCapabilitySubset[T any, M ~map[string]T](prepared map[string]string, fallback M, wanted []string) string {
	if prepared == nil {
		return semanticMapSubset(fallback, wanted)
	}
	return semanticNormalizedMapSubset(prepared, wanted)
}

func semanticStableJSON(value any) string {
	encoded, err := json.Marshal(value)
	if err != nil {
		return fmt.Sprintf("%#v", value)
	}
	return string(encoded)
}

func semanticNormalizedMapSubset(values map[string]string, wanted []string) string {
	if len(values) == 0 || len(wanted) == 0 {
		return ""
	}
	selected := make(map[string]struct{}, len(wanted))
	for _, key := range wanted {
		key = strings.ToLower(strings.TrimSpace(key))
		if _, ok := values[key]; ok {
			selected[key] = struct{}{}
		}
	}
	keys := make([]string, 0, len(selected))
	for key := range selected {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	parts := make([]string, 0, len(keys))
	for _, key := range keys {
		parts = append(parts, key+"="+values[key])
	}
	return strings.Join(parts, ";")
}

func semanticMapSubset[T any, M ~map[string]T](values M, wanted []string) string {
	if len(values) == 0 || len(wanted) == 0 {
		return ""
	}
	lookup := make(map[string]string, len(wanted))
	for _, key := range wanted {
		normalized := strings.ToLower(strings.TrimSpace(key))
		if normalized != "" {
			lookup[normalized] = key
		}
	}
	selected := make(map[string]T, len(lookup))
	for normalized, original := range lookup {
		// Production indexes are normalized at construction time, so this is
		// the common O(1) path. The exact lookup handles compatibility maps
		// whose keys retain source casing.
		if value, ok := values[original]; ok {
			selected[original] = value
			continue
		}
		if value, ok := values[normalized]; ok {
			selected[normalized] = value
			continue
		}
		found := ""
		// Focused callers may provide non-normalized maps. Scan only when a
		// requested leaf was not found directly; avoid the old full-map scan
		// for the normal analyzer path.
		for key := range values {
			if strings.EqualFold(key, normalized) {
				found = key
				break
			}
		}
		if found != "" {
			selected[found] = values[found]
		}
	}
	keys := make([]string, 0, len(selected))
	for key := range selected {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	parts := make([]string, 0, len(keys))
	for _, key := range keys {
		parts = append(parts, fmt.Sprintf("%s=%#v", key, selected[key]))
	}
	return strings.Join(parts, ";")
}
