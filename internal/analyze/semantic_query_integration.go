package analyze

import (
	"context"
	"encoding/json"
	"fmt"
	"sort"
	"strings"
	"sync"

	"github.com/harumiWeb/xlflow/internal/analyze/semanticquery"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
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
	return fmt.Sprintf("%s::%s::%s::%d", file.Path, file.Module, proc.Name, proc.Index)
}

func semanticProcedureQueryKey(a Analyzer, file parsedFile, proc sourceProcedure, kernel string, plan procedureAnalysisPlan, capability string) semanticquery.Key {
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
	configBytes, err := json.Marshal(a.Config.Analyze)
	if err != nil {
		configBytes = []byte(fmt.Sprintf("%#v", a.Config.Analyze))
	}
	fingerprint := semanticquery.Hash(
		a.RootDir,
		file.Path,
		file.Module,
		file.ModuleKind,
		semanticModuleFactsFingerprint(file),
		proc.Name,
		proc.Kind,
		string(proc.ProcedureKind),
		proc.ReturnType,
		fmt.Sprintf("%d", proc.Index),
		fmt.Sprintf("%d:%d:%d:%d", proc.StartLine, proc.EndLine, proc.StartByte, proc.EndByte),
		fmt.Sprintf("option-base:%d:%t", file.ArrayOptionBase, file.ArrayOptionBaseSet),
		fmt.Sprintf("%#v", file.ArrayIntegerModuleConstants),
		fmt.Sprintf("%#v", file.RangeValueModuleConstants),
		fmt.Sprintf("%#v", file.ConstantValues),
		fmt.Sprintf("%#v", file.DataFlowModuleBindings),
		string(body),
		fmt.Sprintf("%#v", proc.Features),
		semanticStableJSON(proc.Effects),
	)
	return semanticquery.Key{
		Procedure:   semanticProcedureIdentity(file, proc),
		Fingerprint: fingerprint,
		Kernel:      kernel + ":" + fmt.Sprintf("%#v", plan),
		Config:      semanticquery.Hash(string(configBytes)),
		Capability:  semanticquery.Hash(capability, semanticAnalyzerCapability(a), semanticTypeDBFingerprint(a)),
	}
}

func semanticQueryDependencies(key semanticquery.Key, file parsedFile, proc sourceProcedure) []semanticquery.Key {
	dependencies := []semanticquery.Key{
		{Procedure: key.Procedure, Fingerprint: key.Fingerprint, Kernel: "source"},
		{Procedure: key.Procedure, Fingerprint: key.Fingerprint, Kernel: "config", Config: key.Config},
		{Procedure: key.Procedure, Fingerprint: key.Fingerprint, Kernel: "capability", Capability: key.Capability},
		{Procedure: file.Path, Fingerprint: semanticquery.Hash(semanticModuleFactsFingerprint(file)), Kernel: "module"},
		{Procedure: key.Procedure, Fingerprint: semanticStableJSON(proc.Effects), Kernel: "effect-summary"},
	}
	for index := 0; index < proc.Calls.Len(); index++ {
		call, ok := proc.Calls.At(index)
		if !ok {
			continue
		}
		resolutionFingerprint := semanticStableJSON(call.Resolution)
		if len(call.Resolution.Candidates) == 0 {
			dependencies = append(dependencies, semanticquery.Key{
				Procedure:   file.Path,
				Fingerprint: semanticquery.Hash(fmt.Sprintf("call:%d", call.ID), resolutionFingerprint),
				Kernel:      "call-resolution",
			})
			continue
		}
		for _, candidate := range call.Resolution.Candidates {
			dependencies = append(dependencies, semanticquery.Key{
				Procedure:   fmt.Sprintf("%s::%s::%s::%d", candidate.File, candidate.QualifiedName, candidate.Kind, candidate.Line),
				Fingerprint: semanticquery.Hash(resolutionFingerprint),
				Kernel:      "call-resolution",
			})
		}
	}
	return dependencies
}

func semanticAnalyzerCapability(a Analyzer) string {
	return semanticquery.Hash(fmt.Sprintf("%#v", a.visibleConstants), fmt.Sprintf("%#v", a.visibleConstantValues))
}

func semanticTypeDBFingerprint(a Analyzer) string {
	if a.typeDB == nil {
		return ""
	}
	if value, ok := semanticTypeDBFingerprints.Load(a.typeDB); ok {
		return value.(string)
	}
	fingerprint := semanticStableJSON(a.typeDB.AllConstantsList())
	actual, _ := semanticTypeDBFingerprints.LoadOrStore(a.typeDB, fingerprint)
	return actual.(string)
}

var semanticTypeDBFingerprints sync.Map // map[*vbadb.DB]string

func semanticModuleFactsFingerprint(file parsedFile) string {
	if file.moduleFactsFingerprint != "" {
		return file.moduleFactsFingerprint
	}
	facts := file.ModuleFacts
	if facts == nil {
		return ""
	}
	return semanticquery.Hash(
		fmt.Sprintf("%#v", facts.moduleDeclarations),
		fmt.Sprintf("%#v", facts.constants),
		fmt.Sprintf("%#v", facts.moduleConstantNames),
		fmt.Sprintf("%#v", facts.procedureNames),
		fmt.Sprintf("%#v", facts.privateModule),
		fmt.Sprintf("%#v", facts.arrayOperationsByName),
		fmt.Sprintf("%#v", facts.arrayOperationsByLine),
		fmt.Sprintf("%#v", facts.lineFacts),
	)
}

func semanticAnalysisCapability(ctx analysisContext, file parsedFile, proc sourceProcedure, kernel string) string {
	// These maps are immutable for one analysis context. Including the relevant
	// capability projection in the key prevents a stale value when resolution,
	// effect summaries, or array participant closure changes.
	switch kernel {
	case "array":
		parts := []string{
			fmt.Sprintf("%t", proc.ArrayParticipant),
			fmt.Sprintf("%t", proc.ArrayParticipantReady),
			semanticStableJSON(proc.Effects),
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
		parts = append(parts,
			semanticMapSubset(ctx.arrayReturns, keys),
			semanticMapSubset(ctx.arrayAllocationGuards, keys),
			semanticMapSubset(ctx.arrayByRefAllocations, keys),
			semanticMapSubset(ctx.arrayByRefConditionalAllocations, keys),
			semanticMapSubset(ctx.arrayByRefLengthAllocations, keys),
			semanticMapSubset(ctx.arrayModuleAllocations, keys),
			semanticMapSubset(ctx.arrayModuleEntryStates, keys),
			semanticMapSubset(ctx.arrayByRefEntryStates, keys),
			semanticMapSubset(ctx.arrayByRefEntryConditions, keys),
			semanticMapSubset(ctx.arrayParticipants, keys),
		)
		return semanticquery.Hash(parts...)
	case "dataflow", "http":
		return semanticquery.Hash(fmt.Sprintf("%#v", file.DataFlowModuleBindings), semanticStableJSON(proc.Effects))
	default:
		return semanticquery.Hash(semanticStableJSON(proc.Effects))
	}
}

func semanticStableJSON(value any) string {
	encoded, err := json.Marshal(value)
	if err != nil {
		return fmt.Sprintf("%#v", value)
	}
	return string(encoded)
}

func semanticMapSubset[T any, M ~map[string]T](values M, wanted []string) string {
	if len(values) == 0 || len(wanted) == 0 {
		return ""
	}
	lookup := make(map[string]struct{}, len(wanted))
	for _, key := range wanted {
		key = strings.ToLower(strings.TrimSpace(key))
		if key != "" {
			lookup[key] = struct{}{}
		}
	}
	keys := make([]string, 0, len(values))
	for key := range values {
		if _, ok := lookup[strings.ToLower(key)]; ok {
			keys = append(keys, key)
		}
	}
	sort.Strings(keys)
	parts := make([]string, 0, len(keys))
	for _, key := range keys {
		parts = append(parts, fmt.Sprintf("%s=%#v", key, values[key]))
	}
	return strings.Join(parts, ";")
}
