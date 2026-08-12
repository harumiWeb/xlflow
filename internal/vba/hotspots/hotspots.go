// Package hotspots computes deterministic, project-relative architectural
// hotspot scores. It intentionally depends only on scalar facts supplied by
// callers; parsing, diagnostics, and configuration remain outside this layer.
package hotspots

import (
	"math"
	"sort"
	"strings"
)

const (
	SchemaVersion = 1
	ScoreModel    = "percentile_equal_weight_v1"
)

// SignalName is a stable raw-signal identifier. Additive signals are
// backwards compatible; changing a signal's meaning requires a new model.
type SignalName string

const (
	SignalComplexity       SignalName = "complexity"
	SignalComplexityMax    SignalName = "complexity_max"
	SignalCallFanIn        SignalName = "call_fan_in"
	SignalCallFanOut       SignalName = "call_fan_out"
	SignalAffectedModules  SignalName = "affected_module_count"
	SignalExcelEffects     SignalName = "excel_effect_count"
	SignalMutableReads     SignalName = "mutable_state_reads"
	SignalMutableWrites    SignalName = "mutable_state_writes"
	SignalMutableMutations SignalName = "mutable_state_mutations"
	SignalCycleCount       SignalName = "cycle_count"
	SignalExternalDeps     SignalName = "external_dependency_count"
	SignalErrorHandling    SignalName = "error_handling_count"
	SignalResourceOwned    SignalName = "resource_ownership_count"
	SignalPublicProcedures SignalName = "public_procedure_count"
)

// UncertaintyName identifies relationships retained for auditability but
// intentionally excluded from the composite score.
type UncertaintyName string

const (
	UncertaintyAmbiguous  UncertaintyName = "ambiguous_call_count"
	UncertaintyUnresolved UncertaintyName = "unresolved_call_count"
	UncertaintyDynamic    UncertaintyName = "dynamic_call_count"
)

// ProcedureSignals and ModuleSignals provide typed construction helpers while
// the wire representation remains an extensible signal map.
type ProcedureSignals struct {
	Complexity, CallFanIn, CallFanOut, AffectedModules                 int
	ExcelEffects, MutableReads, MutableWrites, MutableMutations        int
	CycleCount, ExternalDependencies, ErrorHandling, ResourceOwnership int
}

func (s ProcedureSignals) Map() map[string]int {
	return map[string]int{
		string(SignalComplexity): s.Complexity, string(SignalCallFanIn): s.CallFanIn, string(SignalCallFanOut): s.CallFanOut,
		string(SignalAffectedModules): s.AffectedModules, string(SignalExcelEffects): s.ExcelEffects,
		string(SignalMutableReads): s.MutableReads, string(SignalMutableWrites): s.MutableWrites, string(SignalMutableMutations): s.MutableMutations,
		string(SignalCycleCount): s.CycleCount, string(SignalExternalDeps): s.ExternalDependencies, string(SignalErrorHandling): s.ErrorHandling,
		string(SignalResourceOwned): s.ResourceOwnership,
	}
}

type ModuleSignals struct {
	ComplexityTotal, ComplexityMax, CallFanIn, CallFanOut, AffectedModules               int
	ExcelEffects, MutableReads, MutableWrites, MutableMutations                          int
	CycleCount, ExternalDependencies, ErrorHandling, ResourceOwnership, PublicProcedures int
}

func (s ModuleSignals) Map() map[string]int {
	return map[string]int{
		string(SignalComplexity): s.ComplexityTotal, string(SignalComplexityMax): s.ComplexityMax,
		string(SignalCallFanIn): s.CallFanIn, string(SignalCallFanOut): s.CallFanOut, string(SignalAffectedModules): s.AffectedModules,
		string(SignalExcelEffects): s.ExcelEffects, string(SignalMutableReads): s.MutableReads, string(SignalMutableWrites): s.MutableWrites,
		string(SignalMutableMutations): s.MutableMutations, string(SignalCycleCount): s.CycleCount, string(SignalExternalDeps): s.ExternalDependencies,
		string(SignalErrorHandling): s.ErrorHandling, string(SignalResourceOwned): s.ResourceOwnership, string(SignalPublicProcedures): s.PublicProcedures,
	}
}

// Input is a project-relative entity and its raw scalar facts. Kind must be
// "procedure" or "module". Missing signals are treated as zero.
type Input struct {
	ID              string         `json:"id"`
	Kind            string         `json:"kind"`
	File            string         `json:"file"`
	Module          string         `json:"module"`
	ModuleKind      string         `json:"module_kind,omitempty"`
	Name            string         `json:"name"`
	ProcedureKind   string         `json:"procedure_kind,omitempty"`
	DeclarationByte int            `json:"-"`
	Line            int            `json:"line,omitempty"`
	RawSignals      map[string]int `json:"raw_signals"`
	Uncertainty     map[string]int `json:"uncertainty,omitempty"`
}

// Entity is a ranked hotspot with both raw and normalized contributing facts.
type Entity struct {
	ID                string             `json:"id"`
	Kind              string             `json:"kind"`
	File              string             `json:"file"`
	Module            string             `json:"module,omitempty"`
	ModuleKind        string             `json:"module_kind,omitempty"`
	Name              string             `json:"name"`
	ProcedureKind     string             `json:"procedure_kind,omitempty"`
	Rank              int                `json:"rank"`
	Score             float64            `json:"score"`
	ScoreModel        string             `json:"score_model"`
	ActiveSignalCount int                `json:"active_signal_count"`
	RawSignals        map[string]int     `json:"raw_signals"`
	NormalizedSignals map[string]float64 `json:"normalized_signals"`
	SelectedBy        *Selection         `json:"selected_by,omitempty"`
	DeclarationByte   int                `json:"-"`
	Line              int                `json:"line,omitempty"`
	Uncertainty       map[string]int     `json:"uncertainty,omitempty"`
}

// Selection records why an opt-in finding was selected.
type Selection struct {
	TopN      bool `json:"top_n,omitempty"`
	Threshold bool `json:"threshold,omitempty"`
}

// Report is the versioned hotspot projection embedded in the metrics output.
type Report struct {
	SchemaVersion int      `json:"schema_version"`
	ScoreModel    string   `json:"score_model"`
	Procedures    []Entity `json:"procedures"`
	Modules       []Entity `json:"modules"`
}

// Rank computes percentile-normalized, equal-weight scores and stable ranks.
// The input is copied and may be in arbitrary order.
func Rank(inputs []Input) []Entity {
	entities := make([]Entity, 0, len(inputs))
	for _, in := range inputs {
		raw := cloneInts(in.RawSignals)
		entities = append(entities, Entity{ID: in.ID, Kind: in.Kind, File: in.File, Module: in.Module, ModuleKind: in.ModuleKind, Name: in.Name, ProcedureKind: in.ProcedureKind, DeclarationByte: in.DeclarationByte, Line: in.Line, ScoreModel: ScoreModel, RawSignals: raw, Uncertainty: cloneInts(in.Uncertainty), NormalizedSignals: map[string]float64{}})
	}
	keys := signalKeys(entities)
	for _, key := range keys {
		values := make([]int, len(entities))
		for i := range entities {
			values[i] = entities[i].RawSignals[key]
		}
		ranks := Percentiles(values)
		active := signalVaries(values)
		for i := range entities {
			if active {
				entities[i].NormalizedSignals[key] = ranks[i]
				entities[i].ActiveSignalCount++
			}
		}
	}
	for i := range entities {
		if entities[i].ActiveSignalCount == 0 {
			entities[i].Score = 0
			continue
		}
		total := 0.0
		for _, value := range entities[i].NormalizedSignals {
			total += value
		}
		entities[i].Score = round2(total / float64(entities[i].ActiveSignalCount))
	}
	sort.SliceStable(entities, func(i, j int) bool { return less(entities[i], entities[j]) })
	for i := range entities {
		entities[i].Rank = i + 1
	}
	return entities
}

// BuildReport ranks procedure and module inputs in independent cohorts.
func BuildReport(procedures, modules []Input) Report {
	return Report{SchemaVersion: SchemaVersion, ScoreModel: ScoreModel, Procedures: Rank(procedures), Modules: Rank(modules)}
}

// Select returns the union of top-N and threshold selections. It preserves
// rank order and marks the selection reasons on each entity.
func Select(entities []Entity, topN int, threshold float64) []Entity {
	if topN < 0 {
		topN = 0
	}
	selected := make([]Entity, 0)
	for i, entity := range entities {
		reason := Selection{TopN: topN > 0 && i < topN, Threshold: threshold > 0 && entity.Score >= threshold}
		if !reason.TopN && !reason.Threshold {
			continue
		}
		selection := reason
		entity.SelectedBy = &selection
		selected = append(selected, entity)
	}
	return selected
}

// Percentiles returns average-rank percentiles in input order. Constant and
// singleton cohorts intentionally return zero for every value.
func Percentiles(values []int) []float64 {
	out := make([]float64, len(values))
	if len(values) < 2 || !signalVaries(values) {
		return out
	}
	order := make([]int, len(values))
	for i := range order {
		order[i] = i
	}
	sort.SliceStable(order, func(i, j int) bool {
		if values[order[i]] != values[order[j]] {
			return values[order[i]] < values[order[j]]
		}
		return order[i] < order[j]
	})
	for start := 0; start < len(order); {
		end := start + 1
		for end < len(order) && values[order[end]] == values[order[start]] {
			end++
		}
		averageRank := (float64(start+1) + float64(end)) / 2
		percentile := 100 * (averageRank - 1) / float64(len(values)-1)
		for i := start; i < end; i++ {
			out[order[i]] = percentile
		}
		start = end
	}
	return out
}

func signalKeys(entities []Entity) []string {
	set := map[string]bool{}
	for _, entity := range entities {
		for key := range entity.RawSignals {
			if strings.TrimSpace(key) != "" {
				set[key] = true
			}
		}
	}
	keys := make([]string, 0, len(set))
	for key := range set {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	return keys
}
func signalVaries(values []int) bool {
	if len(values) < 2 {
		return false
	}
	first := values[0]
	for _, value := range values[1:] {
		if value != first {
			return true
		}
	}
	return false
}
func cloneInts(in map[string]int) map[string]int {
	out := make(map[string]int, len(in))
	for key, value := range in {
		out[key] = value
	}
	return out
}
func less(a, b Entity) bool {
	if math.Abs(a.Score-b.Score) > 1e-9 {
		return a.Score > b.Score
	}
	if a.File != b.File {
		return a.File < b.File
	}
	if a.DeclarationByte != b.DeclarationByte {
		return a.DeclarationByte < b.DeclarationByte
	}
	if a.Module != b.Module {
		return a.Module < b.Module
	}
	if a.Name != b.Name {
		return a.Name < b.Name
	}
	if a.ProcedureKind != b.ProcedureKind {
		return a.ProcedureKind < b.ProcedureKind
	}
	return a.ID < b.ID
}
func round2(value float64) float64 { return math.Round(value*100) / 100 }
