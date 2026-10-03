// Package architecture builds a source-only, protocol-neutral report of a VBA
// project's structure, calls, reachability, state, and direct Excel effects.
package architecture

import (
	"context"
	"errors"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/callgraph"
	"github.com/harumiWeb/xlflow/internal/vba/hotspots"
	"github.com/harumiWeb/xlflow/internal/vba/proceduremetrics"
)

// ErrInvalidScope indicates that a requested display path or module does not
// select any source in the project.
var ErrInvalidScope = errors.New("invalid architecture report scope")

// ErrParse indicates that at least one discovered source file could not be
// parsed into a usable VBA document. CollectContext returns no partial report.
var ErrParse = errors.New("architecture source parse failed")

// Options narrows the displayed entities after the complete project has been
// resolved. Path and Module may be combined; when both are set they intersect.
type Options struct {
	Path   string
	Module string
}

// CollectContext creates a deterministic, source-only architecture report.
// The complete project is resolved before Path and Module are applied.
func CollectContext(ctx context.Context, root string, cfg config.Config, options Options) (Report, []map[string]any, error) {
	return collectContext(ctx, root, cfg, options)
}

// Report is the versioned JSON contract consumed by CLI renderers and other
// callers. All collection slices are emitted as [] when empty.
type Report struct {
	SchemaVersion             int                         `json:"schema_version"`
	Scope                     Scope                       `json:"scope"`
	Summary                   Summary                     `json:"summary"`
	ProjectSummary            Summary                     `json:"project_summary"`
	Modules                   []Module                    `json:"modules"`
	Procedures                []Procedure                 `json:"procedures"`
	EntryPoints               []EntryPoint                `json:"entry_points"`
	DynamicReferences         []DynamicReference          `json:"dynamic_references"`
	Dependencies              callgraph.DependencyResult  `json:"dependencies"`
	DependencyBoundaryNodeIDs []string                    `json:"dependency_boundary_node_ids"`
	Cycles                    []callgraph.CyclicComponent `json:"cycles"`
	Unreachable               []callgraph.Node            `json:"unreachable"`
	PossibleReachability      []callgraph.Node            `json:"possible_reachability"`
	Hotspots                  hotspots.Report             `json:"hotspots"`
	ModuleState               ModuleState                 `json:"module_state"`
	ExcelEffects              ExcelEffects                `json:"excel_effects"`
	ExternalDependencies      []callgraph.UncertainEdge   `json:"external_dependencies"`
	Uncertainty               Uncertainty                 `json:"uncertainty"`
}

// Scope records the normalized display filters. Empty fields mean the full
// project is displayed.
type Scope struct {
	Path   string `json:"path"`
	Module string `json:"module"`
}

// Summary contains whole-project or displayed-scope counts. In a filtered
// report ProjectSummary remains the unfiltered summary.
type Summary struct {
	Files                int `json:"files"`
	Modules              int `json:"modules"`
	Procedures           int `json:"procedures"`
	Calls                int `json:"calls"`
	ConfirmedCallEdges   int `json:"confirmed_call_edges"`
	UncertainCallEdges   int `json:"uncertain_call_edges"`
	DependencyCycles     int `json:"dependency_cycles"`
	ConfirmedReachable   int `json:"confirmed_reachable"`
	PossiblyReachable    int `json:"possibly_reachable"`
	Unreachable          int `json:"unreachable"`
	ExternalDependencies int `json:"external_dependencies"`
	ProcedureHotspots    int `json:"procedure_hotspots"`
	ModuleHotspots       int `json:"module_hotspots"`
}

// Module describes one source module. ID uses the canonical dependency graph
// module-node ID so report entities and dependency edges join directly.
type Module struct {
	ID             string `json:"id"`
	Name           string `json:"name"`
	Kind           string `json:"kind"`
	File           string `json:"file"`
	ProcedureCount int    `json:"procedure_count"`
}

// Procedure describes one procedure using its dependency node ID and module ID.
// CallgraphID retains the canonical reachability node ID.
type Procedure struct {
	ID               string                   `json:"id"`
	CallgraphID      string                   `json:"callgraph_id"`
	ModuleID         string                   `json:"module_id"`
	Name             string                   `json:"name"`
	QualifiedName    string                   `json:"qualified_name"`
	Kind             string                   `json:"kind"`
	Module           string                   `json:"module"`
	ModuleKind       string                   `json:"module_kind"`
	File             string                   `json:"file"`
	Location         callgraph.Location       `json:"location"`
	DeclarationRange SourceRange              `json:"declaration_range"`
	Visibility       string                   `json:"visibility"`
	Reachability     string                   `json:"reachability"`
	Metrics          proceduremetrics.Metrics `json:"metrics"`
}

// EntryPoint reports the resolution of a configured root request. Unresolved
// requests remain in the report with an empty NodeID and explicit Status.
type EntryPoint struct {
	Target     string   `json:"target"`
	Status     string   `json:"status"`
	Confidence string   `json:"confidence"`
	Reason     string   `json:"reason"`
	NodeID     string   `json:"node_id"`
	Candidates []string `json:"candidates"`
}

// ModuleState groups module-level state facts and counts for human renderers.
type ModuleState struct {
	Summary ModuleStateSummary `json:"summary"`
	Fields  []ModuleStateField `json:"fields"`
}

// ModuleStateSummary counts direct accesses recorded by the state collector.
type ModuleStateSummary struct {
	MutableStateReads     int `json:"mutable_state_reads"`
	MutableStateWrites    int `json:"mutable_state_writes"`
	MutableStateMutations int `json:"mutable_state_mutations"`
}

// ModuleStateField is a protocol-neutral projection of a collected field.
// Reader, writer, and mutator values use Report Procedure IDs.
type ModuleStateField struct {
	ID              string   `json:"id"`
	File            string   `json:"file"`
	ModuleID        string   `json:"module_id"`
	Module          string   `json:"module"`
	Name            string   `json:"name"`
	Type            string   `json:"type"`
	Visibility      string   `json:"visibility"`
	Kind            string   `json:"kind"`
	Scope           string   `json:"scope"`
	DeclarationLine int      `json:"declaration_line"`
	IsObject        bool     `json:"is_object"`
	IsCollection    bool     `json:"is_collection"`
	IsExcel         bool     `json:"is_excel"`
	Readers         []string `json:"readers"`
	Writers         []string `json:"writers"`
	Mutators        []string `json:"mutators"`
}

// ExcelEffects contains direct evidence only. It deliberately excludes
// filesystem, process, and error-handling effects and does not assert runtime
// certainty.
type ExcelEffects struct {
	DirectEffectCount int              `json:"direct_effect_count"`
	Evidence          []EffectEvidence `json:"evidence"`
}

// DynamicReference is the stable JSON projection of a callback-like reference
// retained by the canonical calls extractor.
type DynamicReference struct {
	File          string      `json:"file"`
	Module        string      `json:"module"`
	CallerID      string      `json:"caller_id"`
	API           string      `json:"api"`
	ArgumentIndex int         `json:"argument_index"`
	ArgumentName  string      `json:"argument_name"`
	Expression    string      `json:"expression"`
	Target        string      `json:"target"`
	Kind          string      `json:"kind"`
	Range         SourceRange `json:"range"`
}

// SourceRange retains line, column, and byte offsets from canonical VBA IR.
type SourceRange struct {
	StartLine   int `json:"start_line"`
	StartColumn int `json:"start_column"`
	EndLine     int `json:"end_line"`
	EndColumn   int `json:"end_column"`
	StartByte   int `json:"start_byte"`
	EndByte     int `json:"end_byte"`
}

// EffectEvidence is one direct recognized Excel/workbook/UI effect.
type EffectEvidence struct {
	ProcedureID   string      `json:"procedure_id"`
	ProcedureName string      `json:"procedure_name"`
	Kind          string      `json:"kind"`
	Target        string      `json:"target"`
	Value         string      `json:"value"`
	File          string      `json:"file"`
	Line          int         `json:"line"`
	Evidence      string      `json:"evidence"`
	Range         SourceRange `json:"range"`
	StatementID   int         `json:"statement_id"`
	CallID        int         `json:"call_id"`
}

// Uncertainty summarizes incomplete or ambiguous static resolution. Its
// counters describe analysis evidence, not runtime behavior.
type Uncertainty struct {
	UnresolvedCallCount  int                       `json:"unresolved_call_count"`
	AmbiguousCallCount   int                       `json:"ambiguous_call_count"`
	DynamicCallCount     int                       `json:"dynamic_call_count"`
	TypeDatabaseLoaded   bool                      `json:"type_database_loaded"`
	TypeDatabaseComplete bool                      `json:"type_database_complete"`
	Reasons              []string                  `json:"reasons"`
	Evidence             []CallUncertaintyEvidence `json:"evidence"`
}

// CallUncertaintyEvidence preserves the source and candidate evidence for one
// unresolved, ambiguous, external, member, builtin-like, or dynamic call.
type CallUncertaintyEvidence struct {
	Status     string                 `json:"status"`
	CallerID   string                 `json:"caller_id"`
	File       string                 `json:"file"`
	Module     string                 `json:"module"`
	Target     string                 `json:"target"`
	Candidates []UncertaintyCandidate `json:"candidates"`
	Range      SourceRange            `json:"range"`
}

// UncertaintyCandidate is a snake_case projection of a calls resolver candidate.
type UncertaintyCandidate struct {
	QualifiedName string `json:"qualified_name"`
	Kind          string `json:"kind"`
	File          string `json:"file"`
	Line          int    `json:"line"`
}
