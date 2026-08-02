// Package effects computes deterministic, protocol-neutral procedure effect
// summaries from resolved procedure IR and its control-flow graphs.
package effects

import (
	"path/filepath"
	"sort"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type EffectKind string

const (
	WritesCells        EffectKind = "writes_cells"
	ChangesWorkbook    EffectKind = "changes_workbook"
	OpensWorkbook      EffectKind = "opens_workbook"
	ClosesWorkbook     EffectKind = "closes_workbook"
	DisablesEvents     EffectKind = "disables_events"
	RestoresEvents     EffectKind = "restores_events"
	ChangesCalculation EffectKind = "changes_calculation"
	Recalculates       EffectKind = "recalculates"
	ChangesSelection   EffectKind = "changes_selection"
	ChangesControls    EffectKind = "changes_controls"
	ShowsDialog        EffectKind = "shows_dialog"
	LaunchesProcess    EffectKind = "launches_process"
	SuppressesErrors   EffectKind = "suppresses_errors"
	RaisesError        EffectKind = "raises_error"
	// ChangesApplicationState and RestoresApplicationState preserve the
	// property-neutral state contract needed by consumers such as VBA203.
	// Target identifies the exact Application property.
	ChangesApplicationState  EffectKind = "changes_application_state"
	RestoresApplicationState EffectKind = "restores_application_state"
)

// ProcedureIdentity is stable across input ordering. DeclarationLine
// disambiguates procedures with otherwise identical source identities.
type ProcedureIdentity struct {
	File            string
	Module          string
	ModuleKind      string
	Name            string
	QualifiedName   string
	Kind            procedureir.ProcedureKind
	Visibility      string
	DeclarationLine int
}

func (id ProcedureIdentity) Key() string {
	return strings.Join([]string{
		canonicalPath(id.File), strings.ToLower(id.Module),
		strings.ToLower(id.QualifiedName), string(id.Kind),
		decimal(id.DeclarationLine),
	}, "\x00")
}

type Evidence struct {
	Effect      EffectKind
	Origin      ProcedureIdentity
	Range       vbaast.Range
	StatementID int
	CallID      int
	Target      string
	Value       string
}

type UncertaintyKind string

const (
	UncertaintyAmbiguous  UncertaintyKind = "ambiguous"
	UncertaintyUnresolved UncertaintyKind = "unresolved"
	UncertaintyExternal   UncertaintyKind = "external"
	UncertaintyDynamic    UncertaintyKind = "dynamic"
)

type CallUncertainty struct {
	Kind        UncertaintyKind
	Origin      ProcedureIdentity
	Range       vbaast.Range
	StatementID int
	CallID      int
	Callee      string
}

type ProcedureSummary struct {
	Identity              ProcedureIdentity
	Direct                []Evidence
	Propagated            []Evidence
	DirectUncertainty     []CallUncertainty
	PropagatedUncertainty []CallUncertainty
}

func (s ProcedureSummary) Has(kind EffectKind) bool {
	for _, evidence := range s.Direct {
		if evidence.Effect == kind {
			return true
		}
	}
	for _, evidence := range s.Propagated {
		if evidence.Effect == kind {
			return true
		}
	}
	return false
}

type ProjectSummary struct {
	procedures []ProcedureSummary
	byKey      map[string]int
}

func (p ProjectSummary) Lookup(id ProcedureIdentity) (ProcedureSummary, bool) {
	i, ok := p.byKey[id.Key()]
	if !ok {
		return ProcedureSummary{}, false
	}
	return cloneProcedureSummary(p.procedures[i]), true
}

// All returns a defensive copy in deterministic procedure order.
func (p ProjectSummary) All() []ProcedureSummary {
	out := make([]ProcedureSummary, len(p.procedures))
	for i := range p.procedures {
		out[i] = cloneProcedureSummary(p.procedures[i])
	}
	return out
}

// Document pairs already-resolved IR with the CFG built from that exact IR.
type Document struct {
	IR  procedureir.DocumentIR
	CFG cfg.Document
}

func identity(doc procedureir.DocumentIR, proc procedureir.ProcedureIR) ProcedureIdentity {
	return ProcedureIdentity{
		File: canonicalPath(doc.Path), Module: doc.ModuleName, ModuleKind: doc.ModuleKind,
		Name: proc.Symbol.Name, QualifiedName: proc.Symbol.QualifiedName, Kind: proc.Symbol.Kind,
		Visibility:      proc.Symbol.Visibility,
		DeclarationLine: proc.Symbol.DeclarationRange.StartLine,
	}
}

func cloneProcedureSummary(in ProcedureSummary) ProcedureSummary {
	in.Direct = append([]Evidence(nil), in.Direct...)
	in.Propagated = append([]Evidence(nil), in.Propagated...)
	in.DirectUncertainty = append([]CallUncertainty(nil), in.DirectUncertainty...)
	in.PropagatedUncertainty = append([]CallUncertainty(nil), in.PropagatedUncertainty...)
	return in
}

func canonicalPath(path string) string {
	return filepath.ToSlash(filepath.Clean(path))
}

func decimal(value int) string {
	if value == 0 {
		return "0"
	}
	const digits = "0123456789"
	var buf [24]byte
	i := len(buf)
	for value > 0 {
		i--
		buf[i] = digits[value%10]
		value /= 10
	}
	return string(buf[i:])
}

func sortSummaries(items []ProcedureSummary) {
	sort.Slice(items, func(i, j int) bool { return items[i].Identity.Key() < items[j].Identity.Key() })
	for i := range items {
		sort.Slice(items[i].Direct, func(a, b int) bool { return evidenceKey(items[i].Direct[a]) < evidenceKey(items[i].Direct[b]) })
		sort.Slice(items[i].Propagated, func(a, b int) bool { return evidenceKey(items[i].Propagated[a]) < evidenceKey(items[i].Propagated[b]) })
		sort.Slice(items[i].DirectUncertainty, func(a, b int) bool {
			return uncertaintyKey(items[i].DirectUncertainty[a]) < uncertaintyKey(items[i].DirectUncertainty[b])
		})
		sort.Slice(items[i].PropagatedUncertainty, func(a, b int) bool {
			return uncertaintyKey(items[i].PropagatedUncertainty[a]) < uncertaintyKey(items[i].PropagatedUncertainty[b])
		})
	}
}

func evidenceKey(e Evidence) string {
	return strings.Join([]string{string(e.Effect), e.Origin.Key(), decimal(e.Range.StartByte), decimal(e.Range.EndByte), decimal(e.StatementID), decimal(e.CallID), strings.ToLower(e.Target), strings.ToLower(e.Value)}, "\x00")
}

func uncertaintyKey(u CallUncertainty) string {
	return strings.Join([]string{string(u.Kind), u.Origin.Key(), decimal(u.Range.StartByte), decimal(u.Range.EndByte), decimal(u.StatementID), decimal(u.CallID), strings.ToLower(u.Callee)}, "\x00")
}
