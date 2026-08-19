package analyze

import (
	"sort"
	"strings"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/effects"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type eventFindingCandidate struct {
	line           int
	statementID    int
	boundary       string
	category       string
	classification string
	uncertainty    string
	effect         effects.Evidence
	same           bool
}

// VBA220 deliberately reports only the initial event surface. The IR records
// broader event metadata, but treating every document procedure as a handler
// would make this safety rule noisier than its published contract.
func eventHandlerKind(proc sourceProcedure) string {
	name := strings.ToLower(proc.Name)
	if proc.ModuleKind == "document" {
		switch name {
		case "worksheet_change", "workbook_sheetchange":
			return "cell"
		case "worksheet_calculate":
			return "calculation"
		case "worksheet_selectionchange":
			return "selection"
		case "workbook_open":
			return "open"
		case "workbook_beforeclose":
			return "close"
		}
	}
	if proc.ModuleKind == "form" {
		if strings.HasSuffix(name, "_change") && !strings.HasPrefix(name, "test") {
			return "control-change"
		}
		if strings.HasSuffix(name, "_click") && !strings.HasPrefix(name, "test") {
			return "control-click"
		}
	}
	return ""
}

func (a Analyzer) eventHandlerReentryFindings(file parsedFile, proc sourceProcedure, project effects.ProjectSummary) []Finding {
	handler := eventHandlerKind(proc)
	if handler == "" || proc.Effects == nil {
		return nil
	}
	candidates := map[string]eventFindingCandidate{}
	record := func(line int, statementID int, boundary string, effect effects.Evidence, uncertainty string) {
		if uncertainty == "" && a.eventSafeProcedures[effect.Origin.Key()] {
			return
		}
		category, same, ok := eventEffectRisk(proc, handler, effect)
		if uncertainty != "" {
			category, same, ok = "unknown", false, true
		}
		if !ok || (category != "control" && eventGuardedAt(proc, statementID)) {
			return
		}
		classification := "broader event-chain risk"
		if same {
			classification = "same-event recursion hazard"
		}
		candidate := eventFindingCandidate{
			line: line, statementID: statementID, boundary: boundary, category: category,
			classification: classification, uncertainty: uncertainty,
			effect: effect, same: same,
		}
		key := strings.Join([]string{strconvItoa(line), boundary}, ":")
		if previous, exists := candidates[key]; exists && !eventFindingCandidatePreferred(candidate, previous) {
			return
		}
		candidates[key] = candidate
	}

	for _, evidence := range proc.Effects.Direct {
		record(evidence.Range.StartLine, evidence.StatementID, eventStatementBoundary(evidence.StatementID), evidence, "")
	}
	for _, uncertainty := range proc.Effects.DirectUncertainty {
		record(uncertainty.Range.StartLine, uncertainty.StatementID, eventStatementBoundary(uncertainty.StatementID), effects.Evidence{}, string(uncertainty.Kind))
	}
	for _, call := range proc.Calls {
		if call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
			continue
		}
		summary, ok := summaryForCandidate(project, call.Resolution.Candidates[0])
		if !ok {
			continue
		}
		safeCallee := a.eventSafeProcedures[summary.Identity.Key()]
		boundary := eventCallBoundary(call)
		for _, evidence := range append(append([]effects.Evidence{}, summary.Direct...), summary.Propagated...) {
			if safeCallee {
				continue
			}
			record(call.Range.StartLine, call.StatementID, boundary, evidence, "")
		}
		for _, uncertainty := range append(append([]effects.CallUncertainty{}, summary.DirectUncertainty...), summary.PropagatedUncertainty...) {
			record(call.Range.StartLine, call.StatementID, boundary, effects.Evidence{}, string(uncertainty.Kind))
		}
	}

	ordered := make([]eventFindingCandidate, 0, len(candidates))
	for _, candidate := range candidates {
		ordered = append(ordered, candidate)
	}
	sort.SliceStable(ordered, func(i, j int) bool {
		if ordered[i].line != ordered[j].line {
			return ordered[i].line < ordered[j].line
		}
		if ordered[i].statementID != ordered[j].statementID {
			return ordered[i].statementID < ordered[j].statementID
		}
		if ordered[i].boundary != ordered[j].boundary {
			return ordered[i].boundary < ordered[j].boundary
		}
		return eventFindingCandidateKey(ordered[i]) < eventFindingCandidateKey(ordered[j])
	})

	out := make([]Finding, 0, len(ordered))
	for _, candidate := range ordered {
		message := "Event handler " + proc.Name + " has a " + candidate.classification + "."
		reason := "This handler can trigger " + candidate.category + " event processing"
		if candidate.uncertainty != "" {
			message = "Event handler " + proc.Name + " reaches an " + candidate.uncertainty + " call that may trigger an event."
			reason = "The call cannot be resolved safely, so its event-triggering effects are uncertain."
		} else if candidate.effect.Origin.QualifiedName != "" && !strings.EqualFold(candidate.effect.Origin.QualifiedName, proc.Effects.Identity.QualifiedName) {
			reason += " through " + candidate.effect.Origin.QualifiedName + "."
		} else {
			reason += "."
		}
		out = append(out, a.simpleFinding(file, proc, candidate.line, "VBA220", "warning", message, reason, "Disable Application.EnableEvents around Excel event-triggering work and restore it on every exit; use a re-entry guard for UserForm controls."))
	}
	return out
}

func eventFindingCandidatePreferred(candidate, previous eventFindingCandidate) bool {
	candidateRank, previousRank := eventFindingCandidateRank(candidate), eventFindingCandidateRank(previous)
	if candidateRank != previousRank {
		return candidateRank > previousRank
	}
	return eventFindingCandidateKey(candidate) < eventFindingCandidateKey(previous)
}

func eventFindingCandidateRank(candidate eventFindingCandidate) int {
	if candidate.uncertainty != "" {
		return 1
	}
	if candidate.same {
		return 3
	}
	return 2
}

func eventFindingCandidateKey(candidate eventFindingCandidate) string {
	return strings.Join([]string{candidate.boundary, candidate.classification, candidate.category, candidate.uncertainty, candidate.effect.Origin.QualifiedName}, ":")
}

func eventStatementBoundary(statementID int) string {
	return "statement:" + strconvItoa(statementID)
}

func eventCallBoundary(call procedureir.CallSite) string {
	if call.ID != 0 {
		return "call:" + strconvItoa(call.ID)
	}
	return strings.Join([]string{
		"call-range",
		strconvItoa(call.Range.StartByte),
		strconvItoa(call.Range.EndByte),
	}, ":")
}

func eventEffectRisk(proc sourceProcedure, handler string, evidence effects.Evidence) (string, bool, bool) {
	category := ""
	switch evidence.Effect {
	case effects.WritesCells:
		category = "cell"
	case effects.Recalculates:
		category = "calculation"
	case effects.ChangesSelection:
		category = "selection"
	case effects.OpensWorkbook:
		category = "open"
	case effects.ClosesWorkbook:
		category = "close"
	case effects.ChangesControls:
		category = "control"
	case effects.ChangesWorkbook:
		lower := strings.ToLower(evidence.Target)
		if strings.Contains(lower, "range(") || strings.Contains(lower, "cells(") || strings.Contains(lower, "rows(") || strings.Contains(lower, "columns(") || strings.HasSuffix(lower, ".value") || strings.HasSuffix(lower, ".value2") || strings.HasSuffix(lower, ".formula") || strings.HasSuffix(lower, ".formular1c1") {
			return "", false, false
		}
		category = "workbook structure"
	default:
		return "", false, false
	}
	if handler == "control-change" && category == "control" {
		control := strings.TrimSuffix(strings.ToLower(proc.Name), "_change")
		return category, strings.EqualFold(eventControlTargetName(evidence.Target), control), true
	}
	return category, handler == category, true
}

func eventControlTargetName(target string) string {
	lower := strings.ToLower(strings.TrimSpace(target))
	lower = strings.TrimPrefix(lower, "me.")
	if strings.HasPrefix(lower, "controls(\"") {
		if end := strings.Index(lower[len("controls(\""):], "\""); end >= 0 {
			return lower[len("controls(\"") : len("controls(\"")+end]
		}
	}
	if index := strings.IndexByte(lower, '.'); index >= 0 {
		return lower[:index]
	}
	return ""
}

func eventSafeProcedures(files []parsedFile, project effects.ProjectSummary) map[string]bool {
	safe := map[string]bool{}
	for _, file := range files {
		procedures := file.procedures()
		for index, proc := range procedures {
			if index >= len(file.IR.Procedures) {
				continue
			}
			id := procedureEffectIdentity(file.IR, file.IR.Procedures[index].Symbol)
			summary, ok := project.Lookup(id)
			if !ok || !summary.Has(effects.DisablesEvents) || !summary.Has(effects.RestoresEvents) {
				continue
			}
			hasTrigger, allGuarded := false, true
			for _, evidence := range summary.Direct {
				if _, _, relevant := eventEffectRisk(proc, "", evidence); relevant {
					hasTrigger = true
					if !eventGuardedAt(proc, evidence.StatementID) {
						allGuarded = false
					}
				}
			}
			for _, call := range proc.Calls {
				if call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
					continue
				}
				callee, ok := summaryForCandidate(project, call.Resolution.Candidates[0])
				if !ok {
					continue
				}
				for _, evidence := range append(append([]effects.Evidence{}, callee.Direct...), callee.Propagated...) {
					if _, _, relevant := eventEffectRisk(proc, "", evidence); relevant {
						hasTrigger = true
						if !eventGuardedAt(proc, call.StatementID) {
							allGuarded = false
						}
					}
				}
			}
			if hasTrigger && allGuarded {
				safe[id.Key()] = true
			}
		}
	}
	return safe
}

func summaryForCandidate(project effects.ProjectSummary, candidate procedureir.Candidate) (effects.ProcedureSummary, bool) {
	return project.LookupCandidate(candidate)
}

// eventGuardedAt accepts a guard only when its False assignment dominates the
// effect on normal paths and VBA203's existing all-exit analysis proves that
// assignment is restored. UserForm controls are intentionally excluded.
func eventGuardedAt(proc sourceProcedure, statementID int) bool {
	if proc.ModuleKind == "form" || proc.Graph == nil || statementID == 0 {
		return false
	}
	byID := map[int]procedureir.Statement{}
	for _, statement := range proc.Statements {
		byID[statement.ID] = statement
	}
	unsafe := applicationStateExitWitnesses(proc, "enableevents", byID)
	target, ok := proc.Graph.BlockForStatement(statementID)
	if !ok {
		return false
	}
	dominators := proc.Graph.Dominators(vbacfg.EdgeFilter{NormalOnly: true})[target.ID]
	for _, statement := range proc.Statements {
		property, value, ok := applicationPropertyAssignment(statement, byID)
		if !ok || property != "enableevents" || unsafe[statement.ID].Kind != "" || !eventDisableValue(value) {
			continue
		}
		block, ok := proc.Graph.BlockForStatement(statement.ID)
		if ok && eventContainsBlock(dominators, block.ID) {
			return true
		}
	}
	return false
}

func eventContainsBlock(blocks []vbacfg.BlockID, target vbacfg.BlockID) bool {
	for _, block := range blocks {
		if block == target {
			return true
		}
	}
	return false
}

func eventDisableValue(value string) bool {
	value = strings.ToLower(strings.TrimSpace(value))
	return value == "false" || value == "0"
}
