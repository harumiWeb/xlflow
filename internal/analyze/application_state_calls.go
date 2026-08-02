package analyze

import (
	"strings"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/effects"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// applicationStateLeakOrigin is a direct Application-state assignment that
// VBA203 has proved can leave its procedure without restoration. It preserves
// enough provenance for callers to explain the same root cause without
// re-running the CFG analysis at every call site.
type applicationStateLeakOrigin struct {
	Property    string
	Identity    effects.ProcedureIdentity
	StatementID int
	Line        int
	Witness     applicationStateExitWitness
}

type applicationStateLeakIndex map[string]applicationStateLeakOrigin

func buildApplicationStateLeakIndex(files []parsedFile, project effects.ProjectSummary) applicationStateLeakIndex {
	index := applicationStateLeakIndex{}
	for _, file := range files {
		for _, proc := range sourceProceduresWithEffects(file, project) {
			for _, origin := range applicationStateLeakOrigins(proc, project) {
				index[applicationStateLeakKey(origin.Identity, origin.StatementID, origin.Property)] = origin
			}
		}
	}
	return index
}

func applicationStateLeakOrigins(proc sourceProcedure, project effects.ProjectSummary) []applicationStateLeakOrigin {
	if proc.Graph == nil || proc.Effects == nil {
		return nil
	}
	byID := make(map[int]procedureir.Statement, len(proc.Statements))
	for _, statement := range proc.Statements {
		byID[statement.ID] = statement
	}
	var out []applicationStateLeakOrigin
	for _, property := range applicationStateProperties() {
		unsafe := applicationStateExitWitnesses(proc, property.Key, byID)
		if len(unsafe) == 0 || hasPairedApplicationRestoreProcedure(proc, property.Key, project) {
			continue
		}
		for _, statement := range proc.Statements {
			witness, found := unsafe[statement.ID]
			if !found {
				continue
			}
			assigned, _, ok := applicationPropertyAssignment(statement, byID)
			if !ok || assigned != property.Key {
				continue
			}
			out = append(out, applicationStateLeakOrigin{
				Property: property.Key, Identity: proc.Effects.Identity,
				StatementID: statement.ID, Line: statement.Range.StartLine, Witness: witness,
			})
		}
	}
	return out
}

func applicationStateLeakKey(identity effects.ProcedureIdentity, statementID int, property string) string {
	return strings.Join([]string{identity.Key(), strconvItoa(statementID), strings.ToLower(property)}, "\x00")
}

func (index applicationStateLeakIndex) lookup(evidence effects.Evidence) (applicationStateLeakOrigin, bool) {
	if evidence.Effect != effects.ChangesApplicationState {
		return applicationStateLeakOrigin{}, false
	}
	property, ok := applicationStatePropertyKey(evidence.Target)
	if !ok {
		return applicationStateLeakOrigin{}, false
	}
	origin, ok := index[applicationStateLeakKey(evidence.Origin, evidence.StatementID, property)]
	return origin, ok
}

func applicationStatePropertyKey(target string) (string, bool) {
	target = strings.ToLower(strings.TrimSpace(target))
	for _, property := range applicationStateProperties() {
		if target == "application."+property.Key {
			return property.Key, true
		}
	}
	return "", false
}

func applicationStatePropertyName(key string) string {
	for _, property := range applicationStateProperties() {
		if property.Key == key {
			return property.Name
		}
	}
	return key
}

// applicationStateCallEffectFindings reports only the immediate call that
// turns a direct root leak into caller context. A transitive callee summary is
// deliberately not enough: reporting it again at every ancestor would make a
// single leak noisy without adding actionable context.
func (a Analyzer) applicationStateCallEffectFindings(file parsedFile, proc sourceProcedure, project effects.ProjectSummary) []Finding {
	if proc.Effects == nil || len(a.applicationStateLeaks) == 0 {
		return nil
	}
	var out []Finding
	reported := map[string]bool{}
	for _, call := range proc.Calls {
		if !applicationStateCallReachable(proc, call) || call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
			continue
		}
		callee, ok := summaryForCandidate(project, call.Resolution.Candidates[0])
		if !ok {
			continue
		}
		for _, evidence := range callee.Direct {
			origin, ok := a.applicationStateLeaks.lookup(evidence)
			if !ok {
				continue
			}
			key := strings.Join([]string{strconvItoa(call.ID), origin.Property}, "\x00")
			if reported[key] {
				continue
			}
			reported[key] = true
			property := "Application." + applicationStatePropertyName(origin.Property)
			originName := origin.Identity.QualifiedName
			if originName == "" {
				originName = origin.Identity.Name
			}
			reason := "This call propagates the un-restored " + property + " change introduced by " + originName + " at line " + strconvItoa(origin.Line) + "."
			if uncertainty := applicationStateCallUncertainty(callee); uncertainty != "" {
				reason += " The callee's final state is also uncertain because it reaches " + uncertainty + " call dispatch."
			}
			out = append(out, a.simpleFinding(
				file, proc, call.Range.StartLine, "VBA221", "warning",
				"Call to "+call.Callee.Text+" can leave "+property+" changed.",
				reason,
				"Restore the previous "+property+" value in the owning cleanup path, or make the helper's state ownership explicit.",
			))
		}
	}
	return out
}

func applicationStateCallReachable(proc sourceProcedure, call procedureir.CallSite) bool {
	for _, statement := range proc.Statements {
		if statement.ID == call.StatementID && statement.Recovered {
			return false
		}
	}
	if proc.Graph == nil {
		return true
	}
	for _, blockID := range proc.Graph.Reachable(vbacfg.EdgeFilter{}) {
		for _, block := range proc.Graph.Blocks {
			if block.ID == blockID && block.Kind == vbacfg.BlockStatement && block.StatementID == call.StatementID {
				return true
			}
		}
	}
	return false
}

func applicationStateCallUncertainty(summary effects.ProcedureSummary) string {
	seen := map[string]bool{}
	var kinds []string
	for _, uncertainty := range append(append([]effects.CallUncertainty{}, summary.DirectUncertainty...), summary.PropagatedUncertainty...) {
		kind := string(uncertainty.Kind)
		if !seen[kind] {
			seen[kind] = true
			kinds = append(kinds, kind)
		}
	}
	return strings.Join(kinds, ", ")
}
