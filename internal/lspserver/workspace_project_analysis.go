package lspserver

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"path/filepath"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type projectProcedureState struct {
	file        string
	fingerprint string
}

type projectDependencyView struct {
	procedures map[string]projectProcedureState
	reverse    map[string][]string
}

// projectImpactPaths returns the transitive caller files affected by changes
// between two coherent snapshots. Both old and new reverse edges participate,
// so removing or redirecting a call still refreshes the former callers.
func projectImpactPaths(before, after intel.ProjectAnalysisSnapshot) []string {
	return projectImpactPathsWithPerformance(before, after, nil)
}

func projectImpactPathsWithPerformance(before, after intel.ProjectAnalysisSnapshot, performance *performanceRecorder) []string {
	return projectImpactPathsWithPerformanceClass(before, after, performance, "background")
}

func projectImpactPathsWithPerformanceClass(before, after intel.ProjectAnalysisSnapshot, performance *performanceRecorder, class string) []string {
	if !before.Complete || !after.Complete {
		return nil
	}
	oldView := buildProjectDependencyViewWithPerformanceClass(before, performance, class)
	newView := buildProjectDependencyViewWithPerformanceClass(after, performance, class)
	changed := make(map[string]bool)
	for key, old := range oldView.procedures {
		if current, ok := newView.procedures[key]; !ok || current.fingerprint != old.fingerprint {
			changed[key] = true
		}
	}
	for key, current := range newView.procedures {
		if old, ok := oldView.procedures[key]; !ok || old.fingerprint != current.fingerprint {
			changed[key] = true
		}
	}
	queue := make([]string, 0, len(changed))
	for key := range changed {
		queue = append(queue, key)
	}
	seen := make(map[string]bool, len(changed))
	files := make(map[string]string)
	for key := range changed {
		if state, ok := newView.procedures[key]; ok {
			files[symbolFileKey(state.file)] = state.file
		} else if state, ok := oldView.procedures[key]; ok {
			files[symbolFileKey(state.file)] = state.file
		}
	}
	for len(queue) > 0 {
		callee := queue[0]
		queue = queue[1:]
		if seen[callee] {
			continue
		}
		seen[callee] = true
		callers := append(append([]string(nil), oldView.reverse[callee]...), newView.reverse[callee]...)
		for _, caller := range callers {
			if state, ok := newView.procedures[caller]; ok {
				files[symbolFileKey(state.file)] = state.file
			} else if state, ok := oldView.procedures[caller]; ok {
				files[symbolFileKey(state.file)] = state.file
			}
			if !seen[caller] {
				queue = append(queue, caller)
			}
		}
	}
	out := make([]string, 0, len(files))
	for _, file := range files {
		out = append(out, file)
	}
	sort.Slice(out, func(i, j int) bool { return symbolFileKey(out[i]) < symbolFileKey(out[j]) })
	return out
}

func buildProjectDependencyViewWithPerformanceClass(snapshot intel.ProjectAnalysisSnapshot, performance *performanceRecorder, class string) projectDependencyView {
	view := projectDependencyView{procedures: make(map[string]projectProcedureState), reverse: make(map[string][]string)}
	for _, document := range snapshot.Documents {
		performance.addCounter(performanceCounterProcedureFingerprintBuilds, uint64(len(document.IR.Procedures)), "workspace/project", performanceStageDependencyUpdate, class, document.IR.Path)
		for procedureIndex, sourceProcedure := range document.IR.Procedures {
			procedure := sourceProcedure
			if resolved, ok := document.Resolution.ResolvedProcedure(procedureIndex); ok {
				procedure = resolved
			}
			key := projectProcedureKey(document.IR.Path, procedure.Symbol.QualifiedName, string(procedure.Symbol.Kind), procedure.Symbol.DeclarationRange.StartLine)
			encoded, _ := json.Marshal(procedure)
			sum := sha256.Sum256(encoded)
			view.procedures[key] = projectProcedureState{file: document.IR.Path, fingerprint: hex.EncodeToString(sum[:])}
			for _, call := range procedure.Calls {
				if call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
					continue
				}
				candidate := call.Resolution.Candidates[0]
				target := projectProcedureKey(candidate.File, candidate.QualifiedName, candidate.Kind, candidate.Line)
				view.reverse[target] = appendUniqueString(view.reverse[target], key)
			}
		}
	}
	return view
}

func projectProcedureKey(file, qualifiedName, kind string, line int) string {
	return strings.Join([]string{
		strings.ToLower(filepath.ToSlash(filepath.Clean(file))),
		strings.ToLower(strings.TrimSpace(qualifiedName)), strings.ToLower(strings.TrimSpace(kind)),
		decimalString(line),
	}, "\x00")
}

func appendUniqueString(values []string, value string) []string {
	for _, current := range values {
		if current == value {
			return values
		}
	}
	return append(values, value)
}

func decimalString(value int) string {
	if value == 0 {
		return "0"
	}
	var digits [24]byte
	i := len(digits)
	for value > 0 {
		i--
		digits[i] = byte('0' + value%10)
		value /= 10
	}
	return string(digits[i:])
}
