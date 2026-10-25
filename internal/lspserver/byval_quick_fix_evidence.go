package lspserver

import (
	"slices"

	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

const byValQuickFixDiagnosticCode = "VBA275"

type quickFixDiagnosticEvidence struct {
	document             intel.Document
	generation           uint64
	dependencyGeneration uint64
	projectAwareFull     bool
	diagnostics          []intel.Diagnostic
}

func newQuickFixDiagnosticEvidence(
	doc intel.Document,
	generation uint64,
	dependencyGeneration uint64,
	projectAwareFull bool,
	diagnostics []intel.Diagnostic,
) *quickFixDiagnosticEvidence {
	if !projectAwareFull {
		return nil
	}
	var accepted []intel.Diagnostic
	for _, diagnostic := range diagnostics {
		if diagnostic.Code != byValQuickFixDiagnosticCode {
			continue
		}
		// VBA275 has no family-specific context. Keep only its standalone
		// diagnostic value and do not retain unrelated diagnostic metadata.
		diagnostic.DefaultMember = nil
		accepted = append(accepted, diagnostic)
	}
	if len(accepted) == 0 {
		return nil
	}
	return &quickFixDiagnosticEvidence{
		document:             doc,
		generation:           generation,
		dependencyGeneration: dependencyGeneration,
		projectAwareFull:     projectAwareFull,
		diagnostics:          accepted,
	}
}

func hasByValQuickFixDiagnostic(actions []intel.CodeAction) bool {
	for _, action := range actions {
		for _, diagnostic := range action.Diagnostics {
			if diagnostic.Code == byValQuickFixDiagnosticCode {
				return true
			}
		}
	}
	return false
}

func retainCurrentByValQuickFixActions(actions []intel.CodeAction, current []intel.Diagnostic) []intel.CodeAction {
	if len(actions) == 0 {
		return actions
	}
	currentEvidence := make(map[intel.Diagnostic]struct{}, len(current))
	for _, diagnostic := range current {
		currentEvidence[diagnostic] = struct{}{}
	}
	retained := make([]intel.CodeAction, 0, len(actions))
	for _, action := range actions {
		isCurrent := true
		for _, diagnostic := range action.Diagnostics {
			if diagnostic.Code != byValQuickFixDiagnosticCode {
				continue
			}
			if _, ok := currentEvidence[diagnostic]; !ok {
				isCurrent = false
				break
			}
		}
		if isCurrent {
			retained = append(retained, action)
		}
	}
	return retained
}

// currentQuickFixDiagnostics returns only VBA275 evidence from the accepted
// Full publication for the requested current open document. It does not run
// analysis or consult project state.
func (s *Server) currentQuickFixDiagnostics(doc intel.Document) []intel.Diagnostic {
	s.diagMu.Lock()
	state := s.diagStates[doc.URI]
	if state == nil {
		s.diagMu.Unlock()
		return nil
	}
	state.mu.Lock()
	s.diagMu.Unlock()
	defer state.mu.Unlock()

	evidence := state.quickFixEvidence
	if !state.open || state.dependentPending || !state.hasPublished ||
		state.publishedMode != intel.DiagnosticModeFull || evidence == nil ||
		!evidence.projectAwareFull ||
		evidence.generation != state.generation ||
		evidence.dependencyGeneration != state.dependencyGeneration ||
		!sameScheduledDocument(state.latest, doc) ||
		!sameScheduledDocument(evidence.document, doc) {
		return nil
	}
	return slices.Clone(evidence.diagnostics)
}
