package analyze

import (
	"context"
	"sync"

	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

func (a Analyzer) implicitApproximateLookupFindingsContext(ctx context.Context, file parsedFile) ([]Finding, error) {
	if !a.Config.Analyze.DetectImplicitApproximateLookups || a.typeDB == nil {
		return nil, nil
	}
	intelAnalyzer := intel.Analyzer{
		RootDir:    a.intelRootDir(),
		Config:     a.Config,
		DB:         a.typeDB,
		SourceOnly: file.sourceProject,
	}
	openDocuments := []intel.Document{file.intelDocument()}
	if len(a.workspaceDocuments) > 0 && a.workspaceSymbolsSnapshot != nil {
		openDocuments = a.workspaceDocuments
	}
	// The project-aware LSP path supplies a coherent open-document set. Keep
	// the callback and that set together so a same-name Public procedure in an
	// unsaved sibling module can shadow the Excel fallback conservatively.
	if a.workspaceSymbolsSnapshot != nil {
		type workspaceResolutionSnapshot struct {
			view *intel.WorkspaceResolutionView
			err  error
		}
		loadWorkspace := sync.OnceValue(func() workspaceResolutionSnapshot {
			if err := ctx.Err(); err != nil {
				return workspaceResolutionSnapshot{err: err}
			}
			symbols, err := a.workspaceSymbolsSnapshot(openDocuments)
			if err != nil {
				return workspaceResolutionSnapshot{err: err}
			}
			if err := ctx.Err(); err != nil {
				return workspaceResolutionSnapshot{err: err}
			}
			return workspaceResolutionSnapshot{view: intel.NewWorkspaceResolutionView(symbols)}
		})
		intelAnalyzer.WorkspaceSymbolQueryFunc = func(_ []intel.Document, query intel.WorkspaceSymbolQuery) ([]intel.Symbol, error) {
			workspace := loadWorkspace()
			if workspace.err != nil {
				return nil, workspace.err
			}
			return workspace.view.Query(query), nil
		}
	}
	diagnostics, err := intelAnalyzer.ImplicitApproximateLookupDiagnosticsContext(ctx, file.intelDocument())
	if err != nil {
		return nil, err
	}
	if len(diagnostics) == 0 {
		return nil, nil
	}
	procedures := file.procedureView()
	out := make([]Finding, 0, len(diagnostics))
	for i, diagnostic := range diagnostics {
		if i&0x3f == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		line := diagnostic.Range.Start.Line + 1
		proc := procedureForLineView(procedures, line, len(file.Lines))
		finding := a.simpleFinding(
			file,
			proc,
			line,
			diagnostic.Code,
			diagnostic.Severity,
			diagnostic.Message,
			"Excel uses approximate matching when this lookup's match-mode argument is omitted, so an unsorted or unexpectedly ordered range can produce a plausible but incorrect result.",
			"Pass the match-mode argument explicitly: use 0 or False for exact matching, or 1, -1, or True when approximate matching is intentional and its ordering contract is verified.",
		)
		finding.Column = diagnostic.Range.Start.Character + 1
		finding.EndLine = diagnostic.Range.End.Line + 1
		finding.EndColumn = diagnostic.Range.End.Character + 1
		out = append(out, finding)
	}
	return out, ctx.Err()
}
