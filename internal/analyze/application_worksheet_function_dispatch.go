package analyze

import (
	"context"
	"sync"

	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

func (a Analyzer) applicationWorksheetFunctionDispatchFindingsContext(ctx context.Context, file parsedFile) ([]Finding, error) {
	if !a.Config.Analyze.DetectApplicationWorksheetFunction || a.typeDB == nil || a.typeDBResolutionIncomplete {
		return nil, nil
	}
	intelAnalyzer := intel.Analyzer{
		RootDir:                    a.intelRootDir(),
		Config:                     a.Config,
		DB:                         a.typeDB,
		SourceOnly:                 file.sourceProject,
		TypeDBResolutionIncomplete: a.typeDBResolutionIncomplete,
	}
	openDocuments := []intel.Document{file.intelDocument()}
	if len(a.workspaceDocuments) > 0 && a.workspaceSymbolsSnapshot != nil {
		openDocuments = a.workspaceDocuments
	}
	if a.workspaceSymbolsSnapshot != nil {
		type workspaceResolutionSnapshot struct {
			view *intel.WorkspaceResolutionView
			err  error
		}
		loadWorkspace := sync.OnceValue(func() workspaceResolutionSnapshot {
			symbols, err := a.workspaceSymbolsSnapshot(openDocuments)
			if err != nil {
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
	diagnostics, err := intelAnalyzer.ApplicationWorksheetFunctionDispatchDiagnosticsContext(ctx, file.intelDocument())
	if err != nil {
		return nil, err
	}
	procedures := file.procedureView()
	out := make([]Finding, 0, len(diagnostics))
	for _, diagnostic := range diagnostics {
		line := diagnostic.Range.Start.Line + 1
		proc := procedureForLineView(procedures, line, len(file.Lines))
		finding := a.simpleFinding(file, proc, line, diagnostic.Code, diagnostic.Severity, diagnostic.Message,
			"Application worksheet-function dispatch and the explicit WorksheetFunction object can expose different error and return-value contracts.",
			"Use Application.WorksheetFunction explicitly when its generated TypeLib contract is the intended behavior.")
		finding.Column = diagnostic.Range.Start.Character + 1
		finding.EndLine = diagnostic.Range.End.Line + 1
		finding.EndColumn = diagnostic.Range.End.Character + 1
		out = append(out, finding)
	}
	return out, ctx.Err()
}
