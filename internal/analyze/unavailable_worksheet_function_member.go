package analyze

import (
	"context"
	"sync"

	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

// unavailableWorksheetFunctionMemberFindingsContext adapts the typed
// WorksheetFunction-member diagnostics to the analyzer Finding envelope. The
// intel checker owns receiver/member resolution; this layer adds procedure
// context and preserves the source range.
func (a Analyzer) unavailableWorksheetFunctionMemberFindingsContext(ctx context.Context, file parsedFile) ([]Finding, error) {
	if !a.Config.Analyze.DetectUnavailableWorksheetFunctionMembers || a.typeDB == nil || a.typeDBResolutionIncomplete {
		return nil, nil
	}

	intelAnalyzer := intel.Analyzer{
		RootDir:                    a.intelRootDir(),
		Config:                     a.Config,
		DB:                         a.typeDB,
		TypeDBResolutionIncomplete: a.typeDBResolutionIncomplete,
	}
	openDocuments := []intel.Document{file.intelDocument()}
	if len(a.workspaceDocuments) > 0 && a.workspaceSymbolsSnapshot != nil {
		openDocuments = a.workspaceDocuments
	}
	// A project-aware realtime request already owns one immutable workspace
	// symbol snapshot. Reuse it so a local procedure with the same name cannot
	// be mistaken for an Excel WorksheetFunction call.
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

	diagnostics, err := intelAnalyzer.UnavailableWorksheetFunctionMemberDiagnosticsContext(ctx, file.intelDocument())
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
			"The complete generated Excel TypeLib does not expose this member on Excel.WorksheetFunction.",
			"Use an available WorksheetFunction member, call the correct Excel API, or refresh the generated Excel TypeLib database.",
		)
		finding.Column = diagnostic.Range.Start.Character + 1
		finding.EndLine = diagnostic.Range.End.Line + 1
		finding.EndColumn = diagnostic.Range.End.Character + 1
		out = append(out, finding)
	}
	return out, ctx.Err()
}
