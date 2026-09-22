package analyze

import (
	"context"
	"sync"

	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

// DefaultMemberContext is the stable machine-readable classification attached
// to VBA253-VBA255 and default-member-owned VBA249 findings.
type DefaultMemberContext struct {
	Kind            string `json:"kind"`
	Binding         string `json:"binding"`
	ExpectedContext string `json:"expected_context"`
	Member          string `json:"member,omitempty"`
	Depth           int    `json:"depth"`
}

func (a Analyzer) defaultMemberFindingsContext(ctx context.Context, file parsedFile) ([]Finding, error) {
	if a.typeDB == nil {
		return nil, nil
	}
	intelAnalyzer := intel.Analyzer{
		RootDir:                    a.intelRootDir(),
		Config:                     a.Config,
		DB:                         a.typeDB,
		TypeDBResolutionIncomplete: a.typeDBResolutionIncomplete,
		ProjectDefaultMembers:      a.projectDefaultMembers,
		ProjectDefaultTypes:        a.projectDefaultTypes,
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
			if err := ctx.Err(); err != nil {
				return workspaceResolutionSnapshot{err: err}
			}
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

	diagnostics, err := intelAnalyzer.DefaultMemberDiagnosticsContext(ctx, file.intelDocument())
	if err != nil || len(diagnostics) == 0 {
		return nil, err
	}
	procedures := file.procedureView()
	out := make([]Finding, 0, len(diagnostics))
	for index, diagnostic := range diagnostics {
		if index&0x3f == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		line := diagnostic.Range.Start.Line + 1
		proc := procedureForLineView(procedures, line, len(file.Lines))
		reason, suggestion := defaultMemberGuidance(diagnostic.Code)
		finding := a.simpleFinding(file, proc, line, diagnostic.Code, diagnostic.Severity, diagnostic.Message, reason, suggestion)
		finding.Column = diagnostic.Range.Start.Character + 1
		finding.EndLine = diagnostic.Range.End.Line + 1
		finding.EndColumn = diagnostic.Range.End.Character + 1
		if diagnostic.DefaultMember != nil {
			finding.DefaultMember = &DefaultMemberContext{
				Kind: diagnostic.DefaultMember.Kind, Binding: diagnostic.DefaultMember.Binding,
				ExpectedContext: diagnostic.DefaultMember.ExpectedContext,
				Member:          diagnostic.DefaultMember.Member, Depth: diagnostic.DefaultMember.Depth,
			}
		}
		if diagnostic.Code == "VBA249" {
			runtimeKind := "default_member_required"
			if diagnostic.DefaultMember != nil && diagnostic.DefaultMember.Kind == "recursive" {
				runtimeKind = "default_member_cycle"
			}
			finding.RuntimeError = &RuntimeErrorContext{Kind: runtimeKind}
		}
		out = append(out, finding)
	}
	return out, ctx.Err()
}

func defaultMemberGuidance(code string) (reason, suggestion string) {
	switch code {
	case "VBA253":
		return "VBA is relying on a type's default member instead of an explicitly named value-producing member.", "Name the intended member explicitly, such as .Item or .Value, so the binding remains visible and stable."
	case "VBA254":
		return "The receiver is Object, Variant, late-bound, or otherwise incomplete, so the invoked default member cannot be verified statically.", "Use an early-bound type and name the intended member explicitly."
	case "VBA255":
		return "Bang notation converts the member name into string-based default-member access.", "Use an explicit typed member or an explicit Item-style access when dynamic lookup is intentional."
	default:
		return "Complete type metadata proves that the required value-producing default member is unavailable.", "Use an explicit supported member or correct the receiver type before evaluating the expression."
	}
}
