package analyze

import (
	"context"

	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

func (a Analyzer) implicitApproximateLookupFindingsContext(ctx context.Context, file parsedFile) ([]Finding, error) {
	if !a.Config.Analyze.DetectImplicitApproximateLookups || a.typeDB == nil {
		return nil, nil
	}
	diagnostics, err := (intel.Analyzer{RootDir: a.intelRootDir(), Config: a.Config, DB: a.typeDB}).ImplicitApproximateLookupDiagnosticsContext(ctx, file.intelDocument())
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
