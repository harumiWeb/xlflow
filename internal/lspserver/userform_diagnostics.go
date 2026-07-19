package lspserver

import (
	"strings"

	"github.com/harumiWeb/xlflow/internal/excel/forms"
	formsintel "github.com/harumiWeb/xlflow/internal/excel/forms/intel"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

// userFormYAMLDiagnostics validates the current editor buffer with the same
// source validator used by form build. The LSP owns only syntax-path-to-range
// conversion; it does not duplicate the UserForm contract.
func userFormYAMLDiagnostics(doc intel.Document) []intel.Diagnostic {
	syntax := formsintel.ParseYAML(doc.Source)
	issues, err := forms.ValidateFormSpecSource(forms.SpecInput{
		DisplayPath: doc.Path,
		Format:      "yaml",
	}, []byte(doc.Source))
	if err != nil {
		parseErr := syntax.ParseError
		if parseErr == nil {
			parseErr = err
		}
		line, character := yamlErrorPosition(doc.Source, parseErr)
		return []intel.Diagnostic{{
			Code:     "UFY001",
			Severity: "error",
			Source:   "xlflow",
			Message:  parseErr.Error(),
			Range: intel.Range{
				Start: intel.Position{Line: line, Character: character},
				End:   intel.Position{Line: line, Character: character + 1},
			},
		}}
	}

	diagnostics := make([]intel.Diagnostic, 0, len(issues))
	for _, issue := range issues {
		diagnostics = append(diagnostics, intel.Diagnostic{
			Code:     issue.Code,
			Severity: string(issue.Severity),
			Source:   "xlflow",
			Message:  issue.Message,
			Range:    userFormValidationIssueRange(syntax, issue),
		})
	}
	return diagnostics
}

func userFormValidationIssueRange(syntax *formsintel.Document, issue forms.ValidationIssue) intel.Range {
	if syntax == nil {
		return intel.Range{}
	}
	if field, ok := syntax.Field(issue.Field); ok {
		if userFormIssueUsesKeyRange(issue.Code) || isEmptyRange(field.ValueRange) {
			return userFormRange(field.KeyRange)
		}
		return userFormRange(field.ValueRange)
	}
	return userFormNearestFieldRange(syntax, issue.Field)
}

func userFormIssueUsesKeyRange(code string) bool {
	switch code {
	case "UFV001", "UFV005", "UFV013":
		return true
	default:
		return false
	}
}

func isEmptyRange(r formsintel.Range) bool {
	return r.Start == r.End
}

// userFormNearestFieldRange handles validation failures that intentionally do
// not have a concrete field, such as missing required properties and a scalar
// where a mapping item is required. It anchors at the nearest mapping sibling
// or collection key rather than underlining the document.
func userFormNearestFieldRange(syntax *formsintel.Document, path string) intel.Range {
	for current := path; ; current = userFormParentPath(current) {
		if field, ok := syntax.Field(current); ok {
			return userFormRange(field.KeyRange)
		}
		if field, ok := userFormFirstChildField(syntax, current); ok {
			return userFormRange(field.KeyRange)
		}
		if current == "" {
			break
		}
	}
	if field, ok := userFormFirstChildField(syntax, ""); ok {
		return userFormRange(field.KeyRange)
	}
	return intel.Range{}
}

func userFormRange(r formsintel.Range) intel.Range {
	return intel.Range{
		Start: intel.Position{Line: r.Start.Line, Character: r.Start.Character},
		End:   intel.Position{Line: r.End.Line, Character: r.End.Character},
	}
}

func userFormParentPath(path string) string {
	dot := strings.LastIndex(path, ".")
	bracket := strings.LastIndex(path, "[")
	if bracket > dot {
		return path[:bracket]
	}
	if dot >= 0 {
		return path[:dot]
	}
	return ""
}

func userFormFirstChildField(syntax *formsintel.Document, parent string) (formsintel.FieldNodes, bool) {
	prefix := parent
	if prefix != "" {
		prefix += "."
	}
	var (
		first formsintel.FieldNodes
		found bool
	)
	for path, field := range syntax.Fields {
		rest := strings.TrimPrefix(path, prefix)
		if path == prefix || !strings.HasPrefix(path, prefix) || strings.Contains(rest, ".") || strings.Contains(rest, "[") {
			continue
		}
		if !found || positionBefore(field.KeyRange.Start, first.KeyRange.Start) {
			first, found = field, true
		}
	}
	return first, found
}

func positionBefore(left, right formsintel.Position) bool {
	return left.Line < right.Line || left.Line == right.Line && left.Character < right.Character
}
