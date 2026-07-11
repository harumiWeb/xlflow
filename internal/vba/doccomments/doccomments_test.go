package doccomments

import (
	"strings"
	"testing"
)

func TestParseDocLinesStructuredSections(t *testing.T) {
	doc := ParseDocLines([]string{
		"''' Calculates total sales for the requested sheet.",
		"'''",
		"''' Longer body text.",
		"'''",
		"''' Args:",
		"'''     ws: Worksheet to aggregate.",
		"'''     includeTax: True to include tax.",
		"'''",
		"''' Returns:",
		"'''     Aggregated sales amount.",
		"'''",
		"''' Errors:",
		"'''     Error 5: ws is Nothing.",
		"'''",
		"''' Custom:",
		"'''     preserved",
	})

	if doc.Summary != "Calculates total sales for the requested sheet." {
		t.Fatalf("summary = %q", doc.Summary)
	}
	if doc.Body != "Longer body text." {
		t.Fatalf("body = %q", doc.Body)
	}
	if doc.Parameters["ws"] != "Worksheet to aggregate." || doc.Parameters["includeTax"] == "" {
		t.Fatalf("parameters = %+v", doc.Parameters)
	}
	if doc.Returns != "Aggregated sales amount." || doc.Errors != "Error 5: ws is Nothing." {
		t.Fatalf("returns/errors = %q/%q", doc.Returns, doc.Errors)
	}
	if doc.UnknownSections["custom"] != "preserved" {
		t.Fatalf("unknown sections = %+v", doc.UnknownSections)
	}
}

func TestDocumentationForTargetMergesRubberduckSummary(t *testing.T) {
	source := `''' Args:
'''     workbook: Target workbook.
'@Description("Processes the workbook.")
Public Sub Process(ByVal workbook As Workbook)
End Sub
`
	doc, start, ok := DocumentationForTarget(source, 4, "symbol")
	if !ok {
		t.Fatal("documentation not found")
	}
	if start != 1 {
		t.Fatalf("start = %d, want 1", start)
	}
	if doc.Summary != "Processes the workbook." {
		t.Fatalf("summary = %q", doc.Summary)
	}
	if doc.Parameters["workbook"] == "" {
		t.Fatalf("parameters = %+v", doc.Parameters)
	}
}

func TestValidateFindsArgumentProblems(t *testing.T) {
	doc := ParseDocLines([]string{
		"''' Args:",
		"'''     wss: typo.",
		"'''     ws: correct name.",
		"'''     WS: case duplicate.",
		"'''",
		"''' Returns:",
		"'''     invalid",
	})
	diagnostics := Validate(Procedure{
		Name:       "Process",
		Kind:       "sub",
		Parameters: []Parameter{{Name: "ws"}},
	}, doc, 1)
	if !hasDiagnostic(diagnostics, "VB040") || !hasDiagnostic(diagnostics, "VB041") || !hasDiagnostic(diagnostics, "VB042") {
		t.Fatalf("diagnostics = %+v", diagnostics)
	}
	if !strings.Contains(diagnostics[0].Message, `Did you mean "ws"`) {
		t.Fatalf("missing suggestion: %+v", diagnostics)
	}
}

func TestGenerateSnippetByProcedureKind(t *testing.T) {
	functionSnippet := GenerateSnippet(Procedure{
		Name:       "FindCustomer",
		Kind:       "function",
		Parameters: []Parameter{{Name: "customerCode"}},
		ReturnType: "Customer",
	})
	if !strings.Contains(functionSnippet.Text, "Args:") || !strings.Contains(functionSnippet.Text, "Returns:") {
		t.Fatalf("function snippet = %q", functionSnippet.Text)
	}

	subSnippet := GenerateSnippet(Procedure{Name: "Initialize", Kind: "sub"})
	if strings.Contains(subSnippet.Text, "Args:") || strings.Contains(subSnippet.Text, "Returns:") {
		t.Fatalf("sub snippet = %q", subSnippet.Text)
	}

	propertySnippet := GenerateSnippet(Procedure{Name: "CurrentCustomer", Kind: "property_get", ReturnType: "Customer"})
	if strings.Contains(propertySnippet.Text, "Args:") || !strings.Contains(propertySnippet.Text, "Returns:") {
		t.Fatalf("property get snippet = %q", propertySnippet.Text)
	}

	indexedPropertySnippet := GenerateSnippet(Procedure{
		Name:       "Item",
		Kind:       "property_get",
		Parameters: []Parameter{{Name: "index"}},
		ReturnType: "Variant",
	})
	if !strings.Contains(indexedPropertySnippet.Text, "Args:") || !strings.Contains(indexedPropertySnippet.Text, "index: ${2:Parameter description.}") || !strings.Contains(indexedPropertySnippet.Text, "Returns:") {
		t.Fatalf("indexed property get snippet = %q", indexedPropertySnippet.Text)
	}
}

func TestGenerateCommentOmitsSnippetPlaceholders(t *testing.T) {
	comment := GenerateComment(Procedure{
		Name:       "FindCustomer",
		Kind:       "function",
		Parameters: []Parameter{{Name: "customerCode"}},
		ReturnType: "Customer",
	})
	if strings.Contains(comment.Text, "${") {
		t.Fatalf("comment should not contain snippet placeholders: %q", comment.Text)
	}
	if !strings.Contains(comment.Text, "customerCode: Parameter description.") || !strings.Contains(comment.Text, "Return value description.") {
		t.Fatalf("comment text = %q", comment.Text)
	}
}

func TestMarkdownRendersUnknownSectionsDeterministically(t *testing.T) {
	doc := SymbolDocumentation{
		Summary: "Summary.",
		UnknownSections: map[string]string{
			"zeta":  "last",
			"alpha": "first",
		},
	}
	markdown := Markdown(doc, "")
	alpha := strings.Index(markdown, "**alpha**")
	zeta := strings.Index(markdown, "**zeta**")
	if alpha < 0 || zeta < 0 || alpha > zeta {
		t.Fatalf("unknown sections not sorted: %q", markdown)
	}
}

func TestParseDocLinesRemovesCommonIndentation(t *testing.T) {
	doc := ParseDocLines([]string{
		"''' Summary.",
		"'''",
		"''' Remarks:",
		"'''     First line.",
		"'''         Nested line.",
	})
	if doc.Remarks != "First line.\n    Nested line." {
		t.Fatalf("remarks = %q", doc.Remarks)
	}
}

func hasDiagnostic(diagnostics []Diagnostic, code string) bool {
	for _, diagnostic := range diagnostics {
		if diagnostic.Code == code {
			return true
		}
	}
	return false
}
