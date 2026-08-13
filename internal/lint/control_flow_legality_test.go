package lint

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
)

func TestLinterProjectsCFGControlFlowLegality(t *testing.T) {
	doc, err := vbaast.ParseDocument("Main.bas", []byte(`Sub Main()
    GoTo Missing
    Label1:
    label1:
    For i = 1 To 2
    Next wrong
    Exit Function
End Sub`))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	issues, err := (Linter{RootDir: ".", Config: config.Default()}).LintParsed(doc)
	if err != nil {
		t.Fatal(err)
	}
	seen := map[string]int{}
	for _, issue := range issues {
		if issue.Code >= "VB055" && issue.Code <= "VB058" {
			seen[issue.Code]++
		}
	}
	for _, code := range []string{"VB055", "VB056", "VB057", "VB058"} {
		if seen[code] != 1 {
			t.Fatalf("%s count = %d, issues = %+v", code, seen[code], issues)
		}
	}
}

func TestLinterDoesNotInlineSuppressCFGControlFlowLegality(t *testing.T) {
	doc, err := vbaast.ParseDocument("Main.bas", []byte(`Sub Main()
    ' xlflow:disable-next-line VB056
    GoTo Missing
End Sub`))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	result, err := (Linter{RootDir: ".", Config: config.Default()}).LintParsedContext(t.Context(), doc)
	if err != nil {
		t.Fatal(err)
	}
	found := false
	for _, issue := range result {
		if issue.Code == "VB056" {
			found = true
		}
	}
	if !found {
		t.Fatalf("VB056 was suppressed unexpectedly: %+v", result)
	}
}
