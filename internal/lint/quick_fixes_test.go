package lint

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
)

func TestQuickFixesUseDiagnosticPredicate(t *testing.T) {
	source := []byte("Option Explicit\nSub Work()\nDim i As Long\nFor i = 1 To 3 Step 1\nDebug.Print i\nNext i\nEnd Sub\n")
	doc, err := vbaast.ParseDocument("Main.bas", source)
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	cfg := config.Default()
	cfg.Lint.DetectRedundantStepOne = true
	linter := Linter{RootDir: ".", Config: cfg}
	fixes, err := linter.QuickFixesParsedContext(t.Context(), doc)
	if err != nil {
		t.Fatal(err)
	}
	if len(fixes) != 1 {
		t.Fatalf("fixes = %+v", fixes)
	}
	issues, err := linter.LintParsed(doc)
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range issues {
		if issue.Code == "VB084" {
			if issue != fixes[0].Diagnostic {
				t.Fatalf("fix diagnostic = %+v, lint = %+v", fixes[0].Diagnostic, issue)
			}
			return
		}
	}
	t.Fatal("VB084 diagnostic missing")
}
