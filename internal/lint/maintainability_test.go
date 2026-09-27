package lint

import (
	"context"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
)

func TestMaintainabilityOptInSyntax(t *testing.T) {
	cases := []struct {
		name, source, code string
		enable             func(*config.LintConfig)
	}{
		{"empty if", "Sub Work()\nIf True Then\nEnd If\nEnd Sub\n", "VB067", func(c *config.LintConfig) { c.DetectEmptyIf = true }},
		{"empty else", "Sub Work()\nIf True Then\nDebug.Print 1\nElse\nEnd If\nEnd Sub\n", "VB068", func(c *config.LintConfig) { c.DetectEmptyElse = true }},
		{"empty case", "Sub Work()\nSelect Case 1\nCase 1\nEnd Select\nEnd Sub\n", "VB069", func(c *config.LintConfig) { c.DetectEmptyCase = true }},
		{"empty for", "Sub Work()\nDim i As Long\nFor i = 1 To 3\nNext i\nEnd Sub\n", "VB070", func(c *config.LintConfig) { c.DetectEmptyFor = true }},
		{"empty for each", "Sub Work(items As Collection)\nDim item As Variant\nFor Each item In items\nNext item\nEnd Sub\n", "VB071", func(c *config.LintConfig) { c.DetectEmptyForEach = true }},
		{"empty do", "Sub Work()\nDo\nLoop\nEnd Sub\n", "VB072", func(c *config.LintConfig) { c.DetectEmptyDo = true }},
		{"empty while", "Sub Work()\nWhile True\nWend\nEnd Sub\n", "VB073", func(c *config.LintConfig) { c.DetectEmptyWhile = true }},
		{"empty procedure", "Sub Work()\n' only comment\nEnd Sub\n", "VB074", func(c *config.LintConfig) { c.DetectEmptyProcedure = true }},
		{"empty module", "Option Explicit\nPrivate value As Long\n", "VB075", func(c *config.LintConfig) { c.DetectEmptyModule = true }},
		{"call", "Sub Work()\nCall Other()\nEnd Sub\nSub Other()\nEnd Sub\n", "VB076", func(c *config.LintConfig) { c.DetectLegacyCall = true }},
		{"rem", "Rem note\nSub Work()\nEnd Sub\n", "VB077", func(c *config.LintConfig) { c.DetectRemComment = true }},
		{"error statement", "Sub Work()\nError 11\nEnd Sub\n", "VB078", func(c *config.LintConfig) { c.DetectErrorStatement = true }},
		{"global declaration", "Global amount As Long\n", "VB079", func(c *config.LintConfig) { c.DetectGlobalDeclaration = true }},
		{"let statement", "Sub Work()\nDim amount As Long\nLet amount = 1\nEnd Sub\n", "VB080", func(c *config.LintConfig) { c.DetectLetAssignment = true }},
		{"on local error", "Sub Work()\nOn Local Error Resume Next\nEnd Sub\n", "VB092", func(c *config.LintConfig) { c.DetectOnLocalError = true }},
		{"suffix", "Sub Work()\nDim amount$\nEnd Sub\n", "VB081", func(c *config.LintConfig) { c.DetectIdentifierTypeSuffix = true }},
		{"while wend", "Sub Work()\nWhile True\nWend\nEnd Sub\n", "VB082", func(c *config.LintConfig) { c.DetectWhileWend = true }},
		{"def type", "DefInt A-Z\nSub Work()\nEnd Sub\n", "VB083", func(c *config.LintConfig) { c.DetectDefType = true }},
		{"step one", "Sub Work()\nDim i As Long\nFor i = 1 To 3 Step 1\nDebug.Print i\nNext i\nEnd Sub\n", "VB084", func(c *config.LintConfig) { c.DetectRedundantStepOne = true }},
		{"step omitted", "Sub Work()\nDim i As Long\nFor i = 1 To 3\nDebug.Print i\nNext i\nEnd Sub\n", "VB085", func(c *config.LintConfig) { c.RequireExplicitStep = true }},
		{"option base", "Option Base 0\nSub Work()\nEnd Sub\n", "VB086", func(c *config.LintConfig) { c.DetectRedundantOptionBaseZero = true }},
		{"module dim", "Dim value As Long\nSub Work()\nEnd Sub\n", "VB087", func(c *config.LintConfig) { c.DetectModuleDim = true }},
		{"implicit public", "Sub Work()\nEnd Sub\n", "VB088", func(c *config.LintConfig) { c.DetectImplicitPublic = true }},
		{"multiple", "Sub Work()\nDim a As Long, b As Long\nEnd Sub\n", "VB089", func(c *config.LintConfig) { c.DetectMultipleDeclarations = true }},
		{"unused label", "Sub Work()\nunused:\nDebug.Print 1\nEnd Sub\n", "VB090", func(c *config.LintConfig) { c.DetectUnusedLabels = true }},
		{"stop", "Sub Work()\nStop\nEnd Sub\n", "VB091", func(c *config.LintConfig) { c.DetectStopStatement = true }},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			cfg := config.Default()
			tc.enable(&cfg.Lint)
			doc, err := vbaast.ParseDocument("Main.bas", []byte(tc.source))
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()
			var found []Issue
			if err := doc.Read(func(view vbaast.ParsedView) error {
				found = (Linter{RootDir: ".", ModuleKind: "standard", Config: cfg}).maintainabilityIssues("Main.bas", view.Source, view.Root)
				return nil
			}); err != nil {
				t.Fatal(err)
			}
			for _, issue := range found {
				if issue.Code == tc.code {
					return
				}
			}
			t.Fatalf("missing %s: %+v", tc.code, found)
		})
	}
}

func TestMaintainabilityDefaultDisabled(t *testing.T) {
	doc, err := vbaast.ParseDocument("Main.bas", []byte("Sub Work()\nStop\nEnd Sub\n"))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	got, err := (Linter{RootDir: ".", Config: config.Default()}).LintParsedContext(context.Background(), doc)
	if err != nil {
		t.Fatal(err)
	}
	for _, issue := range got {
		if issue.Code == "VB091" {
			t.Fatalf("Stop rule enabled by default: %+v", got)
		}
	}
}

func TestMaintainabilityUnusedLabelsRespectsAllBranchTargets(t *testing.T) {
	source := "Sub Work(ByVal choice As Long)\nOn choice GoTo first, second\nfirst:\nGoTo done\nsecond:\nOn Error GoTo handler\ndone:\nExit Sub\nhandler:\nResume done\nEnd Sub\n"
	cfg := config.Default()
	cfg.Lint.DetectUnusedLabels = true
	doc, err := vbaast.ParseDocument("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	if err := doc.Read(func(view vbaast.ParsedView) error {
		if got := (Linter{RootDir: ".", Config: cfg}).maintainabilityIssues("Main.bas", view.Source, view.Root); len(got) != 0 {
			t.Errorf("referenced labels reported unused: %+v", got)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

func TestMaintainabilityNoEmptyBlockWhenExecutableStatementExists(t *testing.T) {
	source := "Sub Work()\nIf True Then\nDebug.Print 1\nElse\nDebug.Print 2\nEnd If\nEnd Sub\n"
	cfg := config.Default()
	cfg.Lint.DetectEmptyIf = true
	cfg.Lint.DetectEmptyElse = true
	doc, err := vbaast.ParseDocument("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	if err := doc.Read(func(view vbaast.ParsedView) error {
		if got := (Linter{RootDir: ".", Config: cfg}).maintainabilityIssues("Main.bas", view.Source, view.Root); len(got) != 0 {
			t.Errorf("nonempty branches reported empty: %+v", got)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

func TestMaintainabilityNestedIfKeepsOuterBranchNonempty(t *testing.T) {
	source := "Sub Work()\nIf True Then\nIf False Then\n' intentionally empty\nEnd If\nEnd If\nEnd Sub\n"
	cfg := config.Default()
	cfg.Lint.DetectEmptyIf = true
	doc, err := vbaast.ParseDocument("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	if err := doc.Read(func(view vbaast.ParsedView) error {
		got := (Linter{RootDir: ".", Config: cfg}).maintainabilityIssues("Main.bas", view.Source, view.Root)
		if len(got) != 1 || got[0].Code != "VB067" || got[0].Line != 3 {
			t.Errorf("nested If findings = %+v, want only inner branch", got)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

func TestMaintainabilityTypeSuffixSkipsIntrinsicCall(t *testing.T) {
	source := "Sub Work()\nDim amount$\namount$ = Chr$(10)\nEnd Sub\n"
	cfg := config.Default()
	cfg.Lint.DetectIdentifierTypeSuffix = true
	doc, err := vbaast.ParseDocument("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	if err := doc.Read(func(view vbaast.ParsedView) error {
		for _, issue := range (Linter{RootDir: ".", Config: cfg}).maintainabilityIssues("Main.bas", view.Source, view.Root) {
			if issue.Code == "VB081" && issue.Line == 3 && issue.Column > 10 {
				t.Errorf("intrinsic Chr$ call reported as type hint: %+v", issue)
			}
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

func TestMaintainabilityConditionalProcedureKeepsLocalScope(t *testing.T) {
	source := "#If VBA7 Then\nPublic Sub Work()\nDim leftValue As Long\n#Else\nPublic Sub Work()\nDim rightValue As Long\n#End If\nDim value As Long\nDebug.Print value\nEnd Sub\n"
	cfg := config.Default()
	cfg.Lint.DetectModuleDim = true
	cfg.Lint.DetectEmptyProcedure = true
	doc, err := vbaast.ParseDocument("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	if err := doc.Read(func(view vbaast.ParsedView) error {
		for _, issue := range (Linter{RootDir: ".", Config: cfg}).maintainabilityIssues("Main.bas", view.Source, view.Root) {
			if issue.Code == "VB087" || issue.Code == "VB074" {
				t.Errorf("conditional procedure finding = %+v", issue)
			}
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

func TestMaintainabilityImplicitPublicIgnoresBodyText(t *testing.T) {
	source := "Sub Work()\nDebug.Print \"Private\"\nEnd Sub\n"
	cfg := config.Default()
	cfg.Lint.DetectImplicitPublic = true
	doc, err := vbaast.ParseDocument("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	if err := doc.Read(func(view vbaast.ParsedView) error {
		got := (Linter{RootDir: ".", ModuleKind: "standard", Config: cfg}).maintainabilityIssues("Main.bas", view.Source, view.Root)
		if len(got) != 1 || got[0].Code != "VB088" || got[0].Line != 1 {
			t.Errorf("implicit Public finding = %+v", got)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

func TestMaintainabilityUnusedNumericLabels(t *testing.T) {
	source := "Sub Work()\n10:\n20 Debug.Print 1\nGoTo 20\nEnd Sub\n"
	cfg := config.Default()
	cfg.Lint.DetectUnusedLabels = true
	doc, err := vbaast.ParseDocument("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	if err := doc.Read(func(view vbaast.ParsedView) error {
		got := (Linter{RootDir: ".", Config: cfg}).maintainabilityIssues("Main.bas", view.Source, view.Root)
		if len(got) != 1 || got[0].Code != "VB090" || got[0].Line != 2 {
			t.Errorf("numeric label findings = %+v", got)
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}

func TestMaintainabilityTypeSuffixReportsDeclaredCallee(t *testing.T) {
	source := "Function Compute$()\nCompute$ = \"x\"\nEnd Function\nSub Work()\nDebug.Print Compute$()\nDebug.Print Chr$(65)\nEnd Sub\n"
	cfg := config.Default()
	cfg.Lint.DetectIdentifierTypeSuffix = true
	doc, err := vbaast.ParseDocument("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	if err := doc.Read(func(view vbaast.ParsedView) error {
		found := false
		for _, issue := range (Linter{RootDir: ".", Config: cfg}).maintainabilityIssues("Main.bas", view.Source, view.Root) {
			if issue.Code != "VB081" {
				continue
			}
			if issue.Line == 5 {
				found = true
			}
			if issue.Line == 6 {
				t.Errorf("intrinsic Chr$ reported: %+v", issue)
			}
		}
		if !found {
			t.Error("declared Compute$ call was not reported")
		}
		return nil
	}); err != nil {
		t.Fatal(err)
	}
}
