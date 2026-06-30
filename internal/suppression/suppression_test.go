package suppression

import (
	"os"
	"path/filepath"
	"testing"
)

func TestDirectivesForFilesParsesInlineSuppressions(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "Main.bas")
	body := `Option Explicit
' xlflow:disable-next-line vb002 VBA205
Range("A1").Select
Debug.Print "not a ' xlflow:disable-line VB003 comment"
ActiveCell.Activate ' xlflow:disable-line VB003
`
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}

	directives, warnings, err := DirectivesForFiles(dir, []string{path})
	if err != nil {
		t.Fatal(err)
	}
	if len(warnings) != 0 {
		t.Fatalf("warnings = %+v, want none", warnings)
	}
	if len(directives) != 3 {
		t.Fatalf("directives = %+v, want 3", directives)
	}
	assertDirective(t, directives, "VB002", 2, 3)
	assertDirective(t, directives, "VBA205", 2, 3)
	assertDirective(t, directives, "VB003", 5, 5)
}

func TestDirectivesForFilesReportsUnknownIDs(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "Main.bas")
	if err := os.WriteFile(path, []byte("' xlflow:disable-next-line VB999\n"), 0o644); err != nil {
		t.Fatal(err)
	}

	directives, warnings, err := DirectivesForFiles(dir, []string{path})
	if err != nil {
		t.Fatal(err)
	}
	if len(directives) != 0 {
		t.Fatalf("directives = %+v, want none", directives)
	}
	if len(warnings) != 1 || warnings[0]["code"] != "unknown_inline_suppression_rule" || warnings[0]["rule"] != "VB999" {
		t.Fatalf("warnings = %+v", warnings)
	}
}

func TestDirectivesForFilesRejectsPreflightBlockingIDs(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "Main.bas")
	body := `' xlflow:disable-next-line VB008 VB031 VB032 VBA104 VBA211
`
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}

	directives, warnings, err := DirectivesForFiles(dir, []string{path})
	if err != nil {
		t.Fatal(err)
	}
	if len(directives) != 0 {
		t.Fatalf("directives = %+v, want none", directives)
	}
	for _, rule := range []string{"VB008", "VB031", "VB032", "VBA104", "VBA211"} {
		if !hasWarning(warnings, "unsupported_inline_suppression_rule", rule) {
			t.Fatalf("missing unsupported warning for %s in %+v", rule, warnings)
		}
	}
}

func TestApplyReportsUnusedSuppressionsForCurrentFamilyOnly(t *testing.T) {
	directives := []Directive{
		{Code: "VB002", File: "src/modules/Main.bas", Line: 1, TargetLine: 2},
		{Code: "VBA205", File: "src/modules/Main.bas", Line: 1, TargetLine: 2},
	}
	diagnostics := []Diagnostic{{Code: "VB002", File: "src/modules/Main.bas", Line: 2}}

	suppressed, warnings := Apply(diagnostics, directives, FamilyLint)
	if len(suppressed) != 1 || !suppressed[0] {
		t.Fatalf("suppressed = %+v, want diagnostic suppressed", suppressed)
	}
	if len(warnings) != 0 {
		t.Fatalf("lint warnings = %+v, want none", warnings)
	}

	_, warnings = Apply(nil, directives, FamilyAnalyze)
	if len(warnings) != 1 || warnings[0]["rule"] != "VBA205" {
		t.Fatalf("analyze warnings = %+v, want unused VBA205 only", warnings)
	}
}

func hasWarning(warnings []map[string]any, code string, rule string) bool {
	for _, warning := range warnings {
		if warning["code"] == code && warning["rule"] == rule {
			return true
		}
	}
	return false
}

func assertDirective(t *testing.T, directives []Directive, code string, line int, targetLine int) {
	t.Helper()
	for _, directive := range directives {
		if directive.Code == code && directive.Line == line && directive.TargetLine == targetLine {
			return
		}
	}
	t.Fatalf("missing directive %s line %d target %d in %+v", code, line, targetLine, directives)
}
