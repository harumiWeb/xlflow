package lint

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestLinterVB029PredeclaredClassAssignment(t *testing.T) {
	t.Parallel()
	for _, tc := range []struct {
		name      string
		attribute string
		local     string
		want      int
	}{
		{name: "default instance", attribute: "Attribute VB_PredeclaredId = True\n"},
		{name: "case insensitive", attribute: "Attribute vb_predeclaredid = true\n"},
		{name: "ordinary class", attribute: "Attribute VB_PredeclaredId = False\n", want: 1},
		{name: "attribute absent", want: 1},
		{name: "local shadows ordinary class", local: "  Dim Predeclared As Object\n"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			dir := t.TempDir()
			classes := filepath.Join(dir, "src", "classes")
			if err := os.MkdirAll(classes, 0o755); err != nil {
				t.Fatal(err)
			}
			// The declared module name, rather than the filename, owns the instance.
			if err := os.WriteFile(filepath.Join(classes, "DifferentFilename.cls"), []byte("Attribute VB_Name = \"Predeclared\"\n"+tc.attribute+"Option Explicit\nPublic Value As Long\n"), 0o644); err != nil {
				t.Fatal(err)
			}
			writeLintModule(t, dir, "UsePredeclared.bas", "Attribute VB_Name = \"UsePredeclared\"\nOption Explicit\nPublic Sub Run()\n"+tc.local+"  Predeclared.Value = 1\n  Debug.Print Predeclared.Value\n  Set pReDeClArEd = Nothing\n  Set missingValue = Nothing\nEnd Sub\n")
			issues, err := (Linter{RootDir: dir, Config: config.Default()}).Run()
			if err != nil {
				t.Fatal(err)
			}
			got := issuesByCode(issues, "VB029")
			if len(got) != tc.want+1 || got[len(got)-1].Symbol != "missingValue" {
				t.Fatalf("VB029 = %+v, want %d class errors and one missingValue", got, tc.want)
			}
			if tc.want > 0 && got[0].Symbol != "pReDeClArEd" {
				t.Fatalf("VB029 = %+v, want ordinary class assignment error", got)
			}
			if blocking := issuesByCode(PushBlockingIssues(issues), "VB029"); len(blocking) != tc.want+1 {
				t.Fatalf("blocking VB029 = %+v", blocking)
			}
		})
	}
}
