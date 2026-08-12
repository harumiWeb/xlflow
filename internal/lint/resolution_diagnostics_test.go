package lint

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestResolutionDiagnosticsProjectNegativeContracts(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "classes")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Public Event Changed(ByVal value As Long)
Private Const Scalar = 1
Sub Run()
    Scalar()
    RaiseEvent Missing(1)
End Sub
`
	if err := os.WriteFile(filepath.Join(src, "Main.cls"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := (Linter{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	seen := map[string]bool{}
	for _, issue := range issues {
		switch issue.Code {
		case "VB052":
			seen[issue.Code] = true
			if issue.Line != 4 || issue.Column != 5 {
				t.Fatalf("VB052 range = %d:%d, want 4:5", issue.Line, issue.Column)
			}
		case "VB054":
			seen[issue.Code] = true
			if issue.Line != 5 || issue.Column != 16 {
				t.Fatalf("VB054 range = %d:%d, want 5:16", issue.Line, issue.Column)
			}
		}
	}
	if !seen["VB052"] || !seen["VB054"] {
		t.Fatalf("resolution diagnostics = %+v, want VB052 and VB054", issues)
	}
}
