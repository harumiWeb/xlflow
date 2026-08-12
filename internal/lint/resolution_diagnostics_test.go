package lint

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestResolutionDiagnosticsProjectNegativeContracts(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(src, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Public Event Changed(ByVal value As Long)
Public Const Scalar = 1
Sub Run()
    Scalar()
    RaiseEvent Missing(1)
End Sub
`
	if err := os.WriteFile(filepath.Join(src, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	issues, err := (Linter{RootDir: dir, Config: config.Default()}).Run()
	if err != nil {
		t.Fatal(err)
	}
	seen := map[string]bool{}
	for _, issue := range issues {
		if issue.Code == "VB052" || issue.Code == "VB054" {
			seen[issue.Code] = true
		}
	}
	if !seen["VB052"] || !seen["VB054"] {
		t.Fatalf("resolution diagnostics = %+v, want VB052 and VB054", issues)
	}
}
