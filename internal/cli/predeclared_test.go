package cli

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestSourcePreflightPredeclaredClassAssignment(t *testing.T) {
	for _, value := range []string{"True", "False"} {
		t.Run(value, func(t *testing.T) {
			dir := t.TempDir()
			for path, source := range map[string]string{
				"src/classes/Predeclared.cls": "Attribute VB_Name = \"Predeclared\"\nAttribute VB_PredeclaredId = " + value + "\nOption Explicit\n",
				"src/modules/Main.bas":        "Attribute VB_Name = \"Main\"\nOption Explicit\nPublic Sub Run()\n  Set Predeclared = Nothing\nEnd Sub\n",
			} {
				path = filepath.Join(dir, filepath.FromSlash(path))
				if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
					t.Fatal(err)
				}
				if err := os.WriteFile(path, []byte(source), 0o644); err != nil {
					t.Fatal(err)
				}
			}
			var stdout bytes.Buffer
			a := &app{cwd: dir, json: true, stdout: &stdout, stderr: &bytes.Buffer{}}
			err := a.runSourcePreflight(t.Context(), "push", config.Default(), "pushing", nil, nil)
			if value == "True" {
				if err != nil {
					t.Fatalf("predeclared instance blocked: %v\n%s", err, stdout.String())
				}
			} else if err == nil || !strings.Contains(stdout.String(), "VB029") || !strings.Contains(stdout.String(), "preflight") {
				t.Fatalf("ordinary class not blocked: %v\n%s", err, stdout.String())
			}
		})
	}
}
