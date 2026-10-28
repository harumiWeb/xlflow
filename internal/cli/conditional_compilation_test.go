package cli

import (
	"bytes"
	"context"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestSourcePreflightConditionalCompilationPreservesDeliverySource(t *testing.T) {
	for _, tc := range []struct {
		name        string
		constants   map[string]string
		wantBlocked bool
	}{
		{name: "undefined custom constant"},
		{name: "explicit false", constants: map[string]string{"FeatureFlag": "False"}},
		{name: "explicit true preserves active error", constants: map[string]string{"FeatureFlag": "True"}, wantBlocked: true},
	} {
		t.Run(tc.name, func(t *testing.T) {
			dir := t.TempDir()
			path := filepath.Join("src", "modules", "CondComp.bas")
			source := "Attribute VB_Name = \"CondComp\"\nOption Explicit\n#If FeatureFlag Then\nThis is plain prose, not VBA code.\n## Heading\n1. item\n#End If\nPublic Function Answer() As Long\nAnswer = 42\nEnd Function\n"
			writePreflightWaiverModule(t, dir, path, source)
			cfg := config.Default()
			cfg.VBA.ConditionalConstants = tc.constants
			var stdout bytes.Buffer
			a := &app{cwd: dir, stdout: &stdout, stderr: &bytes.Buffer{}}
			err := a.runSourcePreflight(context.Background(), "push", cfg, "pushing to Excel", nil, nil)
			if (err != nil) != tc.wantBlocked {
				t.Fatalf("preflight error = %v, blocked = %t, output = %s", err, tc.wantBlocked, stdout.String())
			}
			if tc.wantBlocked && !strings.Contains(stdout.String(), "VB014") {
				t.Fatalf("active syntax recovery missing: %s", stdout.String())
			}
			if warnings := a.preflightWaiverWarnings(); len(warnings) != 0 {
				t.Fatalf("unexpected diagnostic waivers: %v", warnings)
			}
			got, readErr := os.ReadFile(filepath.Join(dir, path))
			if readErr != nil {
				t.Fatal(readErr)
			}
			if string(got) != source {
				t.Fatal("preflight changed the source that push must deliver")
			}
		})
	}
}
