package cli

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/output"
	staticrules "github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
)

func TestRulesCommandWritesV1JSONEnvelope(t *testing.T) {
	var stdout bytes.Buffer
	a := &app{cwd: t.TempDir(), stdout: &stdout, stderr: &bytes.Buffer{}}
	root := a.rootCommand()
	root.SetArgs([]string{"--json", "rules"})

	if err := root.Execute(); err != nil {
		t.Fatalf("rules command error = %v, exit = %d", err, output.ExitCode(err))
	}
	var got struct {
		Status  string              `json:"status"`
		Command string              `json:"command"`
		Rules   staticrules.Catalog `json:"rules"`
	}
	if err := json.Unmarshal(stdout.Bytes(), &got); err != nil {
		t.Fatalf("decode rules JSON: %v\n%s", err, stdout.String())
	}
	if got.Status != output.StatusOK || got.Command != "rules" || got.Rules.SchemaVersion != staticrules.SchemaVersion {
		t.Fatalf("rules envelope = %#v", got)
	}
	if len(got.Rules.Items) == 0 || got.Rules.Items[0].ID == "" {
		t.Fatalf("rules catalog is empty: %#v", got.Rules)
	}
	for _, item := range got.Rules.Items {
		if item.EvidenceClass == "" || len(item.Surfaces) == 0 || len(item.SupportedSeverities) == 0 {
			t.Fatalf("rules metadata missing additive surface/severity fields for %s: %#v", item.ID, item)
		}
	}
	for _, id := range []string{"VB037", "VB045", "VB046", "VB047", "VB048", "VB049", "VBA228", "VBA229"} {
		found := false
		for _, item := range got.Rules.Items {
			if item.ID == id {
				found = true
				if !item.CompileEquivalent || item.DefaultSeverity != staticrules.SeverityError || !item.PreflightBlocking || item.InlineSuppressible {
					t.Fatalf("compile-equivalent metadata for %s = %#v", id, item)
				}
			}
		}
		if !found {
			t.Fatalf("rules metadata missing %s", id)
		}
	}
}

func TestRulesCommandWritesDeterministicHumanInventory(t *testing.T) {
	var stdout bytes.Buffer
	a := &app{cwd: t.TempDir(), stdout: &stdout, stderr: &bytes.Buffer{}}
	root := a.rootCommand()
	root.SetArgs([]string{"rules"})

	if err := root.Execute(); err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"VB001", "lint", "file-local", "Missing Option Explicit"} {
		if !strings.Contains(stdout.String(), want) {
			t.Fatalf("rules output missing %q:\n%s", want, stdout.String())
		}
	}
}
