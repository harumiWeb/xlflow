package main

import (
	"bytes"
	"context"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/staticanalysis/corpus"
)

func TestRunRejectsUnsafeOrIncompleteDeveloperCommands(t *testing.T) {
	tests := []struct {
		args []string
		want string
	}{
		{args: nil, want: "usage:"},
		{args: []string{"review-details"}, want: "--rule is required"},
		{args: []string{"review-draft", "--rule", "VBA225"}, want: "--classification is required"},
		{args: []string{"review-draft", "--rule", "VBA225", "--classification", "true-positive"}, want: "--rationale is required"},
		{args: []string{"review-draft", "--rule", "VBA225", "--classification", "false-positive", "--rationale", "reviewed"}, want: "requires exactly one"},
		{args: []string{"verify"}, want: "focused verify requires"},
		{args: []string{"verify", "--rule", "VBA999"}, want: "unknown diagnostic rule"},
	}
	for _, test := range tests {
		var stdout, stderr bytes.Buffer
		if code := run(context.Background(), test.args, &stdout, &stderr); code == 0 {
			t.Fatalf("run(%q) succeeded", test.args)
		}
		if !strings.Contains(stderr.String(), test.want) {
			t.Fatalf("run(%q) stderr = %q, want %q", test.args, stderr.String(), test.want)
		}
	}
}

func TestSelectSnapshotIDsUsesCommittedSelectionRatherThanActualResults(t *testing.T) {
	ids := []corpus.SnapshotID{
		{Project: "self/selected", Surface: corpus.SurfaceLint},
		{Project: "self/selected", Surface: corpus.SurfaceAnalyze},
		{Project: "self/other", Surface: corpus.SurfaceAnalyze},
	}
	selected := selectSnapshotIDs(ids, "self/selected")
	if len(selected) != 2 || selected[0].Project != "self/selected" || selected[1].Project != "self/selected" {
		t.Fatalf("selected IDs = %#v", selected)
	}
	if all := selectSnapshotIDs(ids, ""); len(all) != len(ids) {
		t.Fatalf("all IDs = %#v", all)
	}
	want := corpus.SnapshotSet{
		selected[0]: {{Project: selected[0].Project, Surface: selected[0].Surface, File: "src/Main.bas", Code: "VBA225", Severity: "warning", Line: 1}},
		selected[1]: {},
	}
	actual := corpus.SnapshotSet{selected[1]: {}}
	if diff := corpus.CompareSnapshotSets(want, actual); len(diff.Removed) != 1 {
		t.Fatalf("missing committed surface was not detected: %#v", diff)
	}
}
