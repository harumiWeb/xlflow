package analyze

import (
	"context"
	"errors"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/sourceproject"
)

func TestAnalyzerAnalyzeProjectUsesCallerSuppliedVirtualFiles(t *testing.T) {
	t.Setenv(typedb.EnvDir, t.TempDir())

	project := sourceproject.SourceProject{Files: []sourceproject.SourceFile{
		{
			Path:       "virtual/Main.bas",
			Source:     []byte("Option Explicit\nPublic Sub Run()\n  Dim count As Long\n  Receiver.ReplaceText count\nEnd Sub\n"),
			ModuleKind: sourceproject.ModuleKindStandard,
		},
		{
			Path:       "virtual/Receiver.bas",
			Source:     []byte("Option Explicit\nPublic Sub ReplaceText(ByRef target As String)\nEnd Sub\n"),
			ModuleKind: sourceproject.ModuleKindStandard,
		},
	}}

	result, err := (Analyzer{Config: config.Default()}).AnalyzeProject(t.Context(), project)
	if err != nil {
		t.Fatalf("AnalyzeProject: %v", err)
	}
	if result.AnalyzedFiles != 2 {
		t.Fatalf("analyzed files = %d, want 2", result.AnalyzedFiles)
	}
	if got := findingsByCode(result.Findings, "VBA228"); len(got) != 1 || got[0].File != "virtual/Main.bas" {
		t.Fatalf("virtual cross-module ByRef finding = %+v", got)
	}
}

func TestAnalyzerAnalyzeProjectAppliesSourceSuppressions(t *testing.T) {
	t.Setenv(typedb.EnvDir, t.TempDir())
	project := sourceproject.SourceProject{Files: []sourceproject.SourceFile{{
		Path:       "virtual/Main.bas",
		ModuleKind: sourceproject.ModuleKindStandard,
		Source:     []byte("Option Explicit\nPublic Sub Run()\n  ' xlflow:disable-next-line VBA205\n  Range(\"A1\").Value = 1\nEnd Sub\n"),
	}}}

	result, err := (Analyzer{Config: config.Default()}).AnalyzeProject(t.Context(), project)
	if err != nil {
		t.Fatalf("AnalyzeProject: %v", err)
	}
	if got := findingsByCode(result.Findings, "VBA205"); len(got) != 0 {
		t.Fatalf("VBA205 should be suppressed from virtual source: %+v", got)
	}
}

func TestAnalyzerAnalyzeProjectValidatesInputMetadata(t *testing.T) {
	t.Setenv(typedb.EnvDir, t.TempDir())
	tests := []struct {
		name    string
		project sourceproject.SourceProject
	}{
		{
			name: "empty path",
			project: sourceproject.SourceProject{Files: []sourceproject.SourceFile{{
				Path: "", Source: []byte("Option Explicit\n"), ModuleKind: sourceproject.ModuleKindStandard,
			}}},
		},
		{
			name: "unsupported module kind",
			project: sourceproject.SourceProject{Files: []sourceproject.SourceFile{{
				Path: "virtual/Main.bas", Source: []byte("Option Explicit\n"), ModuleKind: sourceproject.ModuleKind("unknown"),
			}}},
		},
		{
			name: "duplicate logical path",
			project: sourceproject.SourceProject{Files: []sourceproject.SourceFile{
				{Path: "virtual/Main.bas", Source: []byte("Option Explicit\n"), ModuleKind: sourceproject.ModuleKindStandard},
				{Path: "VIRTUAL\\main.bas", Source: []byte("Option Explicit\n"), ModuleKind: sourceproject.ModuleKindStandard},
			}},
		},
		{
			name: "test must be standard module",
			project: sourceproject.SourceProject{Files: []sourceproject.SourceFile{{
				Path: "virtual/Form.frm", Source: []byte("Option Explicit\n"), ModuleKind: sourceproject.ModuleKindForm, IsTest: true,
			}}},
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			if _, err := (Analyzer{Config: config.Default()}).AnalyzeProject(t.Context(), test.project); err == nil {
				t.Fatal("AnalyzeProject unexpectedly accepted invalid source project")
			}
		})
	}
}

func TestAnalyzerAnalyzeProjectPreservesParseErrorPath(t *testing.T) {
	t.Setenv(typedb.EnvDir, t.TempDir())
	_, err := (Analyzer{Config: config.Default()}).AnalyzeProject(t.Context(), sourceproject.SourceProject{Files: []sourceproject.SourceFile{{
		Path:       "virtual/Broken.bas",
		Source:     []byte("Option Explicit\nPublic Sub Broken(\nEnd Sub\n"),
		ModuleKind: sourceproject.ModuleKindStandard,
	}}})
	parseErr, ok := errors.AsType[*ParseError](err)
	if !ok {
		t.Fatalf("error = %T %v, want ParseError", err, err)
	}
	if parseErr.Path != "virtual/Broken.bas" {
		t.Fatalf("parse error path = %q, want virtual/Broken.bas", parseErr.Path)
	}
}

func TestAnalyzerAnalyzeProjectReturnsCancellation(t *testing.T) {
	ctx, cancel := context.WithCancel(t.Context())
	cancel()

	_, err := (Analyzer{Config: config.Default()}).AnalyzeProject(ctx, sourceproject.SourceProject{})
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("AnalyzeProject error = %v, want context.Canceled", err)
	}
}
