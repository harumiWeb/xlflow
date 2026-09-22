package analyze

import (
	"context"
	"errors"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/sourceproject"
	"github.com/harumiWeb/xlflow/internal/vbadb"
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

func TestAnalyzerAnalyzeProjectUsesEmbeddedTypeDBWithoutRuntimeLookup(t *testing.T) {
	t.Setenv(typedb.EnvDir, t.TempDir())
	project := sourceproject.SourceProject{Files: []sourceproject.SourceFile{{
		Path:       "virtual/Main.bas",
		ModuleKind: sourceproject.ModuleKindStandard,
		Source: []byte(`Option Explicit
Public Sub Run()
    Dim rng As Range
    Dim result As Variant
    result = WorksheetFunction.Match("key", rng)
End Sub
`),
	}}}

	result, err := (Analyzer{
		Config:                 config.Default(),
		TypeDBGeneratorVersion: "stale-build",
	}).AnalyzeProject(t.Context(), project)
	if err != nil {
		t.Fatalf("AnalyzeProject: %v", err)
	}
	if got := findingsByCode(result.Findings, "VBA251"); len(got) != 1 {
		t.Fatalf("embedded TypeDB VBA251 findings = %+v, want one finding", got)
	}
	for _, warning := range result.Warnings {
		if warning["code"] == "type_db_load_warning" {
			t.Fatalf("in-memory analysis unexpectedly loaded the runtime TypeDB: %+v", warning)
		}
	}
}

func TestAnalyzerAnalyzeProjectUsesSuppliedTypeDBAndCompleteness(t *testing.T) {
	t.Setenv(typedb.EnvDir, t.TempDir())
	newDatabase := func() *vbadb.DB {
		db := vbadb.New()
		if err := db.MergeJSON([]byte(`{
  "types": [{
    "name": "Excel.WorksheetFunction",
    "library": "Excel",
    "kind": "interface",
    "confidence": "generated",
    "source": "typelib",
    "methods": [{ "name": "CustomFunction", "return_type": "Double" }]
  }],
  "global_values": {"WorksheetFunction": "Excel.WorksheetFunction"}
}`)); err != nil {
			t.Fatalf("MergeJSON: %v", err)
		}
		return db
	}
	project := func(member string) sourceproject.SourceProject {
		return sourceproject.SourceProject{Files: []sourceproject.SourceFile{{
			Path:       "virtual/Main.bas",
			ModuleKind: sourceproject.ModuleKindStandard,
			Source:     []byte("Option Explicit\nPublic Sub Run()\n    Dim result As Variant\n    result = WorksheetFunction." + member + "(1)\nEnd Sub\n"),
		}}}
	}

	complete := &TypeDatabase{DB: newDatabase(), Complete: true}
	result, err := (Analyzer{
		Config:                 config.Default(),
		TypeDB:                 complete,
		TypeDBGeneratorVersion: "stale-build",
	}).AnalyzeProject(t.Context(), project("CustomFunction"))
	if err != nil {
		t.Fatalf("AnalyzeProject with supplied TypeDB: %v", err)
	}
	if got := findingsByCode(result.Findings, "VBA252"); len(got) != 0 {
		t.Fatalf("supplied TypeDB member was not authoritative: %+v", got)
	}

	incomplete := &TypeDatabase{DB: newDatabase(), Complete: false}
	result, err = (Analyzer{
		Config:                 config.Default(),
		TypeDB:                 incomplete,
		TypeDBGeneratorVersion: "stale-build",
	}).AnalyzeProject(t.Context(), project("MissingFunction"))
	if err != nil {
		t.Fatalf("AnalyzeProject with incomplete supplied TypeDB: %v", err)
	}
	if got := findingsByCode(result.Findings, "VBA252"); len(got) != 0 {
		t.Fatalf("incomplete supplied TypeDB must fail open for absent members: %+v", got)
	}

	result, err = (Analyzer{
		Config:                 config.Default(),
		TypeDB:                 complete,
		TypeDBGeneratorVersion: "stale-build",
	}).AnalyzeProject(t.Context(), project("MissingFunction"))
	if err != nil {
		t.Fatalf("AnalyzeProject with complete supplied TypeDB: %v", err)
	}
	if got := findingsByCode(result.Findings, "VBA252"); len(got) != 1 {
		t.Fatalf("complete supplied TypeDB must report absent members: %+v", got)
	}
}

func TestAnalyzerAnalyzeProjectReportsDefaultMemberDiagnostics(t *testing.T) {
	t.Setenv(typedb.EnvDir, t.TempDir())
	db := vbadb.New()
	if err := db.MergeJSON([]byte(`{
  "types": [
    {
      "name": "Demo.ValueObject",
      "library": "Demo",
      "kind": "interface",
      "confidence": "generated",
      "source": "typelib",
      "default_member": "Value",
      "default_member_type": "String",
      "properties": [{"name": "Value", "return_type": "String", "default": true}]
    },
    {
      "name": "Demo.Items",
      "library": "Demo",
      "kind": "interface",
      "confidence": "generated",
      "source": "typelib",
      "default_member": "Item",
      "default_member_type": "String",
      "properties": [{"name": "Item", "return_type": "String", "default": true, "parameters": [{"name": "Index", "type": "Long"}]}]
    }
  ]
}`)); err != nil {
		t.Fatalf("MergeJSON: %v", err)
	}
	project := sourceproject.SourceProject{Files: []sourceproject.SourceFile{{
		Path:       "virtual/Main.bas",
		ModuleKind: sourceproject.ModuleKindStandard,
		Source: []byte(`Option Explicit
Public Sub Run()
    Dim valueObject As Demo.ValueObject
    Dim items As Demo.Items
    Dim result As Variant
    result = valueObject
    result = items(1)
End Sub
`),
	}}}

	cfg := config.Default()
	cfg.Analyze.DetectImplicitDefaultMemberAccess = true
	result, err := (Analyzer{
		Config: cfg,
		TypeDB: &TypeDatabase{DB: db, Complete: true},
	}).AnalyzeProject(t.Context(), project)
	if err != nil {
		t.Fatalf("AnalyzeProject: %v", err)
	}
	findings := findingsByCode(result.Findings, "VBA253")
	if len(findings) != 2 {
		t.Fatalf("VBA253 findings = %+v, want implicit and indexed findings", findings)
	}
	if findings[0].DefaultMember == nil || findings[0].DefaultMember.Kind != "implicit" || findings[0].DefaultMember.Member != "Value" {
		t.Fatalf("implicit default-member context = %+v", findings[0].DefaultMember)
	}
	if findings[1].DefaultMember == nil || findings[1].DefaultMember.Kind != "indexed" || findings[1].DefaultMember.Member != "Item" {
		t.Fatalf("indexed default-member context = %+v", findings[1].DefaultMember)
	}
}

func TestAnalyzerAnalyzeProjectKeepsAdvisoryDefaultMemberRulesOptIn(t *testing.T) {
	t.Setenv(typedb.EnvDir, t.TempDir())
	project := sourceproject.SourceProject{Files: []sourceproject.SourceFile{{
		Path:       "virtual/Main.bas",
		ModuleKind: sourceproject.ModuleKindStandard,
		Source: []byte(`Option Explicit
Public Sub Run()
    Dim lateBound As Object
    Dim result As Variant
    result = lateBound
    result = lateBound!Name
End Sub
`),
	}}}

	base := config.Default()
	result, err := (Analyzer{Config: base}).AnalyzeProject(t.Context(), project)
	if err != nil {
		t.Fatalf("AnalyzeProject defaults: %v", err)
	}
	if got := append(findingsByCode(result.Findings, "VBA254"), findingsByCode(result.Findings, "VBA255")...); len(got) != 0 {
		t.Fatalf("opt-in findings enabled by default: %+v", got)
	}

	base.Analyze.DetectUnboundDefaultMemberAccess = true
	base.Analyze.DetectBangNotation = true
	result, err = (Analyzer{Config: base}).AnalyzeProject(t.Context(), project)
	if err != nil {
		t.Fatalf("AnalyzeProject opt-in: %v", err)
	}
	if got := findingsByCode(result.Findings, "VBA254"); len(got) != 1 || got[0].DefaultMember == nil || got[0].DefaultMember.Binding != "unbound" {
		t.Fatalf("VBA254 findings = %+v", got)
	}
	if got := findingsByCode(result.Findings, "VBA255"); len(got) != 1 || got[0].DefaultMember == nil || got[0].DefaultMember.Kind != "bang" {
		t.Fatalf("VBA255 findings = %+v", got)
	}
}

func TestAnalyzerAnalyzeProjectDefaultMemberAbsenceRequiresCompleteGeneratedMetadata(t *testing.T) {
	t.Setenv(typedb.EnvDir, t.TempDir())
	newDatabase := func() *vbadb.DB {
		db := vbadb.New()
		if err := db.MergeJSON([]byte(`{
  "types": [{
    "name": "Demo.NoDefault",
    "library": "Demo",
    "kind": "interface",
    "confidence": "generated",
    "source": "typelib",
    "properties": [{"name": "Name", "return_type": "String"}]
  }]
}`)); err != nil {
			t.Fatalf("MergeJSON: %v", err)
		}
		return db
	}
	project := sourceproject.SourceProject{Files: []sourceproject.SourceFile{{
		Path:       "virtual/Main.bas",
		ModuleKind: sourceproject.ModuleKindStandard,
		Source: []byte(`Option Explicit
Public Sub Run()
    Dim receiver As Demo.NoDefault
    Dim result As Variant
    Set receiver = New Demo.NoDefault
    result = receiver
End Sub
`),
	}}}

	for _, test := range []struct {
		name     string
		complete bool
		want     int
	}{
		{name: "incomplete fails open", complete: false, want: 0},
		{name: "complete proves absence", complete: true, want: 1},
	} {
		t.Run(test.name, func(t *testing.T) {
			cfg := config.Default()
			cfg.Analyze.DetectImplicitDefaultMemberAccess = true
			result, err := (Analyzer{
				Config: cfg,
				TypeDB: &TypeDatabase{DB: newDatabase(), Complete: test.complete},
			}).AnalyzeProject(t.Context(), project)
			if err != nil {
				t.Fatalf("AnalyzeProject: %v", err)
			}
			findings := findingsByCode(result.Findings, "VBA249")
			if len(findings) != test.want {
				t.Fatalf("VBA249 findings = %+v, want %d", findings, test.want)
			}
			if test.want == 1 && (findings[0].RuntimeError == nil || findings[0].RuntimeError.Kind != "default_member_required") {
				t.Fatalf("runtime context = %+v", findings[0].RuntimeError)
			}
		})
	}
}

func TestAnalyzerAnalyzeProjectResolvesProjectDefaultMemberAttribute(t *testing.T) {
	t.Setenv(typedb.EnvDir, t.TempDir())
	cfg := config.Default()
	cfg.Analyze.DetectImplicitDefaultMemberAccess = true
	project := sourceproject.SourceProject{Files: []sourceproject.SourceFile{
		{
			Path:       "virtual/Entry.cls",
			ModuleKind: sourceproject.ModuleKindClass,
			Source: []byte(`VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Entry"
Option Explicit
Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    Value = "ok"
End Property
`),
		},
		{
			Path:       "virtual/Main.bas",
			ModuleKind: sourceproject.ModuleKindStandard,
			Source: []byte(`Option Explicit
Public Sub Run()
    Dim entryValue As Entry
    Dim result As Variant
    Set entryValue = New Entry
    result = entryValue
End Sub
`),
		},
	}}

	result, err := (Analyzer{Config: cfg}).AnalyzeProject(t.Context(), project)
	if err != nil {
		t.Fatalf("AnalyzeProject: %v", err)
	}
	findings := findingsByCode(result.Findings, "VBA253")
	if len(findings) != 1 || findings[0].DefaultMember == nil || findings[0].DefaultMember.Member != "Value" {
		t.Fatalf("project default-member findings = %+v", findings)
	}
}

func TestAnalyzerAnalyzeProjectDerivesModuleNameFromLogicalPathSeparators(t *testing.T) {
	t.Setenv(typedb.EnvDir, t.TempDir())
	project := sourceproject.SourceProject{Files: []sourceproject.SourceFile{{
		Path:       `virtual\Main.bas`,
		ModuleKind: sourceproject.ModuleKindStandard,
		Source:     []byte("Option Explicit\nPublic Sub Run()\n  Range(\"A1\").Value = 1\nEnd Sub\n"),
	}}}

	result, err := (Analyzer{Config: config.Default()}).AnalyzeProject(t.Context(), project)
	if err != nil {
		t.Fatalf("AnalyzeProject: %v", err)
	}
	findings := findingsByCode(result.Findings, "VBA205")
	if len(findings) != 1 || findings[0].Module != "Main" {
		t.Fatalf("logical-path module = %+v, want one VBA205 finding for Main", findings)
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
