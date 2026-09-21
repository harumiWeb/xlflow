package analyze

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/sourceproject"
)

func TestAnalyzeProjectUserFormDoesNotReadHostDesigner(t *testing.T) {
	root := t.TempDir()
	formsDir := filepath.Join(root, "src", "forms")
	if err := os.MkdirAll(formsDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(formsDir, "Dialog.frm"), []byte(`VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dialog
   Begin MSForms.TextBox TestButton
   End
End
`), 0o644); err != nil {
		t.Fatal(err)
	}

	result, err := (Analyzer{RootDir: root, Config: config.Default()}).AnalyzeProject(t.Context(), sourceproject.SourceProject{Files: []sourceproject.SourceFile{{
		Path:       "virtual/forms/code/Dialog.bas",
		ModuleKind: sourceproject.ModuleKindForm,
		Source: []byte(`Option Explicit
Private Sub TestButton_Change()
	Me.TestButton.Value = "updated"
End Sub
`),
	}}})
	if err != nil {
		t.Fatalf("AnalyzeProject: %v", err)
	}
	if got := findingsByCode(result.Findings, "VBA220"); len(got) != 0 {
		t.Fatalf("host-only UserForm metadata changed source-only findings: %+v", got)
	}
	if len(result.Warnings) != 1 {
		t.Fatalf("warnings = %+v, want one capability warning", result.Warnings)
	}
	warning := result.Warnings[0]
	if warning["code"] != "analysis_capability_unavailable" || warning["capability"] != "userform_control_metadata" || warning["file"] != "virtual/forms/code/Dialog.bas" {
		t.Fatalf("capability warning = %+v", warning)
	}
	if got, ok := warning["rules"].([]string); !ok || len(got) != 1 || got[0] != "VBA220" {
		t.Fatalf("capability warning rules = %#v, want [VBA220]", warning["rules"])
	}
}

func TestAnalyzeProjectUsesSuppliedUserFormDesignerSource(t *testing.T) {
	project := sourceproject.SourceProject{Files: []sourceproject.SourceFile{
		{
			Path:       `virtual\forms\code\Dialog.bas`,
			ModuleKind: sourceproject.ModuleKindForm,
			Source: []byte(`Option Explicit
Private Sub TextBox1_Change()
  Application.EnableEvents = False
  Me.TextBox1.Value = "updated"
  Application.EnableEvents = True
End Sub
`),
		},
		{
			Path:       `virtual\forms\Dialog.frm`,
			ModuleKind: sourceproject.ModuleKindForm,
			Source: []byte(`VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dialog
	   Begin MSForms.TextBox TextBox1
   End
End
`),
		},
	}}

	result, err := (Analyzer{Config: config.Default()}).AnalyzeProject(t.Context(), project)
	if err != nil {
		t.Fatalf("AnalyzeProject: %v", err)
	}
	if got := findingsByCode(result.Findings, "VBA220"); len(got) != 1 || got[0].Procedure != "TextBox1_Change" {
		t.Fatalf("supplied UserForm designer did not classify the event: %+v", got)
	}
	for _, warning := range result.Warnings {
		if warning["capability"] == "userform_control_metadata" {
			t.Fatalf("supplied UserForm designer unexpectedly produced capability warning: %+v", warning)
		}
	}
}

func TestAnalyzeProjectDoesNotReadUserFormFRXReference(t *testing.T) {
	project := sourceproject.SourceProject{Files: []sourceproject.SourceFile{
		{
			Path:       `virtual\forms\code\Dialog.bas`,
			ModuleKind: sourceproject.ModuleKindForm,
			Source: []byte(`Option Explicit
Private Sub TestButton_Change()
  Me.TestButton.Value = "updated"
End Sub
`),
		},
		{
			Path:       `virtual\forms\Dialog.frm`,
			ModuleKind: sourceproject.ModuleKindForm,
			Source: []byte(`VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dialog
   OleObjectBlob   =   "Dialog.frx":0000
   Begin MSForms.TextBox TestButton
   End
End
`),
		},
	}}

	result, err := (Analyzer{Config: config.Default()}).AnalyzeProject(t.Context(), project)
	if err != nil {
		t.Fatalf("AnalyzeProject: %v", err)
	}
	if got := findingsByCode(result.Findings, "VBA220"); len(got) != 0 {
		t.Fatalf("external FRX metadata unexpectedly made an event certain: %+v", got)
	}
	if len(result.Warnings) != 1 || result.Warnings[0]["capability"] != "userform_control_metadata" {
		t.Fatalf("FRX capability warning = %+v, want one userform metadata warning", result.Warnings)
	}
}

func TestAnalyzeProjectUserFormCapabilityWarningsArePathOrdered(t *testing.T) {
	project := sourceproject.SourceProject{Files: []sourceproject.SourceFile{
		{
			Path:       "virtual/z/forms/code/Zed.bas",
			ModuleKind: sourceproject.ModuleKindForm,
			Source: []byte(`Private Sub TestZed_Change()
  Application.EnableEvents = False
  Me.TestZed.Value = "updated"
  Application.EnableEvents = True
End Sub
`),
		},
		{
			Path:       "virtual/a/forms/code/Alpha.bas",
			ModuleKind: sourceproject.ModuleKindForm,
			Source: []byte(`Private Sub TestAlpha_Change()
  Application.EnableEvents = False
  Me.TestAlpha.Value = "updated"
  Application.EnableEvents = True
End Sub
`),
		},
	}}

	result, err := (Analyzer{Config: config.Default()}).AnalyzeProject(t.Context(), project)
	if err != nil {
		t.Fatalf("AnalyzeProject: %v", err)
	}
	if len(result.Warnings) != 2 {
		t.Fatalf("warnings = %+v, want two capability warnings", result.Warnings)
	}
	if result.Warnings[0]["file"] != "virtual/a/forms/code/Alpha.bas" || result.Warnings[1]["file"] != "virtual/z/forms/code/Zed.bas" {
		t.Fatalf("capability warning order = %+v", result.Warnings)
	}
}
