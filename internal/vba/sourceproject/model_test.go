package sourceproject_test

import (
	"bytes"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/sourceproject"
)

func TestSourceProjectRepresentsVirtualMultiModuleProject(t *testing.T) {
	project := sourceproject.SourceProject{Files: []sourceproject.SourceFile{
		{
			Path:       "Module1.bas",
			Source:     []byte("Public Sub Run()\nEnd Sub\n"),
			ModuleKind: sourceproject.ModuleKindStandard,
		},
		{
			Path:       "Customer.cls",
			Source:     []byte("Public Name As String\n"),
			ModuleKind: sourceproject.ModuleKindClass,
		},
		{
			Path:       "LoginForm.frm",
			Source:     []byte("Private Sub Submit_Click()\nEnd Sub\n"),
			ModuleKind: sourceproject.ModuleKindForm,
		},
		{
			Path:       "ThisWorkbook.cls",
			Source:     []byte("Private Sub Workbook_Open()\nEnd Sub\n"),
			ModuleKind: sourceproject.ModuleKindDocument,
		},
		{
			Path:       "Sheet1.cls",
			Source:     []byte("Private Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\n"),
			ModuleKind: sourceproject.ModuleKindDocument,
		},
		{
			Path:       "tests/Module1Tests.bas",
			Source:     []byte("Public Sub TestRun()\nEnd Sub\n"),
			ModuleKind: sourceproject.ModuleKindStandard,
			IsTest:     true,
		},
	}}

	if len(project.Files) != 6 {
		t.Fatalf("files = %d, want 6", len(project.Files))
	}

	testModule := project.Files[5]
	if testModule.Path != "tests/Module1Tests.bas" {
		t.Fatalf("test module path = %q", testModule.Path)
	}
	if testModule.ModuleKind != sourceproject.ModuleKindStandard || !testModule.IsTest {
		t.Fatalf("test module metadata = %+v", testModule)
	}
	if !bytes.Equal(testModule.Source, []byte("Public Sub TestRun()\nEnd Sub\n")) {
		t.Fatalf("test module source = %q", testModule.Source)
	}

	if project.Files[3].ModuleKind != sourceproject.ModuleKindDocument ||
		project.Files[4].ModuleKind != sourceproject.ModuleKindDocument {
		t.Fatalf("document module kinds = %q, %q", project.Files[3].ModuleKind, project.Files[4].ModuleKind)
	}
}

func TestModuleKindValuesMatchExistingAnalysisVocabulary(t *testing.T) {
	tests := map[sourceproject.ModuleKind]string{
		sourceproject.ModuleKindStandard: "standard",
		sourceproject.ModuleKindClass:    "class",
		sourceproject.ModuleKindForm:     "form",
		sourceproject.ModuleKindDocument: "document",
	}
	for kind, want := range tests {
		if string(kind) != want {
			t.Errorf("module kind = %q, want %q", kind, want)
		}
	}
}
