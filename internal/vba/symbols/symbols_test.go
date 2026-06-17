package symbols

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestInspectExtractsRepresentativeStandardModuleSymbols(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	moduleDir := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `Attribute VB_Name = "Main"
Option Explicit
Public Const MaxValue As Long = 10
Private cache As Object
Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Public Type Customer
    Name As String
End Type
Public Enum Color
    Red = 1
End Enum
Public Function ParseJson(ByVal JsonString As String, Optional ByVal Strict As Boolean = False) As Object
Start:
10  Debug.Print JsonString
End Function
Private Sub Hidden()
End Sub
`
	if err := os.WriteFile(filepath.Join(moduleDir, "Main.bas"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}

	result, err := Inspect(Options{RootDir: dir, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if result.Summary.Files != 1 {
		t.Fatalf("files = %d, want 1", result.Summary.Files)
	}
	file := result.Files[0]
	if file.ModuleName != "Main" || file.ModuleKind != "standard" {
		t.Fatalf("unexpected module metadata: %+v", file)
	}
	assertSymbol(t, file.Symbols, "Main", "module")
	assertSymbol(t, file.Symbols, "MaxValue", "const")
	assertSymbol(t, file.Symbols, "GetTickCount", "declare_function")
	assertSymbol(t, file.Symbols, "Customer", "type")
	assertSymbol(t, file.Symbols, "Color", "enum")
	parseJson := assertSymbol(t, file.Symbols, "ParseJson", "function")
	if parseJson.ReturnType != "Object" {
		t.Fatalf("return type = %q, want Object", parseJson.ReturnType)
	}
	if len(parseJson.Parameters) != 2 || parseJson.Parameters[0].Name != "JsonString" || parseJson.Parameters[0].Passing != "ByVal" || !parseJson.Parameters[1].Optional {
		t.Fatalf("unexpected parameters: %+v", parseJson.Parameters)
	}
	assertNoSymbol(t, file.Symbols, "Hidden")
	assertNoSymbol(t, file.Symbols, "Start")

	withPrivate, err := Inspect(Options{RootDir: dir, Config: cfg, IncludePrivate: true, IncludeLabels: true})
	if err != nil {
		t.Fatal(err)
	}
	privateFile := withPrivate.Files[0]
	assertSymbol(t, privateFile.Symbols, "cache", "module_variable")
	assertSymbol(t, privateFile.Symbols, "Hidden", "sub")
	assertSymbol(t, privateFile.Symbols, "Start", "label")
	assertSymbol(t, privateFile.Symbols, "10", "line_number_label")
}

func TestInspectExtractsClassFieldsPropertiesAndImplements(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	classDir := filepath.Join(dir, "src", "classes")
	if err := os.MkdirAll(classDir, 0o755); err != nil {
		t.Fatal(err)
	}
	body := `VERSION 1.0 CLASS
Attribute VB_Name = "OrderService"
Option Explicit
Implements IService
Public WithEvents App As Excel.Application
Private mCount As Long
Public Property Get Count() As Long
End Property
`
	if err := os.WriteFile(filepath.Join(classDir, "OrderService.cls"), []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}

	result, err := Inspect(Options{RootDir: dir, Config: cfg, IncludePrivate: true})
	if err != nil {
		t.Fatal(err)
	}
	file := result.Files[0]
	if file.ModuleKind != "class" {
		t.Fatalf("module kind = %q, want class", file.ModuleKind)
	}
	assertSymbol(t, file.Symbols, "IService", "implements")
	assertSymbol(t, file.Symbols, "App", "withevents_field")
	assertSymbol(t, file.Symbols, "mCount", "field")
	assertSymbol(t, file.Symbols, "Count", "property_get")
}

func TestInspectSidecarFormCodeAvoidsDuplicateFrmSymbols(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	cfg.UserForm.CodeSource = "sidecar"
	formsDir := filepath.Join(dir, "src", "forms")
	if err := os.MkdirAll(filepath.Join(formsDir, "code"), 0o755); err != nil {
		t.Fatal(err)
	}
	frm := "VERSION 5.00\nBegin {GUID} UserForm1\nEnd\nAttribute VB_Name = \"UserForm1\"\n\nPrivate Sub StaleFrm()\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(formsDir, "UserForm1.frm"), []byte(frm), 0o644); err != nil {
		t.Fatal(err)
	}
	sidecar := "Option Explicit\nPublic Sub UserForm_Initialize()\nEnd Sub\n"
	if err := os.WriteFile(filepath.Join(formsDir, "code", "UserForm1.bas"), []byte(sidecar), 0o644); err != nil {
		t.Fatal(err)
	}

	result, err := Inspect(Options{RootDir: dir, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if result.Summary.Files != 1 {
		t.Fatalf("files = %d, want 1: %+v", result.Summary.Files, result.Files)
	}
	file := result.Files[0]
	if file.ModuleKind != "form" || file.Path != "src/forms/code/UserForm1.bas" {
		t.Fatalf("unexpected form sidecar file: %+v", file)
	}
	assertSymbol(t, file.Symbols, "UserForm_Initialize", "sub")
	assertNoSymbol(t, file.Symbols, "StaleFrm")
}

func TestInspectModuleFilter(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default()
	moduleDir := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(moduleDir, "A.bas"), []byte("Attribute VB_Name = \"A\"\nPublic Sub RunA()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(moduleDir, "B.bas"), []byte("Attribute VB_Name = \"B\"\nPublic Sub RunB()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}

	result, err := Inspect(Options{RootDir: dir, Config: cfg, Module: "b"})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Files) != 1 || result.Files[0].ModuleName != "B" {
		t.Fatalf("unexpected module filter result: %+v", result.Files)
	}
}

func assertSymbol(t *testing.T, symbols []Symbol, name, kind string) Symbol {
	t.Helper()
	for _, symbol := range symbols {
		if symbol.Name == name && symbol.Kind == kind {
			return symbol
		}
	}
	t.Fatalf("missing symbol %s/%s in %+v", name, kind, symbols)
	return Symbol{}
}

func assertNoSymbol(t *testing.T, symbols []Symbol, name string) {
	t.Helper()
	for _, symbol := range symbols {
		if symbol.Name == name {
			t.Fatalf("unexpected symbol %s in %+v", name, symbols)
		}
	}
}
