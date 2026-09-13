package intel

import (
	"context"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestByRefArgumentDiagnosticsAllowLocalUserDefinedArrayToPointerArray(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Private Type RawStorage
    value As Long
End Type

Private Sub WritePointerWords(ByRef words() As LONG_PTR, ByVal value As LongPtr)
    words(0) = value
End Sub

Private Sub WriteText(ByRef values() As String)
    values(0) = "changed"
End Sub

Public Sub Run()
    Dim raw(0) As RawStorage
    WritePointerWords raw, 0
    WriteText raw
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}

	diagnostics := diagnosticsByCode(analyzer.ByRefArgumentDiagnostics(doc), "VBA228")
	if len(diagnostics) != 1 {
		t.Fatalf("VBA228 diagnostics = %+v, want only the non-pointer array mismatch", diagnostics)
	}
	if !strings.Contains(diagnostics[0].Message, "requires String()") {
		t.Fatalf("unexpected non-pointer array mismatch diagnostic: %+v", diagnostics[0])
	}
}

func TestByRefArgumentDiagnosticsUseProjectUserDefinedArrayToPointerArray(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	analyzer.WorkspaceSymbolQueryFunc = func(_ []Document, query WorkspaceSymbolQuery) ([]Symbol, error) {
		if query.Mode != WorkspaceSymbolQueryKind {
			return nil, nil
		}
		return []Symbol{{Name: "SharedStorage", Kind: "type", Module: "Types"}}, nil
	}
	source := `Option Explicit
Private Sub WritePointerWords(ByRef words() As LONG_PTR)
End Sub

Public Sub Run()
    Dim raw(0) As SharedStorage
    WritePointerWords raw
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}

	if diagnostics := diagnosticsByCode(analyzer.ByRefArgumentDiagnostics(doc), "VBA228"); len(diagnostics) != 0 {
		t.Fatalf("project user-defined array should be accepted for pointer reinterpretation: %+v", diagnostics)
	}
}

func TestByRefArgumentDiagnosticsUseLegacyWorkspaceSymbolsForUDT(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	analyzer.WorkspaceSymbolsFunc = func(_ []Document, query string) ([]Symbol, error) {
		if query != "" {
			t.Fatalf("legacy workspace query = %q, want all-symbol query", query)
		}
		return []Symbol{{Name: "SharedStorage", Kind: "type", Module: "Types", Visibility: "Public"}}, nil
	}
	source := `Option Explicit
Private Sub WritePointerWords(ByRef words() As LongPtr)
End Sub

Public Sub Run()
    Dim raw(0) As SharedStorage
    WritePointerWords raw
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}

	if diagnostics := diagnosticsByCode(analyzer.ByRefArgumentDiagnostics(doc), "VBA228"); len(diagnostics) != 0 {
		t.Fatalf("legacy workspace UDT should be accepted for pointer reinterpretation: %+v", diagnostics)
	}
}

func TestByRefArgumentDiagnosticsUseBuiltInWorkspaceSymbolsForUDT(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	typesPath := filepath.Join(moduleDir, "Types.bas")
	if err := os.WriteFile(typesPath, []byte(`Attribute VB_Name = "Types"
Option Explicit
Public Type SharedStorage
    value As Long
End Type
`), 0o644); err != nil {
		t.Fatal(err)
	}
	analyzer := newTestAnalyzer(t)
	analyzer.RootDir = root
	doc := Document{
		Path:       filepath.Join(moduleDir, "Main.bas"),
		ModuleKind: "standard",
		Source: `Attribute VB_Name = "Main"
Option Explicit
Private Sub WritePointerWords(ByRef words() As LongPtr)
End Sub

Public Sub Run()
    Dim raw(0) As SharedStorage
    WritePointerWords raw
End Sub
`,
	}

	if diagnostics := diagnosticsByCode(analyzer.ByRefArgumentDiagnostics(doc), "VBA228"); len(diagnostics) != 0 {
		t.Fatalf("built-in workspace UDT should be accepted for pointer reinterpretation: %+v", diagnostics)
	}
}

func TestWorkspaceSymbolsQueryContextHonorsCancellationForBuiltInFallback(t *testing.T) {
	analyzer := Analyzer{RootDir: t.TempDir()}
	ctx, cancel := context.WithCancel(t.Context())
	cancel()

	_, err := analyzer.WorkspaceSymbolsQueryContext(ctx, nil, WorkspaceSymbolQuery{Mode: WorkspaceSymbolQueryKind})
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("built-in workspace query error = %v, want context.Canceled", err)
	}
}

func TestWorkspaceSymbolsQueryContextTreatsParserRecoveryAsIncomplete(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(moduleDir, "Broken.bas"), []byte(`Attribute VB_Name = "Broken"
Option Explicit
Public Function Broken(ByVal x As String
End Function
`), 0o644); err != nil {
		t.Fatal(err)
	}

	analyzer := newTestAnalyzer(t)
	analyzer.RootDir = root
	_, err := analyzer.WorkspaceSymbolsQueryContext(t.Context(), nil, WorkspaceSymbolQuery{Mode: WorkspaceSymbolQueryKind})
	if !errors.Is(err, errWorkspaceSymbolsIncomplete) {
		t.Fatalf("workspace query error = %v, want parser-recovery incompleteness", err)
	}
}

func TestWorkspaceSymbolsQueryContextTreatsOpenDocumentSymbolErrorAsIncomplete(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	analyzer.DocumentSymbolsFunc = func(Document, DocumentSymbolLoader) ([]Symbol, error) {
		return nil, errors.New("open document symbols unavailable")
	}
	doc := Document{
		Path:   filepath.Join(t.TempDir(), "Main.bas"),
		Source: "Option Explicit\nPublic Sub Run()\nEnd Sub\n",
	}

	_, err := analyzer.WorkspaceSymbolsQueryContext(t.Context(), []Document{doc}, WorkspaceSymbolQuery{Mode: WorkspaceSymbolQueryKind})
	if !errors.Is(err, errWorkspaceSymbolsIncomplete) {
		t.Fatalf("workspace query error = %v, want open-document incompleteness", err)
	}
}

func TestByRefArrayReinterpretationKeepsQualifiedTypeIdentity(t *testing.T) {
	localTypes := map[string]struct{}{"rawstorage": {}, "main.rawstorage": {}}
	param := Parameter{Type: "LongPtr", IsArray: true}

	if !byRefArrayReinterpretation(inferredType{Type: "Main.RawStorage", IsArray: true}, param, localTypes) {
		t.Fatal("qualified local user-defined type should be recognized")
	}
	if byRefArrayReinterpretation(inferredType{Type: "Other.RawStorage", IsArray: true}, param, localTypes) {
		t.Fatal("a same-named type from another module must not be treated as local")
	}
	if byRefArrayReinterpretation(inferredType{Type: "Excel.RawStorage", IsArray: true}, param, localTypes) {
		t.Fatal("an external qualified type must not be treated as local")
	}
}

func TestByRefPointerSizedTypeUsesExactAliases(t *testing.T) {
	for _, typ := range []string{"LongPtr", "LONG_PTR", "Long_Ptr"} {
		if !isByRefPointerSizedType(typ) {
			t.Errorf("isByRefPointerSizedType(%q) = false, want true", typ)
		}
	}
	for _, typ := range []string{"Long__Ptr", "LongPtrValue", "Pointer"} {
		if isByRefPointerSizedType(typ) {
			t.Errorf("isByRefPointerSizedType(%q) = true, want false", typ)
		}
	}
}

func TestWorkspaceUserDefinedTypeIndexFiltersPrivateAndConditionalTypes(t *testing.T) {
	index := NewWorkspaceUserDefinedTypeIndex([]Symbol{
		{Name: "PublicStorage", Kind: "type", Module: "Types", Visibility: "Public"},
		{Name: "PrivateStorage", Kind: "type", Module: "Types", Visibility: "Private"},
		{Name: "ConditionalStorage", Kind: "type", Module: "Types", Visibility: "Public", ConditionalBranches: []procedureir.ConditionalBranch{{Group: "platform", Branch: 0}}},
	})
	param := Parameter{Type: "LongPtr", IsArray: true}

	if !byRefArrayReinterpretationWithWorkspace(inferredType{Type: "PublicStorage", IsArray: true}, param, nil, index) {
		t.Fatal("public unconditional type should be recognized")
	}
	if byRefArrayReinterpretationWithWorkspace(inferredType{Type: "PrivateStorage", IsArray: true}, param, nil, index) {
		t.Fatal("private type from another module must not be recognized")
	}
	if byRefArrayReinterpretationWithWorkspace(inferredType{Type: "ConditionalStorage", IsArray: true}, param, nil, index) {
		t.Fatal("conditional type must not be recognized")
	}
}

func TestWorkspaceUserDefinedTypeIndexAcceptsProjectQualifiedTypes(t *testing.T) {
	index := NewWorkspaceUserDefinedTypeIndexForProject([]Symbol{
		{Name: "PointerAccessor", Kind: "type", Module: "Types", Visibility: "Public"},
	}, "ProjectName")
	param := Parameter{Type: "LongPtr", IsArray: true}

	if !byRefArrayReinterpretationWithWorkspace(inferredType{Type: "ProjectName.PointerAccessor", IsArray: true}, param, nil, index) {
		t.Fatal("project-qualified public type should be recognized")
	}
	if byRefArrayReinterpretationWithWorkspace(inferredType{Type: "OtherProject.PointerAccessor", IsArray: true}, param, nil, index) {
		t.Fatal("a type from another project must not be treated as local")
	}
}

func TestByRefArgumentDiagnosticsHoldUnknownArrayMismatchForIncompleteWorkspace(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	analyzer.WorkspaceUserDefinedTypes = NewWorkspaceUserDefinedTypeIndex(nil)
	analyzer.WorkspaceUserDefinedTypesComplete = false
	source := `Option Explicit
Private Sub WritePointerWords(ByRef words() As LongPtr)
End Sub

Public Sub Run()
    Dim raw(0) As SharedStorage
    WritePointerWords raw
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}

	if diagnostics := diagnosticsByCode(analyzer.ByRefArgumentDiagnostics(doc), "VBA228"); len(diagnostics) != 0 {
		t.Fatalf("incomplete workspace should not prove an unknown array mismatch: %+v", diagnostics)
	}

	analyzer.WorkspaceUserDefinedTypesComplete = true
	if diagnostics := diagnosticsByCode(analyzer.ByRefArgumentDiagnostics(doc), "VBA228"); len(diagnostics) != 1 {
		t.Fatalf("complete workspace should report the unknown non-UDT mismatch: %+v", diagnostics)
	}
}

func TestByRefLocalUserDefinedTypesRespectModuleVisibilityAndConditions(t *testing.T) {
	analyzer := Analyzer{}
	doc := Document{Path: "Main.bas", Source: "Attribute VB_Name = \"Main\"\n"}
	symbols := []Symbol{
		{Name: "CurrentPrivate", Kind: "type", Module: "Main", Visibility: "Private", File: doc.Path},
		{Name: "OtherPrivate", Kind: "type", Module: "Other", Visibility: "Private", File: "Other.bas"},
		{Name: "CurrentConditional", Kind: "type", Module: "Main", ConditionalBranches: []procedureir.ConditionalBranch{{Group: "platform", Branch: 0}}, File: doc.Path},
	}
	types := analyzer.byRefLocalUserDefinedTypesForDocument(doc, symbols)
	if _, ok := types["currentprivate"]; !ok {
		t.Fatal("current-module private type should be available")
	}
	if _, ok := types["otherprivate"]; ok {
		t.Fatal("other-module private type must not be available")
	}
	if _, ok := types["currentconditional"]; ok {
		t.Fatal("conditional type must not be available")
	}
}

func TestByRefArgumentDiagnosticsTreatWorkspaceQueryErrorAsIncomplete(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	analyzer.WorkspaceSymbolQueryFunc = func(_ []Document, query WorkspaceSymbolQuery) ([]Symbol, error) {
		if query.Mode == WorkspaceSymbolQueryKind {
			return nil, errors.New("workspace index is still building")
		}
		return nil, nil
	}
	source := `Option Explicit
Private Sub WritePointerWords(ByRef words() As LongPtr)
End Sub

Public Sub Run()
    Dim raw(0) As SharedStorage
    WritePointerWords raw
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}

	if diagnostics := diagnosticsByCode(analyzer.ByRefArgumentDiagnostics(doc), "VBA228"); len(diagnostics) != 0 {
		t.Fatalf("workspace query failure should hold the mismatch: %+v", diagnostics)
	}
}
