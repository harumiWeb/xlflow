package testdiscover

import (
	"errors"
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestDiscoverSourceDefinedVBATests(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, filepath.Join("src", "modules", "SmokeTests.bas"), `Attribute VB_Name = "SmokeTests"
Option Explicit

'@Tag("smoke")
'' Plain documentation line
'@ExpectedError(5, "Invalid ""値""", "InvoiceParser")
'@Tag("fast")
Public Sub Test足し算()
End Sub

Public Sub EndsWith_Test()
End Sub

Private Sub TestPrivate()
End Sub

Public Function TestFunction() As Boolean
End Function

Public Sub TestWithArg(ByVal value As Long)
End Sub
`)
	writeModule(t, dir, filepath.Join("src", "classes", "Ignored.cls"), `VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Ignored"
Option Explicit
Public Sub TestClass()
End Sub
`)

	result, err := Discover(Options{RootDir: dir, Config: config.Default()})
	if err != nil {
		t.Fatalf("Discover() error = %v", err)
	}
	if result.Root != "src" || result.Summary.Files != 1 || result.Summary.Tests != 2 {
		t.Fatalf("unexpected summary: %+v", result)
	}
	if len(result.Items) != 2 {
		t.Fatalf("items = %d, want 2: %+v", len(result.Items), result.Items)
	}
	first := result.Items[0]
	if first.ID != "SmokeTests.Test足し算" || first.Module != "SmokeTests" || first.Name != "Test足し算" || first.QualifiedName != "SmokeTests.Test足し算" {
		t.Fatalf("unexpected first test identity: %+v", first)
	}
	if first.SourcePath != "src/modules/SmokeTests.bas" || first.Line != 8 {
		t.Fatalf("unexpected first test location: %+v", first)
	}
	if !reflect.DeepEqual(first.Tags, []string{"fast", "smoke"}) {
		t.Fatalf("tags = %#v, want fast/smoke", first.Tags)
	}
	if first.ExpectedError == nil || first.ExpectedError.Number != 5 {
		t.Fatalf("expected error missing: %+v", first.ExpectedError)
	}
	if first.ExpectedError.Description == nil || *first.ExpectedError.Description != `Invalid "値"` {
		t.Fatalf("description = %+v", first.ExpectedError.Description)
	}
	if first.ExpectedError.Source == nil || *first.ExpectedError.Source != "InvoiceParser" {
		t.Fatalf("source = %+v", first.ExpectedError.Source)
	}
	second := result.Items[1]
	if second.Name != "EndsWith_Test" || len(second.Tags) != 0 {
		t.Fatalf("unexpected second test: %+v", second)
	}
	if second.ExpectedError != nil {
		t.Fatalf("unexpected expected error on second test: %+v", second.ExpectedError)
	}
}

func TestDiscoverAllowsSameProcedureNameInDifferentModules(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, filepath.Join("src", "modules", "InvoiceTests.bas"), `Attribute VB_Name = "InvoiceTests"
Option Explicit
Public Sub Test_Export()
End Sub
`)
	writeModule(t, dir, filepath.Join("src", "modules", "CustomerTests.bas"), `Attribute VB_Name = "CustomerTests"
Option Explicit
Public Sub Test_Export()
End Sub
`)

	result, err := Discover(Options{RootDir: dir, Config: config.Default()})
	if err != nil {
		t.Fatalf("Discover() error = %v", err)
	}
	if result.Summary.Tests != 2 {
		t.Fatalf("tests = %d, want 2: %+v", result.Summary.Tests, result.Items)
	}
	got := map[string]bool{}
	for _, item := range result.Items {
		got[item.ID] = true
		if item.ID != item.QualifiedName {
			t.Fatalf("id and qualified_name diverged: %+v", item)
		}
	}
	for _, want := range []string{"InvoiceTests.Test_Export", "CustomerTests.Test_Export"} {
		if !got[want] {
			t.Fatalf("missing qualified id %q in %+v", want, result.Items)
		}
	}
}

func TestDiscoverRejectsDuplicateProcedureNameInSameModule(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, filepath.Join("src", "modules", "InvoiceTests.bas"), `Attribute VB_Name = "InvoiceTests"
Option Explicit
Public Sub Test_Export()
End Sub
Public Sub test_export()
End Sub
`)

	_, err := Discover(Options{RootDir: dir, Config: config.Default()})
	if err == nil {
		t.Fatal("expected duplicate test procedure error")
	}
}

func TestDiscoverEmptySourceReturnsEmptyItems(t *testing.T) {
	dir := t.TempDir()
	if err := os.MkdirAll(filepath.Join(dir, "src", "modules"), 0o755); err != nil {
		t.Fatal(err)
	}
	result, err := Discover(Options{RootDir: dir, Config: config.Default()})
	if err != nil {
		t.Fatalf("Discover() error = %v", err)
	}
	if result.Summary.Files != 0 || result.Summary.Tests != 0 || len(result.Items) != 0 {
		t.Fatalf("unexpected empty result: %+v", result)
	}
}

func TestDiscoverExpectedErrorNumberOnly(t *testing.T) {
	dir := t.TempDir()
	writeModule(t, dir, filepath.Join("src", "modules", "ParserTests.bas"), `Attribute VB_Name = "ParserTests"
Option Explicit

'@ExpectedError(5)
Public Sub Test_InvalidArgument()
End Sub
`)

	result, err := Discover(Options{RootDir: dir, Config: config.Default()})
	if err != nil {
		t.Fatalf("Discover() error = %v", err)
	}
	if len(result.Items) != 1 || result.Items[0].ExpectedError == nil {
		t.Fatalf("expected metadata: %+v", result.Items)
	}
	if got := result.Items[0].ExpectedError.Number; got != 5 {
		t.Fatalf("number = %d, want 5", got)
	}
}

func TestDiscoverRejectsInvalidExpectedErrorMetadata(t *testing.T) {
	cases := []struct {
		name string
		body string
	}{
		{
			name: "duplicate",
			body: `'@ExpectedError(5)
'@ExpectedError(6)
Public Sub Test_InvalidArgument()
End Sub
`,
		},
		{
			name: "non_numeric",
			body: `'@ExpectedError(foo)
Public Sub Test_InvalidArgument()
End Sub
`,
		},
		{
			name: "bad_arg_count",
			body: `'@ExpectedError(5, "a", "b", "c")
Public Sub Test_InvalidArgument()
End Sub
`,
		},
		{
			name: "malformed_string",
			body: `'@ExpectedError(5, "unterminated)
Public Sub Test_InvalidArgument()
End Sub
`,
		},
		{
			name: "non_test",
			body: `'@ExpectedError(5)
Public Sub Helper()
End Sub
`,
		},
	}

	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			dir := t.TempDir()
			writeModule(t, dir, filepath.Join("src", "modules", "ParserTests.bas"), "Attribute VB_Name = \"ParserTests\"\nOption Explicit\n\n"+c.body)

			_, err := Discover(Options{RootDir: dir, Config: config.Default()})
			if err == nil {
				t.Fatal("Discover() error = nil, want invalid metadata")
			}
			var metadataErr InvalidMetadataError
			if !errors.As(err, &metadataErr) {
				t.Fatalf("error = %T %v, want InvalidMetadataError", err, err)
			}
			if metadataErr.Path == "" || metadataErr.Line == 0 || metadataErr.Module != "ParserTests" {
				t.Fatalf("metadata location missing: %+v", metadataErr)
			}
		})
	}
}

func writeModule(t *testing.T, dir, rel, body string) {
	t.Helper()
	path := filepath.Join(dir, rel)
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
}
