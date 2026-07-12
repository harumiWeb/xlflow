package testdiscover

import (
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
	if first.SourcePath != "src/modules/SmokeTests.bas" || first.Line != 7 {
		t.Fatalf("unexpected first test location: %+v", first)
	}
	if !reflect.DeepEqual(first.Tags, []string{"fast", "smoke"}) {
		t.Fatalf("tags = %#v, want fast/smoke", first.Tags)
	}
	second := result.Items[1]
	if second.Name != "EndsWith_Test" || len(second.Tags) != 0 {
		t.Fatalf("unexpected second test: %+v", second)
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
