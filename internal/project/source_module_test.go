package project

import (
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"gopkg.in/yaml.v3"
)

func TestRemoveModuleRemovesStandardClassAndFormArtifacts(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default().Src
	writeProjectFile(t, dir, "src/modules/InvoiceAggregator.bas", `Attribute VB_Name = "InvoiceAggregator"`+"\n")
	writeProjectFile(t, dir, "src/classes/InvoiceService.cls", `Attribute VB_Name = "InvoiceService"`+"\n")
	writeProjectFile(t, dir, "src/forms/CustomerForm.frm", `Begin VB.Form CustomerForm
Attribute VB_Name = "CustomerForm"
`)
	writeProjectFile(t, dir, "src/forms/CustomerForm.frx", "binary")
	writeProjectFile(t, dir, "src/forms/code/CustomerForm.bas", "Option Explicit\n")
	writeProjectFile(t, dir, "src/forms/specs/CustomerForm.yaml", "form:\n  name: CustomerForm\n")

	standard, err := RemoveModule(dir, "InvoiceAggregator", cfg)
	if err != nil {
		t.Fatalf("RemoveModule standard error = %v", err)
	}
	if standard.Kind != "standard" || !containsString(standard.Removed, "src/modules/InvoiceAggregator.bas") || !standard.RequiresPush {
		t.Fatalf("unexpected standard result: %+v", standard)
	}
	if pathExists(filepath.Join(dir, "src", "modules", "InvoiceAggregator.bas")) {
		t.Fatal("standard module should be removed")
	}

	class, err := RemoveModule(dir, "InvoiceService", cfg)
	if err != nil {
		t.Fatalf("RemoveModule class error = %v", err)
	}
	if class.Kind != "class" || !containsString(class.Removed, "src/classes/InvoiceService.cls") {
		t.Fatalf("unexpected class result: %+v", class)
	}

	form, err := RemoveModule(dir, "CustomerForm", cfg)
	if err != nil {
		t.Fatalf("RemoveModule form error = %v", err)
	}
	for _, want := range []string{
		"src/forms/CustomerForm.frm",
		"src/forms/CustomerForm.frx",
		"src/forms/code/CustomerForm.bas",
		"src/forms/specs/CustomerForm.yaml",
	} {
		if !containsString(form.Removed, want) {
			t.Fatalf("expected removed %s in %+v", want, form)
		}
		if pathExists(filepath.Join(dir, filepath.FromSlash(want))) {
			t.Fatalf("%s should be removed", want)
		}
	}
}

func TestRenameModuleUpdatesStandardAndClassVBName(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default().Src
	writeProjectFile(t, dir, "src/modules/OldStandard.bas", `Attribute VB_Name = "OldStandard"
Option Explicit
`)
	writeProjectFile(t, dir, "src/classes/OldClass.cls", `VERSION 1.0 CLASS
BEGIN
END
Attribute VB_Name = "OldClass"
Option Explicit
`)

	standard, err := RenameModule(dir, "OldStandard", "NewStandard", cfg)
	if err != nil {
		t.Fatalf("RenameModule standard error = %v", err)
	}
	if standard.Kind != "standard" || standard.OldName != "OldStandard" || standard.NewName != "NewStandard" || !standard.RequiresPush {
		t.Fatalf("unexpected standard result: %+v", standard)
	}
	assertFileContains(t, filepath.Join(dir, "src", "modules", "NewStandard.bas"), `Attribute VB_Name = "NewStandard"`)
	if pathExists(filepath.Join(dir, "src", "modules", "OldStandard.bas")) {
		t.Fatal("old standard path should be gone")
	}

	class, err := RenameModule(dir, "OldClass", "NewClass", cfg)
	if err != nil {
		t.Fatalf("RenameModule class error = %v", err)
	}
	if class.Kind != "class" {
		t.Fatalf("unexpected class result: %+v", class)
	}
	assertFileContains(t, filepath.Join(dir, "src", "classes", "NewClass.cls"), `Attribute VB_Name = "NewClass"`)
}

func TestRenameModuleUpdatesFormArtifactsAndSpecName(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default().Src
	writeProjectFile(t, dir, "src/forms/CustomerForm.frm", `VERSION 5.00
Begin VB.Form CustomerForm
   Caption = "Keep Caption"
End
Attribute VB_Name = "CustomerForm"
Option Explicit
`)
	writeProjectFile(t, dir, "src/forms/CustomerForm.frx", "binary")
	writeProjectFile(t, dir, "src/forms/code/CustomerForm.bas", "Option Explicit\nPrivate Sub UserForm_Initialize()\nEnd Sub\n")
	writeProjectFile(t, dir, "src/forms/specs/CustomerForm.yaml", `schemaVersion: 1
kind: xlflow.userform
form:
  name: CustomerForm
  caption: Keep Caption
controls:
  - name: txtName
    type: TextBox
`)

	result, err := RenameModule(dir, "CustomerForm", "RegistrationForm", cfg)
	if err != nil {
		t.Fatalf("RenameModule form error = %v", err)
	}
	if result.Kind != "form" || len(result.Renamed) != 4 {
		t.Fatalf("unexpected form result: %+v", result)
	}
	frmPath := filepath.Join(dir, "src", "forms", "RegistrationForm.frm")
	assertFileContains(t, frmPath, "Begin VB.Form RegistrationForm")
	assertFileContains(t, frmPath, `Attribute VB_Name = "RegistrationForm"`)
	assertFileContains(t, filepath.Join(dir, "src", "forms", "code", "RegistrationForm.bas"), "UserForm_Initialize")
	if !pathExists(filepath.Join(dir, "src", "forms", "RegistrationForm.frx")) {
		t.Fatal("renamed .frx should exist")
	}
	specBody, err := os.ReadFile(filepath.Join(dir, "src", "forms", "specs", "RegistrationForm.yaml"))
	if err != nil {
		t.Fatal(err)
	}
	var spec map[string]any
	if err := yaml.Unmarshal(specBody, &spec); err != nil {
		t.Fatal(err)
	}
	form := spec["form"].(map[string]any)
	if form["name"] != "RegistrationForm" || form["caption"] != "Keep Caption" {
		t.Fatalf("unexpected spec form: %+v", form)
	}
}

func TestRenameModuleUpdatesJSONFormSpecName(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default().Src
	writeProjectFile(t, dir, "src/forms/UserForm1.frm", `Begin VB.Form UserForm1
Attribute VB_Name = "UserForm1"
`)
	writeProjectFile(t, dir, "src/forms/specs/UserForm1.json", `{"form":{"name":"UserForm1","caption":"Keep"},"controls":[{"name":"txtName"}]}`)

	if _, err := RenameModule(dir, "UserForm1", "OrderForm", cfg); err != nil {
		t.Fatalf("RenameModule error = %v", err)
	}
	body, err := os.ReadFile(filepath.Join(dir, "src", "forms", "specs", "OrderForm.json"))
	if err != nil {
		t.Fatal(err)
	}
	var spec map[string]any
	if err := json.Unmarshal(body, &spec); err != nil {
		t.Fatal(err)
	}
	form := spec["form"].(map[string]any)
	if form["name"] != "OrderForm" || form["caption"] != "Keep" {
		t.Fatalf("unexpected JSON spec form: %+v", form)
	}
}

func TestModuleMutationsRejectProtectedDocumentModules(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default().Src
	writeProjectFile(t, dir, "src/workbook/ThisWorkbook.bas", "Option Explicit\n")

	if _, err := RemoveModule(dir, "ThisWorkbook", cfg); !errors.Is(err, ErrProtectedModule) {
		t.Fatalf("RemoveModule protected error = %v, want ErrProtectedModule", err)
	}
	if _, err := RenameModule(dir, "ThisWorkbook", "BookHost", cfg); !errors.Is(err, ErrProtectedModule) {
		t.Fatalf("RenameModule protected error = %v, want ErrProtectedModule", err)
	}
}

func TestModuleMutationsRejectInvalidMissingDuplicateAndAmbiguousNames(t *testing.T) {
	dir := t.TempDir()
	cfg := config.Default().Src
	writeProjectFile(t, dir, "src/modules/Existing.bas", `Attribute VB_Name = "Existing"`+"\n")
	writeProjectFile(t, dir, "src/classes/Collision.cls", `Attribute VB_Name = "Collision"`+"\n")
	writeProjectFile(t, dir, "src/modules/Ambiguous.bas", `Attribute VB_Name = "Ambiguous"`+"\n")
	writeProjectFile(t, dir, "src/classes/Ambiguous.cls", `Attribute VB_Name = "Ambiguous"`+"\n")

	if _, err := RemoveModule(dir, "Missing", cfg); !errors.Is(err, ErrModuleNotFound) {
		t.Fatalf("RemoveModule missing error = %v, want ErrModuleNotFound", err)
	}
	if _, err := RenameModule(dir, "Existing", "123Bad", cfg); !errors.Is(err, ErrInvalidComponentName) {
		t.Fatalf("RenameModule invalid error = %v, want ErrInvalidComponentName", err)
	}
	if _, err := RenameModule(dir, "Existing", "Collision", cfg); !errors.Is(err, ErrModuleAlreadyExists) {
		t.Fatalf("RenameModule collision error = %v, want ErrModuleAlreadyExists", err)
	}
	if _, err := RemoveModule(dir, "Ambiguous", cfg); !errors.Is(err, ErrModuleAmbiguous) {
		t.Fatalf("RemoveModule ambiguous error = %v, want ErrModuleAmbiguous", err)
	}
}

func writeProjectFile(t *testing.T, root string, relPath string, body string) {
	t.Helper()
	path := filepath.Join(root, filepath.FromSlash(relPath))
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
}

func assertFileContains(t *testing.T, path string, want string) {
	t.Helper()
	body, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(body), want) {
		t.Fatalf("%s missing %q:\n%s", path, want, string(body))
	}
}

func pathExists(path string) bool {
	_, err := os.Stat(path)
	return err == nil
}
