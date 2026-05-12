package excel

import (
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestResolveFormSnapshotOutputValidatesAndNormalizes(t *testing.T) {
	root := t.TempDir()
	resolved, err := ResolveFormSnapshotOutput(root, " artifacts\\UserForm1.form.yaml ")
	if err != nil {
		t.Fatal(err)
	}
	if resolved.Format != "yaml" {
		t.Fatalf("format = %q, want yaml", resolved.Format)
	}
	if resolved.DisplayPath != "artifacts/UserForm1.form.yaml" {
		t.Fatalf("display path = %q", resolved.DisplayPath)
	}
	if resolved.Path != filepath.Join(root, "artifacts", "UserForm1.form.yaml") {
		t.Fatalf("path = %q", resolved.Path)
	}

	if _, err := ResolveFormSnapshotOutput(root, "artifacts\\UserForm1.form.txt"); err == nil || !strings.Contains(err.Error(), ".json, .yaml, or .yml") {
		t.Fatalf("expected extension validation error, got %v", err)
	}

	dirPath := filepath.Join(root, "artifacts")
	if err := os.MkdirAll(dirPath, 0o755); err != nil {
		t.Fatal(err)
	}
	if _, err := ResolveFormSnapshotOutput(root, dirPath); err == nil || !strings.Contains(err.Error(), ".json, .yaml, or .yml") {
		t.Fatalf("expected directory extension validation error, got %v", err)
	}
}

func TestFormSpecFromInspectSnapshotConvertsDesignerPayload(t *testing.T) {
	spec, err := FormSpecFromInspectSnapshot(map[string]any{
		"name":              "UserForm1",
		"basis":             "designer",
		"caption":           "Order Entry",
		"width":             308.0,
		"height":            372.0,
		"coordinate_system": "parent-relative",
		"controls": []any{
			map[string]any{
				"name":           "txtCustomer",
				"type":           "TextBox",
				"prog_id":        "Forms.TextBox.1",
				"left":           24.0,
				"top":            36.0,
				"width":          120.0,
				"height":         18.0,
				"tab_index":      0.0,
				"enabled":        true,
				"visible":        true,
				"selected_index": -1.0,
				"list":           []any{"Alpha", "Beta"},
				"controls": []any{
					map[string]any{
						"name": "lblNested",
						"type": "Label",
					},
				},
			},
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	if spec.SchemaVersion != 1 || spec.Kind != "xlflow.userform" || spec.Basis != "designer" {
		t.Fatalf("unexpected top-level spec: %#v", spec)
	}
	if spec.CoordinateSystem != "parent-relative" {
		t.Fatalf("coordinate system = %q", spec.CoordinateSystem)
	}
	if spec.Form.Name != "UserForm1" || spec.Form.Caption == nil || *spec.Form.Caption != "Order Entry" {
		t.Fatalf("form summary = %#v", spec.Form)
	}
	if len(spec.Controls) != 1 {
		t.Fatalf("controls = %#v", spec.Controls)
	}
	control := spec.Controls[0]
	if control.ProgID != "Forms.TextBox.1" {
		t.Fatalf("progId = %q", control.ProgID)
	}
	if control.TabIndex == nil || *control.TabIndex != 0 {
		t.Fatalf("tabIndex = %#v", control.TabIndex)
	}
	if control.SelectedIndex == nil || *control.SelectedIndex != -1 {
		t.Fatalf("selectedIndex = %#v", control.SelectedIndex)
	}
	if len(control.List) != 2 || control.List[0] != "Alpha" {
		t.Fatalf("list = %#v", control.List)
	}
	if len(control.Controls) != 1 || control.Controls[0].Name != "lblNested" {
		t.Fatalf("nested controls = %#v", control.Controls)
	}
	if spec.Warnings == nil || len(spec.Warnings) != 0 {
		t.Fatalf("warnings = %#v, want empty slice", spec.Warnings)
	}
}

func TestWriteFormSnapshotWritesJSONAndYAML(t *testing.T) {
	root := t.TempDir()
	spec := FormSpec{
		SchemaVersion:    1,
		Kind:             "xlflow.userform",
		Basis:            "designer",
		CoordinateSystem: "parent-relative",
		Form:             FormSpecForm{Name: "UserForm1"},
		Controls: []FormSpecControl{{
			Type:   "TextBox",
			Name:   "txtCustomer",
			ProgID: "Forms.TextBox.1",
		}},
		Warnings: []FormSpecWarning{},
	}

	jsonOutput, err := ResolveFormSnapshotOutput(root, "artifacts\\UserForm1.form.json")
	if err != nil {
		t.Fatal(err)
	}
	if err := WriteFormSnapshot(jsonOutput, spec); err != nil {
		t.Fatal(err)
	}
	jsonBody, err := os.ReadFile(jsonOutput.Path)
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"\"schemaVersion\": 1", "\"coordinateSystem\": \"parent-relative\"", "\"progId\": \"Forms.TextBox.1\"", "\"warnings\": []"} {
		if !strings.Contains(string(jsonBody), want) {
			t.Fatalf("json snapshot missing %q:\n%s", want, string(jsonBody))
		}
	}
	var decoded map[string]any
	if err := json.Unmarshal(jsonBody, &decoded); err != nil {
		t.Fatalf("json snapshot should remain valid: %v\n%s", err, string(jsonBody))
	}

	yamlOutput, err := ResolveFormSnapshotOutput(root, "artifacts\\UserForm1.form.yaml")
	if err != nil {
		t.Fatal(err)
	}
	if err := WriteFormSnapshot(yamlOutput, spec); err != nil {
		t.Fatal(err)
	}
	yamlBody, err := os.ReadFile(yamlOutput.Path)
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"schemaVersion: 1", "coordinateSystem: parent-relative", "progId: Forms.TextBox.1", "warnings: []"} {
		if !strings.Contains(string(yamlBody), want) {
			t.Fatalf("yaml snapshot missing %q:\n%s", want, string(yamlBody))
		}
	}
}
