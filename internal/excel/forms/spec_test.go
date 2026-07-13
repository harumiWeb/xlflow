package forms

import (
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestResolveSnapshotOutputValidatesAndNormalizes(t *testing.T) {
	root := t.TempDir()
	resolved, err := ResolveSnapshotOutput(root, " artifacts\\UserForm1.form.yaml ")
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

	if _, err := ResolveSnapshotOutput(root, "artifacts\\UserForm1.form.txt"); err == nil || !strings.Contains(err.Error(), ".json, .yaml, or .yml") {
		t.Fatalf("expected extension validation error, got %v", err)
	}

	dirPath := filepath.Join(root, "artifacts")
	if err := os.MkdirAll(dirPath, 0o755); err != nil {
		t.Fatal(err)
	}
	if _, err := ResolveSnapshotOutput(root, dirPath); err == nil || !strings.Contains(err.Error(), ".json, .yaml, or .yml") {
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
		"warnings": []any{
			map[string]any{
				"code":    "unsupported_property",
				"message": "The snapshot omitted a designer-only property.",
			},
		},
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
	if spec.Form.Observed == nil || spec.Form.Build == nil {
		t.Fatalf("expected observed/build form values, got %#v", spec.Form)
	}
	if len(spec.Controls) != 2 {
		t.Fatalf("controls = %#v", spec.Controls)
	}
	control := spec.Controls[0]
	if control.ID == "" {
		t.Fatalf("expected generated control id: %#v", control)
	}
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
	if control.Observed == nil || control.Observed.Width == nil || *control.Observed.Width != 120.0 {
		t.Fatalf("observed control values = %#v", control.Observed)
	}
	child := spec.Controls[1]
	if child.Name != "lblNested" || child.ParentID != control.ID {
		t.Fatalf("child control = %#v, parent id = %q", child, control.ID)
	}
	if len(spec.Warnings) != 1 || spec.Warnings[0].Code != "unsupported_property" {
		t.Fatalf("warnings = %#v", spec.Warnings)
	}
}

func TestWriteSnapshotWritesJSONAndYAML(t *testing.T) {
	root := t.TempDir()
	spec := FormSpec{
		SchemaVersion:    1,
		Kind:             "xlflow.userform",
		Basis:            "designer",
		CoordinateSystem: "parent-relative",
		Form: FormSpecForm{
			Name: "UserForm1",
			Observed: &FormSpecObservedForm{
				Width:  ptrFloat(308),
				Height: ptrFloat(372),
			},
			Build: &FormSpecBuildForm{
				Width:  ptrFloat(308),
				Height: ptrFloat(372),
			},
		},
		Controls: []FormSpecControl{{
			ID:     "txtCustomer",
			Type:   "TextBox",
			Name:   "txtCustomer",
			ProgID: "Forms.TextBox.1",
			Observed: &FormSpecObservedControl{
				Width: ptrFloat(120),
			},
		}},
		Warnings: []FormSpecWarning{},
	}

	jsonOutput, err := ResolveSnapshotOutput(root, "artifacts\\UserForm1.form.json")
	if err != nil {
		t.Fatal(err)
	}
	if err := WriteSnapshot(jsonOutput, spec); err != nil {
		t.Fatal(err)
	}
	jsonBody, err := os.ReadFile(jsonOutput.Path)
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{`"schemaVersion": 1`, `"coordinateSystem": "parent-relative"`, `"progId": "Forms.TextBox.1"`, `"warnings": []`} {
		if !strings.Contains(string(jsonBody), want) {
			t.Fatalf("json snapshot missing %q:\n%s", want, string(jsonBody))
		}
	}
	var decoded map[string]any
	if err := json.Unmarshal(jsonBody, &decoded); err != nil {
		t.Fatalf("json snapshot should remain valid: %v\n%s", err, string(jsonBody))
	}

	yamlOutput, err := ResolveSnapshotOutput(root, "artifacts\\UserForm1.form.yaml")
	if err != nil {
		t.Fatal(err)
	}
	if err := WriteSnapshot(yamlOutput, spec); err != nil {
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

func TestFormSpecFromInspectSnapshotAssignsPlaceholderToUnnamedControl(t *testing.T) {
	spec, err := FormSpecFromInspectSnapshot(map[string]any{
		"name":  "UserForm1",
		"basis": "designer",
		"controls": []any{
			map[string]any{
				"name": "",
				"type": "Label",
				"controls": []any{
					map[string]any{
						"type": "TextBox",
					},
				},
			},
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(spec.Controls) != 2 || spec.Controls[0].Name != "<unnamed_1>" {
		t.Fatalf("unexpected top-level unnamed control placeholder: %#v", spec.Controls)
	}
	if spec.Controls[1].Name != "<unnamed_2>" || spec.Controls[1].ParentID != spec.Controls[0].ID {
		t.Fatalf("unexpected nested unnamed control placeholder: %#v", spec.Controls[1])
	}
	if len(spec.Warnings) != 2 {
		t.Fatalf("warnings = %#v", spec.Warnings)
	}
	if spec.Warnings[0].Code != "unnamed_control_placeholder" || spec.Warnings[1].Code != "unnamed_control_placeholder" {
		t.Fatalf("unexpected warnings = %#v", spec.Warnings)
	}
}

func TestLoadFormSpecValidatesSchemaAndControls(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "UserForm1.form.json")
	body := `{
  "schemaVersion": 1,
  "kind": "xlflow.userform",
  "basis": "designer",
  "form": { "name": "UserForm1" },
  "controls": [
    { "id": "txt_customer", "name": "txtCustomer", "type": "TextBox" }
  ],
  "warnings": []
}`
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	input, err := ResolveSpecInput(root, path)
	if err != nil {
		t.Fatal(err)
	}
	spec, err := LoadFormSpec(input)
	if err != nil {
		t.Fatal(err)
	}
	if spec.Form.Name != "UserForm1" {
		t.Fatalf("form name = %q", spec.Form.Name)
	}
	if len(spec.Controls) != 1 || spec.Controls[0].ID != "txt_customer" {
		t.Fatalf("expected control id, got %#v", spec.Controls)
	}

	if _, err := ResolveSpecInput(root, filepath.Join(root, "missing.form.json")); err == nil {
		t.Fatal("expected missing file error")
	}
}

func TestLoadFormSpecFlattensLegacyNestedControls(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "UserForm1.form.yaml")
	body := `
schemaVersion: 1
kind: xlflow.userform
basis: designer
form:
  name: UserForm1
controls:
  - id: frame_main
    name: Frame1
    type: Frame
    controls:
      - id: txt_customer
        name: txtCustomer
        type: TextBox
warnings: []
`
	if err := os.WriteFile(path, []byte(strings.TrimSpace(body)+"\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	input, err := ResolveSpecInput(root, path)
	if err != nil {
		t.Fatal(err)
	}
	spec, err := LoadFormSpec(input)
	if err != nil {
		t.Fatal(err)
	}
	if len(spec.Controls) != 2 {
		t.Fatalf("controls = %#v", spec.Controls)
	}
	if spec.Controls[1].ParentID != spec.Controls[0].ID {
		t.Fatalf("expected nested child parent id, got %#v", spec.Controls)
	}
}

func TestLoadFormSpecRejectsDuplicateExplicitControlIDs(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "UserForm1.form.json")
	body := `{
  "schemaVersion": 1,
  "kind": "xlflow.userform",
  "basis": "designer",
  "form": { "name": "UserForm1" },
  "controls": [
    { "id": "shared", "name": "Frame1", "type": "Frame" },
    { "id": "shared", "parentId": "shared", "name": "txtCustomer", "type": "TextBox" }
  ],
  "warnings": []
}`
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	input, err := ResolveSpecInput(root, path)
	if err != nil {
		t.Fatal(err)
	}
	_, err = LoadFormSpec(input)
	if err == nil || !strings.Contains(err.Error(), `id "shared" is duplicated`) {
		t.Fatalf("expected duplicate id validation error, got %v", err)
	}
	var specErr *SpecError
	if !errors.As(err, &specErr) {
		t.Fatalf("expected SpecError, got %T", err)
	}
	if specErr.Code != "spec_validation_failed" || specErr.Field != "controls[1].id" {
		t.Fatalf("unexpected spec error: %+v", specErr)
	}
}

func TestValidateFormSpecSourceReportsStrictStructuralIssues(t *testing.T) {
	body := []byte(`schemaVersion: 2
kind: xlflow.userform
basis: designer
extraRoot: true
form:
  name: UserForm1
  build:
    width: wide
  observed:
    insideWidth: 200
    extraObserved: true
controls:
  - id: label_status
    name: LabelStatus
    type: Label
    list:
      - A
      - B
    observed:
      missing: true
  - id: button_ok
    name: OKButton
    type: CommandButton
    selectedIndex: 1
warnings: []
`)
	issues, err := ValidateFormSpecSource(SpecInput{Format: "yaml", DisplayPath: "UserForm1.yaml"}, body)
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []struct {
		code  string
		field string
	}{
		{"UFV003", "schemaVersion"},
		{"UFV001", "extraRoot"},
		{"UFV002", "form.build.width"},
		{"UFV001", "form.observed.extraObserved"},
		{"UFV005", "controls[0].list"},
		{"UFV001", "controls[0].observed.missing"},
		{"UFV005", "controls[1].selectedIndex"},
	} {
		if !hasValidationIssue(issues, want.code, want.field) {
			t.Fatalf("missing validation issue %s at %s in %+v", want.code, want.field, issues)
		}
	}
}

func TestValidateFormSpecSourceReportsReferenceAndProgIDIssues(t *testing.T) {
	body := []byte(`{
  "schemaVersion": 1,
  "kind": "xlflow.userform",
  "basis": "designer",
  "form": { "name": "UserForm1" },
  "controls": [
    { "id": "frame_a", "name": "FrameA", "type": "Frame", "parentId": "frame_b" },
    { "id": "frame_b", "name": "FrameB", "type": "Frame", "parentId": "frame_a" },
    { "id": "txt_parent", "name": "TextBox1", "type": "TextBox" },
    { "id": "lbl_child", "name": "Label1", "type": "Label", "parentId": "txt_parent" },
    { "id": "self", "name": "SelfFrame", "type": "Frame", "parentId": "self" },
    { "id": "missing", "name": "MissingParent", "type": "Label", "parentId": "nope" },
    { "id": "bad_prog", "name": "BadProg", "type": "TextBox", "progId": "Forms.Label.1" }
  ],
  "warnings": []
}`)
	issues, err := ValidateFormSpecSource(SpecInput{Format: "json", DisplayPath: "UserForm1.json"}, body)
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []struct {
		code  string
		field string
	}{
		{"UFV010", "controls"},
		{"UFV011", "controls[3].parentId"},
		{"UFV009", "controls[4].parentId"},
		{"UFV008", "controls[5].parentId"},
		{"UFV012", "controls[6].progId"},
	} {
		if !hasValidationIssue(issues, want.code, want.field) {
			t.Fatalf("missing validation issue %s at %s in %+v", want.code, want.field, issues)
		}
	}
}

func TestValidateFormSpecSourceAcceptsCustomProgIDWithWarnings(t *testing.T) {
	body := []byte(`schemaVersion: 1
kind: xlflow.userform
basis: designer
form:
  name: UserForm1
  width: 240
controls:
  - id: custom_parent
    name: CustomParent
    type: VendorWidget
    progId: Vendor.Widget.1
    properties:
      customCaption: Details
  - id: label_child
    parentId: custom_parent
    name: Label1
    type: Label
    caption: Name
warnings:
  - code: captured
    message: captured warning
`)
	issues, err := ValidateFormSpecSource(SpecInput{Format: "yaml", DisplayPath: "UserForm1.yaml"}, body)
	if err != nil {
		t.Fatal(err)
	}
	if hasValidationErrors(issues) {
		t.Fatalf("custom ProgID should not produce errors: %+v", issues)
	}
	for _, want := range []struct {
		code    string
		field   string
		support SupportLevel
	}{
		{"UFV014", "controls[0].progId", SupportLevelCustomUnchecked},
		{"UFV013", "form.width", SupportLevelBestEffort},
		{"UFV013", "warnings", SupportLevelSnapshotOnly},
	} {
		if !hasValidationIssueWithSupport(issues, want.code, want.field, want.support) {
			t.Fatalf("missing warning %s at %s support %s in %+v", want.code, want.field, want.support, issues)
		}
	}
}

func TestLoadFormSpecReturnsMultipleValidationIssues(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "UserForm1.form.yaml")
	body := `schemaVersion: 1
kind: xlflow.userform
basis: designer
form:
  name: UserForm1
controls:
  - id: shared
    name: Label1
    type: Label
    list: [A]
  - id: shared
    name: TextBox1
    type: TextBox
    parentId: missing_parent
warnings: []
`
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	input, err := ResolveSpecInput(root, path)
	if err != nil {
		t.Fatal(err)
	}
	_, err = LoadFormSpec(input)
	var specErr *SpecError
	if !errors.As(err, &specErr) {
		t.Fatalf("expected SpecError, got %T", err)
	}
	if len(specErr.Issues) < 3 {
		t.Fatalf("expected multiple issues, got %+v", specErr.Issues)
	}
	if !hasValidationIssue(specErr.Issues, "UFV005", "controls[0].list") ||
		!hasValidationIssue(specErr.Issues, "UFV007", "controls[1].id") ||
		!hasValidationIssue(specErr.Issues, "UFV008", "controls[1].parentId") {
		t.Fatalf("unexpected issues: %+v", specErr.Issues)
	}
}

func TestLoadFormSpecReturnsParseMetadataAndSuggestion(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "UserForm1.form.yaml")
	body := `
schemaVersion: 1
kind: xlflow.userform
basis: designer
form:
  name: UserForm1
  caption: -
controls: []
warnings: []
`
	if err := os.WriteFile(path, []byte(strings.TrimSpace(body)+"\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	input, err := ResolveSpecInput(root, path)
	if err != nil {
		t.Fatal(err)
	}
	_, err = LoadFormSpec(input)
	var specErr *SpecError
	if !errors.As(err, &specErr) {
		t.Fatalf("expected SpecError, got %T", err)
	}
	if specErr.Code != "spec_parse_failed" {
		t.Fatalf("code = %q", specErr.Code)
	}
	if specErr.Path != "UserForm1.form.yaml" || specErr.Format != "yaml" {
		t.Fatalf("unexpected spec metadata: %+v", specErr)
	}
	if !strings.Contains(specErr.Suggestion, `caption: ""`) {
		t.Fatalf("suggestion = %q", specErr.Suggestion)
	}
}

func TestLoadFormSpecReturnsJSONSpecificParseSuggestion(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "UserForm1.form.json")
	body := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"UserForm1",},"controls":[]}`
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	input, err := ResolveSpecInput(root, path)
	if err != nil {
		t.Fatal(err)
	}
	_, err = LoadFormSpec(input)
	var specErr *SpecError
	if !errors.As(err, &specErr) {
		t.Fatalf("expected SpecError, got %T", err)
	}
	if specErr.Code != "spec_parse_failed" || specErr.Format != "json" {
		t.Fatalf("unexpected parse error: %+v", specErr)
	}
	if !strings.Contains(specErr.Suggestion, "Fix JSON syntax") {
		t.Fatalf("suggestion = %q", specErr.Suggestion)
	}
	if strings.Contains(specErr.Suggestion, "Try using JSON") {
		t.Fatalf("json suggestion should not tell the user to switch to JSON: %q", specErr.Suggestion)
	}
}

func ptrFloat(value float64) *float64 {
	return &value
}

func hasValidationIssue(issues []ValidationIssue, code, field string) bool {
	for _, issue := range issues {
		if issue.Code == code && issue.Field == field {
			return true
		}
	}
	return false
}

func hasValidationIssueWithSupport(issues []ValidationIssue, code, field string, support SupportLevel) bool {
	for _, issue := range issues {
		if issue.Code == code && issue.Field == field && issue.Support == support {
			return true
		}
	}
	return false
}
