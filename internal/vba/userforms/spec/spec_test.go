package spec

import (
	"bytes"
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
	"slices"
	"strings"
	"testing"

	"gopkg.in/yaml.v3"
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

func TestFormSpecFromInspectSnapshotMapsPagesTabsAndObservedPageGeometry(t *testing.T) {
	spec, err := FormSpecFromInspectSnapshot(map[string]any{
		"name": "NavigationForm",
		"controls": []any{
			map[string]any{
				"name":           "Pages",
				"type":           "MultiPage",
				"selected_index": 1,
				"controls": []any{
					map[string]any{
						"name":             "Details",
						"type":             "Page",
						"caption":          "Details",
						"control_tip_text": "Show details",
						"tag":              "details-page",
						"accelerator":      "D",
						"enabled":          true,
						"visible":          true,
						"left":             3.0,
						"top":              4.0,
						"width":            220.0,
						"height":           120.0,
					},
				},
			},
			map[string]any{
				"name":           "Navigation",
				"type":           "TabStrip",
				"selected_index": 0,
				"tabs": []any{
					map[string]any{
						"name":             "Main",
						"caption":          "Main",
						"control_tip_text": "Main view",
						"tag":              "main-tab",
						"accelerator":      "M",
						"enabled":          true,
						"visible":          false,
					},
				},
			},
			map[string]any{
				"name": "EmptyNavigation",
				"type": "TabStrip",
				"tabs": []any{},
			},
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(spec.Controls) != 4 {
		t.Fatalf("controls = %d, want MultiPage, Page, and two TabStrips", len(spec.Controls))
	}
	page := spec.Controls[1]
	if page.Type != "Page" || page.ParentID != spec.Controls[0].ID || page.Caption == nil || *page.Caption != "Details" {
		t.Fatalf("Page mapping = %#v", page)
	}
	if page.Tag == nil || *page.Tag != "details-page" || page.ControlTipText == nil || *page.ControlTipText != "Show details" || page.Accelerator == nil || *page.Accelerator != "D" {
		t.Fatalf("Page typed properties = %#v", page)
	}
	if page.Left != nil || page.Top != nil || page.Width != nil || page.Height != nil || page.Observed == nil || page.Observed.Width == nil || *page.Observed.Width != 220 {
		t.Fatalf("Page geometry should remain observed-only: %#v", page)
	}
	if spec.Controls[0].SelectedIndex == nil || *spec.Controls[0].SelectedIndex != 1 {
		t.Fatalf("MultiPage selectedIndex = %#v", spec.Controls[0].SelectedIndex)
	}
	tabStrip := spec.Controls[2]
	if len(tabStrip.Tabs) != 1 || tabStrip.Tabs[0].Name != "Main" || tabStrip.Tabs[0].ControlTipText == nil || *tabStrip.Tabs[0].ControlTipText != "Main view" || tabStrip.Tabs[0].Visible == nil || *tabStrip.Tabs[0].Visible {
		t.Fatalf("TabStrip mapping = %#v", tabStrip)
	}
	if tabStrip.Observed == nil || len(tabStrip.Observed.Tabs) != 1 {
		t.Fatalf("observed TabStrip tabs = %#v", tabStrip.Observed)
	}
	if spec.Controls[3].Tabs == nil || len(spec.Controls[3].Tabs) != 0 {
		t.Fatalf("empty TabStrip tabs lost: %#v", spec.Controls[3].Tabs)
	}
	if spec.Controls[3].Observed == nil || spec.Controls[3].Observed.Tabs == nil || len(spec.Controls[3].Observed.Tabs) != 0 {
		t.Fatalf("empty observed TabStrip tabs lost: %#v", spec.Controls[3].Observed)
	}
}

func TestFormSpecTabSlicePreservesNilAndExplicitEmptyInJSONAndYAML(t *testing.T) {
	for _, format := range []string{"json", "yaml"} {
		t.Run(format, func(t *testing.T) {
			for _, test := range []struct {
				name string
				tabs []FormSpecTab
				want bool
			}{
				{name: "nil means unspecified"},
				{name: "empty means delete all", tabs: []FormSpecTab{}, want: true},
			} {
				t.Run(test.name, func(t *testing.T) {
					body, err := MarshalSnapshot(format, FormSpec{
						SchemaVersion: 1,
						Kind:          "xlflow.userform",
						Basis:         "designer",
						Form:          FormSpecForm{Name: "TabsForm"},
						Controls: []FormSpecControl{{
							ID: "tabs", Name: "Tabs", Type: "TabStrip", Tabs: test.tabs,
							Observed: &FormSpecObservedControl{Tabs: test.tabs},
						}},
					})
					if err != nil {
						t.Fatal(err)
					}
					var decoded FormSpec
					if format == "json" {
						err = json.Unmarshal(body, &decoded)
					} else {
						err = yaml.Unmarshal(body, &decoded)
					}
					if err != nil {
						t.Fatal(err)
					}
					if got := decoded.Controls[0].Tabs != nil; got != test.want {
						t.Fatalf("decoded tabs nonnil = %v, want %v; body:\n%s", got, test.want, body)
					}
					if got := decoded.Controls[0].Observed != nil && decoded.Controls[0].Observed.Tabs != nil; got != test.want {
						t.Fatalf("decoded observed tabs nonnil = %v, want %v; body:\n%s", got, test.want, body)
					}
					if test.want {
						tabsToken := []byte("tabs: []")
						if format == "json" {
							tabsToken = []byte(`"tabs": []`)
						}
						if !bytes.Contains(body, tabsToken) {
							t.Fatalf("snapshot does not preserve explicit empty tabs:\n%s", body)
						}
					} else {
						fieldToken := []byte("tabs:")
						if format == "json" {
							fieldToken = []byte(`"tabs":`)
						}
						if bytes.Contains(body, fieldToken) {
							t.Fatalf("snapshot should omit unspecified tabs:\n%s", body)
						}
					}
				})
			}
		})
	}
}

func TestValidateFormSpecSourceEnforcesPageAndTabStripContract(t *testing.T) {
	tests := []struct {
		name   string
		source string
		code   string
		field  string
	}{
		{
			name: "MultiPage only accepts Page children",
			source: "controls:\n" +
				"  - id: pages\n    name: Pages\n    type: MultiPage\n" +
				"  - id: label\n    name: Label\n    type: Label\n    parentId: pages\n",
			code: "UFV011", field: "controls[1].parentId",
		},
		{
			name: "Page requires MultiPage parent",
			source: "controls:\n" +
				"  - id: frame\n    name: Frame\n    type: Frame\n" +
				"  - id: page\n    name: Page\n    type: Page\n    parentId: frame\n",
			code: "UFV011", field: "controls[1].parentId",
		},
		{
			name:   "Page requires a parent",
			source: "controls:\n  - id: page\n    name: Page\n    type: Page\n",
			code:   "UFV011", field: "controls[0].parentId",
		},
		{
			name: "TabStrip cannot contain controls",
			source: "controls:\n" +
				"  - id: tabs\n    name: Tabs\n    type: TabStrip\n" +
				"  - id: label\n    name: Label\n    type: Label\n    parentId: tabs\n",
			code: "UFV011", field: "controls[1].parentId",
		},
		{
			name: "Tab names are case insensitive within their owner",
			source: "controls:\n  - id: tabs\n    name: Tabs\n    type: TabStrip\n    tabs:\n" +
				"      - name: Main\n      - name: mAiN\n",
			code: tabNameValidationCode, field: "controls[0].tabs[1].name",
		},
		{
			name:   "Tab names are required",
			source: "controls:\n  - id: tabs\n    name: Tabs\n    type: TabStrip\n    tabs:\n      - caption: Missing name\n",
			code:   "UFV004", field: "controls[0].tabs[0].name",
		},
		{
			name:   "Tab property values are typed",
			source: "controls:\n  - id: tabs\n    name: Tabs\n    type: TabStrip\n    tabs:\n      - name: Main\n        enabled: yes\n",
			code:   "UFV002", field: "controls[0].tabs[0].enabled",
		},
		{
			name: "Page geometry is not an authoring field",
			source: "controls:\n" +
				"  - id: pages\n    name: Pages\n    type: MultiPage\n" +
				"  - id: page\n    name: Page\n    type: Page\n    parentId: pages\n    width: 10\n",
			code: "UFV005", field: "controls[1].width",
		},
		{
			name:   "MultiPage accepts a partial Page list and selection",
			source: "controls:\n  - id: pages\n    name: Pages\n    type: MultiPage\n    selectedIndex: 8\n",
		},
		{
			name:   "TabStrip selected index is checked against explicit tabs",
			source: "controls:\n  - id: tabs\n    name: Tabs\n    type: TabStrip\n    selectedIndex: 0\n    tabs: []\n",
			code:   selectedIndexValidationCode, field: "controls[0].selectedIndex",
		},
		{
			name: "TabStrip rejects no-selection when tabs are nonempty",
			source: "controls:\n  - id: tabs\n    name: Tabs\n    type: TabStrip\n    selectedIndex: -1\n" +
				"    tabs:\n      - name: First\n",
			code: selectedIndexValidationCode, field: "controls[0].selectedIndex",
		},
		{
			name:   "selectedIndex accepts no-selection sentinel",
			source: "controls:\n  - id: tabs\n    name: Tabs\n    type: TabStrip\n    selectedIndex: -1\n    tabs: []\n",
		},
		{
			name: "MultiPage rejects no-selection when Page topology is known nonempty",
			source: "controls:\n  - id: pages\n    name: Pages\n    type: MultiPage\n    selectedIndex: -1\n" +
				"  - id: page\n    parentId: pages\n    name: Page\n    type: Page\n",
			code: selectedIndexValidationCode, field: "controls[0].selectedIndex",
		},
		{
			name:   "MultiPage rejects an index when known empty",
			source: "controls:\n  - id: pages\n    name: Pages\n    type: MultiPage\n    selectedIndex: 0\n    controls: []\n",
			code:   selectedIndexValidationCode, field: "controls[0].selectedIndex",
		},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			source := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: TestForm\n" + test.source
			issues, err := ValidateFormSpecSource(SpecInput{Format: "yaml"}, []byte(source))
			if err != nil {
				t.Fatal(err)
			}
			if test.code == "" {
				if hasValidationErrors(issues) {
					t.Fatalf("issues = %#v", issues)
				}
				return
			}
			if !hasValidationIssue(issues, test.code, test.field) {
				t.Fatalf("issues = %#v, want %s at %s", issues, test.code, test.field)
			}
		})
	}
}

func TestImagePictureSourceContract(t *testing.T) {
	for _, test := range []struct {
		name        string
		controlType string
		action      string
		valid       bool
	}{
		{"path", "Image", "picture: {path: assets/logo.jpg}", true},
		{"remove", "Image", "picture: {remove: true}", true},
		{"null", "Image", "picture: null", false},
		{"empty", "Image", "picture: {}", false},
		{"false-remove", "Image", "picture: {remove: false}", false},
		{"mixed", "Image", "picture: {path: assets/logo.bmp, remove: true}", false},
		{"unknown", "Image", "picture: {data: opaque}", false},
		{"empty-path", "Image", "picture: {path: '  '} ", false},
		{"wrong-control", "TextBox", "picture: {path: assets/logo.bmp}", false},
	} {
		t.Run(test.name, func(t *testing.T) {
			control := "  - id: image\n    name: Logo\n    type: " + test.controlType + "\n    " + test.action + "\n"
			source := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: Form\ncontrols:\n" + control
			issues, err := ValidateFormSpecSource(SpecInput{Format: "yaml"}, []byte(source))
			if err != nil {
				t.Fatal(err)
			}
			if test.valid == hasValidationErrors(issues) {
				t.Fatalf("valid=%t issues=%#v", test.valid, issues)
			}
		})
	}
}

func TestFormSpecPictureDataAndRemoveFalseAreNotSerialized(t *testing.T) {
	control := FormSpecControl{
		Type: "Image", Name: "Logo",
		Picture: &FormSpecPicture{Path: "assets/logo.bmp", Data: []byte{1, 2, 3}},
	}
	jsonBody, err := json.Marshal(control)
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(string(jsonBody), "Data") || strings.Contains(string(jsonBody), "\"data\"") || strings.Contains(string(jsonBody), "\"remove\"") {
		t.Fatalf("internal bytes or false remove were serialized: %s", jsonBody)
	}
	yamlBody, err := yaml.Marshal(control)
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(string(yamlBody), "data:") || strings.Contains(string(yamlBody), "remove:") {
		t.Fatalf("internal bytes or false remove were serialized: %s", yamlBody)
	}
}

func TestValidateFormSpecStrictSelectedIndexUsesKnownPageTopology(t *testing.T) {
	for _, test := range []struct {
		name     string
		controls []FormSpecControl
		invalid  bool
	}{
		{
			name: "partial MultiPage defers unknown page bounds",
			controls: []FormSpecControl{
				{ID: "pages", Name: "Pages", Type: "MultiPage", SelectedIndex: new(8)},
			},
		},
		{
			name: "nonempty MultiPage rejects no selection",
			controls: []FormSpecControl{
				{ID: "pages", Name: "Pages", Type: "MultiPage", SelectedIndex: new(-1)},
				{ID: "page", ParentID: "pages", Name: "Page", Type: "Page"},
			},
			invalid: true,
		},
		{
			name: "explicitly empty MultiPage rejects a nonnegative index",
			controls: []FormSpecControl{
				{ID: "pages", Name: "Pages", Type: "MultiPage", SelectedIndex: new(0), Controls: []FormSpecControl{}},
			},
			invalid: true,
		},
	} {
		t.Run(test.name, func(t *testing.T) {
			spec := FormSpec{
				SchemaVersion: 1,
				Kind:          "xlflow.userform",
				Basis:         "designer",
				Form:          FormSpecForm{Name: "TestForm"},
				Controls:      test.controls,
			}
			issues := ValidateFormSpecStrict(spec)
			if got := hasValidationIssue(issues, selectedIndexValidationCode, "controls[0].selectedIndex"); got != test.invalid {
				t.Fatalf("selectedIndex issue present = %t, want %t; issues = %#v", got, test.invalid, issues)
			}
		})
	}
}

func TestNormalizeFormSpecPreservesObservedOuterDimensionsForClientBuild(t *testing.T) {
	outerWidth, outerHeight := 308.0, 372.0
	clientWidth, clientHeight := 280.0, 344.0
	spec := NormalizeFormSpec(FormSpec{
		Form: FormSpecForm{
			Name: "UserForm1",
			Observed: &FormSpecObservedForm{
				Width:  &outerWidth,
				Height: &outerHeight,
			},
			Build: &FormSpecBuildForm{
				ClientWidth:  &clientWidth,
				ClientHeight: &clientHeight,
			},
		},
	})
	if spec.Form.Observed == nil || spec.Form.Observed.Width == nil || *spec.Form.Observed.Width != outerWidth || spec.Form.Observed.Height == nil || *spec.Form.Observed.Height != outerHeight {
		t.Fatalf("observed dimensions = %#v, want outer dimensions", spec.Form.Observed)
	}
	if spec.Form.Width != nil || spec.Form.Height != nil {
		t.Fatalf("normalization synthesized legacy top-level dimensions: %#v/%#v", spec.Form.Width, spec.Form.Height)
	}
	if spec.Form.Build == nil || spec.Form.Build.Width != nil || spec.Form.Build.Height != nil {
		t.Fatalf("client build inherited legacy outer dimensions: %#v", spec.Form.Build)
	}
	if spec.Form.Build.ClientWidth == nil || *spec.Form.Build.ClientWidth != clientWidth || spec.Form.Build.ClientHeight == nil || *spec.Form.Build.ClientHeight != clientHeight {
		t.Fatalf("client build dimensions = %#v", spec.Form.Build)
	}

	validSpec := spec
	validSpec.SchemaVersion = 1
	validSpec.Kind = "xlflow.userform"
	validSpec.Basis = "designer"
	if err := ValidateFormSpec(validSpec); err != nil {
		t.Fatalf("client build with observed outer dimensions should validate: %v", err)
	}
	jsonBody, err := MarshalSnapshot("json", spec)
	if err != nil {
		t.Fatal(err)
	}
	for _, field := range []string{"\"clientWidth\"", "\"clientHeight\""} {
		if !strings.Contains(string(jsonBody), field) {
			t.Fatalf("JSON snapshot missing %s: %s", field, jsonBody)
		}
	}
}

func TestValidateFormSpecRejectsConflictingBuildDimensions(t *testing.T) {
	spec := FormSpec{
		SchemaVersion: 1,
		Kind:          "xlflow.userform",
		Basis:         "designer",
		Form: FormSpecForm{
			Name: "UserForm1",
			Build: &FormSpecBuildForm{
				Width:       ptrFloat(308),
				ClientWidth: ptrFloat(280),
			},
		},
	}
	if !hasValidationIssue(ValidateFormSpecStrict(spec), "UFV017", "form.build") {
		t.Fatalf("missing build dimension conflict: %+v", ValidateFormSpecStrict(spec))
	}

	spec.Form.Width = ptrFloat(308)
	spec.Form.Build.Width = nil
	if !hasValidationIssue(ValidateFormSpecStrict(spec), "UFV017", "form.build") {
		t.Fatalf("missing top-level legacy/build client dimension conflict: %+v", ValidateFormSpecStrict(spec))
	}
}

func TestFormSpecFromInspectSnapshotOmitsUnsupportedBuiltInValue(t *testing.T) {
	spec, err := FormSpecFromInspectSnapshot(map[string]any{
		"name": "SnapshotForm",
		"controls": []any{map[string]any{
			"name": "Button1", "type": "CommandButton", "value": "False",
		}},
	})
	if err != nil {
		t.Fatal(err)
	}
	control := spec.Controls[0]
	if control.Value != nil {
		t.Fatalf("CommandButton value = %#v, want omitted", control.Value)
	}
	if !slices.Contains(control.Unsupported, "value") {
		t.Fatalf("unsupported = %q, want value", control.Unsupported)
	}
	if len(spec.Warnings) != 1 || spec.Warnings[0].Code != "unsupported_properties" {
		t.Fatalf("warnings = %#v, want one unsupported_properties warning", spec.Warnings)
	}

	path := filepath.Join(t.TempDir(), "SnapshotForm.json")
	if err := WriteSnapshot(SnapshotOutput{Path: path, DisplayPath: path, Format: "json"}, spec); err != nil {
		t.Fatalf("WriteSnapshot: %v", err)
	}
	if _, err := LoadFormSpec(SpecInput{Path: path, DisplayPath: path, Format: "json"}); err != nil {
		t.Fatalf("LoadFormSpec: %v", err)
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
	marshaledJSON, err := MarshalSnapshot("json", spec)
	if err != nil || !bytes.Equal(marshaledJSON, jsonBody) {
		t.Fatalf("MarshalSnapshot JSON mismatch: err=%v\n%s", err, marshaledJSON)
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
	marshaledYAML, err := MarshalSnapshot("yaml", spec)
	if err != nil || !bytes.Equal(marshaledYAML, yamlBody) {
		t.Fatalf("MarshalSnapshot YAML mismatch: err=%v\n%s", err, marshaledYAML)
	}
	if _, err := MarshalSnapshot("toml", spec); err == nil {
		t.Fatal("MarshalSnapshot accepted unsupported format")
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

func TestValidateFormSpecSourceRejectsNonFiniteAndOutOfRangeDimensions(t *testing.T) {
	body := []byte(`schemaVersion: 1
kind: xlflow.userform
basis: designer
form:
  name: UserForm1
  build:
    clientWidth: .nan
    clientHeight: .inf
    height: -1
controls: []
warnings: []
`)
	issues, err := ValidateFormSpecSource(SpecInput{Format: "yaml", DisplayPath: "UserForm1.yaml"}, body)
	if err != nil {
		t.Fatal(err)
	}
	if !hasValidationIssue(issues, "UFV002", "form.build.clientWidth") || !hasValidationIssue(issues, "UFV002", "form.build.clientHeight") {
		t.Fatalf("non-finite dimensions were not rejected as invalid numbers: %+v", issues)
	}
	if hasValidationIssue(issues, "UFV016", "form.build.clientWidth") || hasValidationIssue(issues, "UFV016", "form.build.clientHeight") || !hasValidationIssue(issues, "UFV016", "form.build.height") {
		t.Fatalf("dimension range issues = %+v", issues)
	}
}

func TestValidateFormSpecSourceRejectsOuterAndClientBuildDimensionsTogether(t *testing.T) {
	body := []byte(`schemaVersion: 1
kind: xlflow.userform
basis: designer
form:
  name: UserForm1
  width: 308
  build:
    clientWidth: 280
controls: []
warnings: []
`)
	issues, err := ValidateFormSpecSource(SpecInput{Format: "yaml", DisplayPath: "UserForm1.yaml"}, body)
	if err != nil {
		t.Fatal(err)
	}
	if !hasValidationIssue(issues, "UFV017", "form.build") {
		t.Fatalf("missing top-level/build client dimension conflict: %+v", issues)
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
		{"UFV010", "controls[0].parentId"},
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

func TestValidateFormSpecSourceReportsOnlyCycleMembersAtParentFields(t *testing.T) {
	body := []byte(`schemaVersion: 1
kind: xlflow.userform
basis: designer
form:
  name: UserForm1
controls:
  - id: frame_a
    name: FrameA
    type: Frame
    parentId: frame_b
  - id: frame_b
    name: FrameB
    type: Frame
    parentId: frame_a
  - id: label_child
    name: LabelChild
    type: Label
    parentId: frame_a
`)
	issues, err := ValidateFormSpecSource(SpecInput{Format: "yaml", DisplayPath: "UserForm1.yaml"}, body)
	if err != nil {
		t.Fatal(err)
	}
	cycles := make([]ValidationIssue, 0)
	for _, issue := range issues {
		if issue.Code == "UFV010" {
			cycles = append(cycles, issue)
		}
	}
	if len(cycles) != 1 || cycles[0].Field != "controls[0].parentId" {
		t.Fatalf("cycle issues = %+v, want one issue at controls[0].parentId", cycles)
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
    caption: Details
    text: Ready
    value: 7
    observed:
      caption: Details
      text: Ready
      value: 7
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
		{"UFV013", "controls[0].caption", SupportLevelCustomUnchecked},
		{"UFV013", "controls[0].observed.text", SupportLevelCustomUnchecked},
		{"UFV013", "form.width", SupportLevelBestEffort},
		{"UFV013", "warnings", SupportLevelSnapshotOnly},
	} {
		if !hasValidationIssueWithSupport(issues, want.code, want.field, want.support) {
			t.Fatalf("missing warning %s at %s support %s in %+v", want.code, want.field, want.support, issues)
		}
	}
}

func TestValidateFormSpecSourceAcceptsUnsupportedSnapshotPlaceholder(t *testing.T) {
	body := []byte(`schemaVersion: 1
kind: xlflow.userform
basis: designer
coordinateSystem: parent-relative
form:
  name: UserForm1
controls:
  - id: control_001
    name: VendorControl1
    type: Control
    unsupported:
      - controlType
warnings: []
`)
	issues, err := ValidateFormSpecSource(SpecInput{Format: "yaml", DisplayPath: "UserForm1.yaml"}, body)
	if err != nil {
		t.Fatal(err)
	}
	if hasValidationErrors(issues) {
		t.Fatalf("placeholder source issues = %#v", issues)
	}
	found := false
	for _, issue := range issues {
		if issue.Code == "UFV015" && issue.Severity == SeverityWarning && issue.Support == SupportLevelSnapshotOnly {
			found = true
		}
	}
	if !found {
		t.Fatalf("placeholder source issues = %#v, want UFV015 warning", issues)
	}
}

func TestValidateFormSpecForAuthoringRejectsUnsupportedSnapshotPlaceholder(t *testing.T) {
	spec := FormSpec{
		SchemaVersion: 1,
		Kind:          "xlflow.userform",
		Basis:         "designer",
		Form:          FormSpecForm{Name: "UserForm1"},
		Controls: []FormSpecControl{{
			ID: "control_001", Name: "VendorControl1", Type: UnsupportedControlPlaceholderType,
			Unsupported: []string{UnsupportedControlTypeProperty},
		}},
	}
	err := ValidateFormSpecForAuthoring(SpecInput{Format: "yaml", DisplayPath: "UserForm1.yaml"}, spec)
	var specErr *SpecError
	if !errors.As(err, &specErr) {
		t.Fatalf("authoring error = %#v, want SpecError", err)
	}
	if specErr.Code != "spec_validation_failed" || specErr.Field != "controls[0].type" || len(specErr.Issues) != 1 || specErr.Issues[0].Code != "UFV006" {
		t.Fatalf("authoring error = %#v", specErr)
	}
}

func TestValidateFormSpecForAuthoringRejectsClientDimensions(t *testing.T) {
	spec := FormSpec{
		SchemaVersion: 1,
		Kind:          "xlflow.userform",
		Basis:         "designer",
		Form: FormSpecForm{
			Name:  "UserForm1",
			Build: &FormSpecBuildForm{ClientWidth: ptrFloat(280)},
		},
	}
	err := ValidateFormSpecForAuthoring(SpecInput{Format: "yaml", DisplayPath: "UserForm1.yaml"}, spec)
	var specErr *SpecError
	if !errors.As(err, &specErr) {
		t.Fatalf("authoring error = %#v, want SpecError", err)
	}
	if specErr.Field != "form.build.clientWidth" || len(specErr.Issues) != 1 || specErr.Issues[0].Code != "UFV018" || !strings.Contains(specErr.Message, "refusing it instead of ignoring it") {
		t.Fatalf("authoring client dimension error = %#v", specErr)
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

func TestLoadFormSpecSelectsDeterministicFirstValidationIssue(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "UserForm1.form.yaml")
	body := `schemaVersion: 1
kind: xlflow.userform
basis: designer
unknownRoot: true
form:
  name: UserForm1
  zUnknown: true
controls:
  - id: label1
    name: Label1
    type: Label
    list: [A]
warnings: []
`
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	input, err := ResolveSpecInput(root, path)
	if err != nil {
		t.Fatal(err)
	}
	for i := 0; i < 20; i++ {
		_, err = LoadFormSpec(input)
		var specErr *SpecError
		if !errors.As(err, &specErr) {
			t.Fatalf("expected SpecError, got %T", err)
		}
		if specErr.Field != "unknownRoot" {
			t.Fatalf("iteration %d: field = %q, issues = %+v", i, specErr.Field, specErr.Issues)
		}
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
