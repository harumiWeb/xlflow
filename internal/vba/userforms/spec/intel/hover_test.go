package intel

import (
	"strings"
	"testing"
)

func TestHoverYAMLUsesCanonicalContract(t *testing.T) {
	source := `schemaVersion: 1
kind: xlflow.userform
basis: designer
warnings: []
form:
  width: 100
  height: 120
  observed:
    insideWidth: 110
controls:
  - type: Page
    tag: page-tag
    controlTipText: Page hint
    accelerator: P
  - type: TabStrip
    tabs:
      - name: Main
        tag: tab-tag
        controlTipText: Tab hint
        accelerator: M
  - type: ComboBox
    progId: Forms.ComboBox.1
    selectedIndex: 0
    properties: {}
    unsupported: []
    observed:
      value: East
  - type: VendorWidget
    progId: Vendor.Widget.1
`
	tests := []struct {
		name string
		line int
		text string
		want []string
	}{
		{"Page tag", 11, "tag", []string{"### `tag`", "Applies to:** `Page`", "string"}},
		{"Page tooltip", 12, "controlTipText", []string{"### `controlTipText`", "Page"}},
		{"Tab tag", 17, "tag", []string{"### `tag`", "TabStrip tab", "string"}},
		{"Tab tooltip", 18, "controlTipText", []string{"### `controlTipText`", "TabStrip tab"}},
		{"root field", 0, "schemaVersion", []string{"### `schemaVersion`", "**Type:** integer", "**Required:** yes", "document root"}},
		{"snapshot warnings", 3, "warnings", []string{"snapshot-only", "not a guaranteed normal build input"}},
		{"supported form field", 5, "width", []string{"supported", "client content width in points"}},
		{"supported form height", 6, "height", []string{"supported", "client content height in points"}},
		{"snapshot field", 7, "observed", []string{"snapshot-only", "not a guaranteed normal build input"}},
		{"nested observed form field", 8, "insideWidth", []string{"snapshot-only", "captured form state"}},
		{"list state", 22, "selectedIndex", []string{"observed-only", "best-effort basis", "not guaranteed"}},
		{"properties", 23, "properties", []string{"custom/unchecked", "not an unrestricted build escape hatch"}},
		{"unsupported", 24, "unsupported", []string{"snapshot-only", "not a guaranteed normal build input"}},
		{"nested observed control field", 26, "value", []string{"snapshot-only", "Applies to:** `ComboBox`"}},
		{"built in type", 20, "ComboBox", []string{"### `ComboBox`", "Forms.ComboBox.1", "Container:** no"}},
		{"built in progid", 21, "Forms.ComboBox.1", []string{"### `Forms.ComboBox.1`", "ComboBox", "supported"}},
		{"custom progid", 28, "Vendor.Widget.1", []string{"custom/unchecked", "common structural fields", "installed control"}},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			column := strings.Index(sourceLine(source, tc.line), tc.text) + 1
			hover := HoverYAML(source, Position{Line: tc.line, Character: column})
			if hover == nil {
				t.Fatal("HoverYAML() = nil")
			}
			for _, want := range tc.want {
				if !strings.Contains(hover.Contents, want) {
					t.Fatalf("hover missing %q:\n%s", want, hover.Contents)
				}
			}
			if hover.Range.Start.Line != tc.line {
				t.Fatalf("hover range = %#v, want line %d", hover.Range, tc.line)
			}
		})
	}
}

func TestHoverYAMLReturnsNilForUnknownFields(t *testing.T) {
	if hover := HoverYAML("form:\n  unknown: true\n", Position{Line: 1, Character: 3}); hover != nil {
		t.Fatalf("HoverYAML() = %#v, want nil", hover)
	}
}

func TestHoverYAMLDocumentsImagePictureAction(t *testing.T) {
	source := "controls:\n  - type: Image\n    picture:\n      path: assets/logo.bmp\n"
	line := 3
	column := strings.Index(sourceLine(source, line), "path") + 1
	hover := HoverYAML(source, Position{Line: line, Character: column})
	if hover == nil {
		t.Fatal("HoverYAML() = nil for Image picture path")
	}
	for _, want := range []string{"Image picture action", "Project-root-relative BMP or JPEG asset path."} {
		if !strings.Contains(hover.Contents, want) {
			t.Fatalf("hover missing %q:\n%s", want, hover.Contents)
		}
	}
}

func sourceLine(source string, line int) string {
	return strings.Split(source, "\n")[line]
}
