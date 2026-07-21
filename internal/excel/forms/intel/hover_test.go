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
  observed: {}
controls:
  - type: ComboBox
    progId: Forms.ComboBox.1
    selectedIndex: 0
    properties: {}
    unsupported: []
  - type: VendorWidget
    progId: Vendor.Widget.1
`
	tests := []struct {
		name string
		line int
		text string
		want []string
	}{
		{"root field", 0, "schemaVersion", []string{"### `schemaVersion`", "**Type:** integer", "**Required:** yes", "document root"}},
		{"snapshot warnings", 3, "warnings", []string{"snapshot-only", "not a guaranteed normal build input"}},
		{"best effort form field", 5, "width", []string{"best-effort", "inspect the rebuilt Designer"}},
		{"snapshot field", 6, "observed", []string{"snapshot-only", "not a guaranteed normal build input"}},
		{"list state", 10, "selectedIndex", []string{"observed-only", "best-effort basis", "not guaranteed"}},
		{"properties", 11, "properties", []string{"custom/unchecked", "not an unrestricted build escape hatch"}},
		{"unsupported", 12, "unsupported", []string{"snapshot-only", "not a guaranteed normal build input"}},
		{"built in type", 8, "ComboBox", []string{"### `ComboBox`", "Forms.ComboBox.1", "Container:** no"}},
		{"built in progid", 9, "Forms.ComboBox.1", []string{"### `Forms.ComboBox.1`", "ComboBox", "supported"}},
		{"custom progid", 14, "Vendor.Widget.1", []string{"custom/unchecked", "common structural fields", "installed control"}},
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

func sourceLine(source string, line int) string {
	return strings.Split(source, "\n")[line]
}
