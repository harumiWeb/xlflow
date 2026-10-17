package intel

import (
	"reflect"
	"slices"
	"strings"
	"testing"

	forms "github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func TestTypedPropertyChildCompletionUsesContract(t *testing.T) {
	for _, target := range []struct {
		name, source, field string
		properties          map[string]forms.PropertyContract
	}{
		{"form font", "form:\n  font:\n    ", "font", forms.UserFormContract().FormProperties},
		{"form color", "form:\n  backColor:\n    ", "backColor", forms.UserFormContract().FormProperties},
		{"Label font", "controls:\n  - type: Label\n    font:\n      ", "font", forms.ControlProperties("Label")},
		{"Label color", "controls:\n  - type: Label\n    foreColor:\n      ", "foreColor", forms.ControlProperties("Label")},
		{"Image background", "controls:\n  - type: Image\n    backColor:\n      ", "backColor", forms.ControlProperties("Image")},
		{"nested Frame child", "controls:\n  - type: Frame\n    controls:\n      - type: TextBox\n        font:\n          ", "font", forms.ControlProperties("TextBox")},
	} {
		t.Run(target.name, func(t *testing.T) {
			lines := strings.Split(target.source, "\n")
			items := CompleteYAML(target.source, Position{Line: len(lines) - 1, Character: len(lines[len(lines)-1])})
			var expected []string
			for name, property := range target.properties[target.field].Children {
				expected = append(expected, name)
				item, ok := completionByLabel(items, name)
				if !ok || !strings.Contains(item.Documentation, property.Description) || item.InsertText != name+": " {
					t.Fatalf("child contract %s lost: %+v", name, item)
				}
			}
			actual := labels(items)
			slices.Sort(actual)
			slices.Sort(expected)
			if !reflect.DeepEqual(actual, expected) {
				t.Fatalf("child fields = %v, want contract %v", actual, expected)
			}
		})
	}
}

func TestTypedPropertyChildValueCompletion(t *testing.T) {
	for _, test := range []struct {
		source string
		want   []string
	}{
		{"form:\n  font:\n    bold: ", []string{"false", "true"}},
		{"controls:\n  - type: Label\n    font:\n      italic: ", []string{"false", "true"}},
		{"controls:\n  - type: Label\n    font:\n      underline: f", []string{"false"}},
		{"form:\n  backColor:\n    kind: ", []string{"ole", "rgb", "system"}},
		{"controls:\n  - type: Image\n    backColor:\n      kind: s", []string{"system"}},
		{"form:\n  font:\n    charset: ", nil},
	} {
		lines := strings.Split(test.source, "\n")
		line := len(lines) - 1
		items := CompleteYAML(test.source, Position{Line: line, Character: len(lines[line])})
		actual := labels(items)
		slices.Sort(actual)
		if !slices.Equal(actual, test.want) {
			t.Fatalf("values in %q = %v, want %v", test.source, actual, test.want)
		}
	}
}

func TestTypedPropertyChildCompletionPreservesContext(t *testing.T) {
	source := "form:\n  font:\n    name: Tahoma\n    size: 8.25\n    st"
	items := CompleteYAML(source, Position{Line: 4, Character: 6})
	if len(items) != 1 || items[0].Label != "strikeThrough" || items[0].Replace != (Range{Start: Position{Line: 4, Character: 4}, End: Position{Line: 4, Character: 6}}) {
		t.Fatalf("incomplete child prefix/range: %+v", items)
	}
	source = "form:\n  font:\n    bold: false\n    "
	items = CompleteYAML(source, Position{Line: 3, Character: 4})
	if hasLabel(items, "bold") || !hasLabel(items, "italic") || hasLabel(items, "caption") {
		t.Fatalf("child siblings/context: %v", labels(items))
	}
}

func TestTypedPropertyChildrenRespectControlApplicability(t *testing.T) {
	for _, test := range []struct{ declaration, field, child string }{
		{"type: Image", "font", "bold"},
		{"type: Image", "foreColor", "kind"},
		{"type: SpinButton", "font", "name"},
		{"type: Label\n    progId: Vendor.Custom.1", "font", "bold"},
		{"type: Label\n    progId: Vendor.Custom.1", "backColor", "kind"},
		{"type: Label\n    progId: Forms.TextBox.1", "font", "name"},
		{"type: Image\n    progId: Vendor.Custom.1", "picture", "path"},
	} {
		prefix := "controls:\n  - " + test.declaration + "\n    " + test.field + ":\n      "
		lines := strings.Split(prefix, "\n")
		line := len(lines) - 1
		if items := CompleteYAML(prefix, Position{Line: line, Character: 6}); len(items) != 0 {
			t.Fatalf("unsupported children exposed in %q: %v", prefix, labels(items))
		}
		source := prefix + test.child + ": true\n"
		if hover := HoverYAML(source, Position{Line: line, Character: 7}); hover != nil {
			t.Fatalf("unsupported child hover in %q: %+v", source, hover)
		}
	}
}

func TestTypedPropertyChildHoverUsesContract(t *testing.T) {
	for _, test := range []struct {
		source, field, child, scope string
		properties                  map[string]forms.PropertyContract
	}{
		{"form:\n  font:\n    size: 8.25\n", "font", "size", "form font", forms.UserFormContract().FormProperties},
		{"form:\n  foreColor:\n    kind: system\n", "foreColor", "kind", "form foreColor", forms.UserFormContract().FormProperties},
		{"controls:\n  - type: TextBox\n    font:\n      charset: 128\n", "font", "charset", "TextBox font", forms.ControlProperties("TextBox")},
		{"controls:\n  - type: Image\n    backColor:\n      value: buttonFace\n", "backColor", "value", "Image backColor", forms.ControlProperties("Image")},
	} {
		lines := strings.Split(strings.TrimSuffix(test.source, "\n"), "\n")
		line := len(lines) - 1
		hover := HoverYAML(test.source, Position{Line: line, Character: strings.Index(lines[line], test.child) + 1})
		property := test.properties[test.field].Children[test.child]
		if hover == nil || !strings.Contains(hover.Contents, property.Description) || !strings.Contains(hover.Contents, test.scope) || !strings.Contains(hover.Contents, "**Type:** "+string(property.ValueType)) {
			t.Fatalf("child hover for %q: %+v", test.source, hover)
		}
		if hover.Range.Start.Line != line || hover.Range.End.Character-hover.Range.Start.Character != len(test.child) {
			t.Fatalf("child key hover range: %+v", hover.Range)
		}
	}
	if hover := HoverYAML("form:\n  font:\n    unknown: true\n", Position{Line: 2, Character: 5}); hover != nil {
		t.Fatalf("unknown nested field has hover: %+v", hover)
	}
}
