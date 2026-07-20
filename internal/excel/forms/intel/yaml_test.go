package intel

import "testing"

func TestParseYAMLIndexesFieldPathsAndUTF16Ranges(t *testing.T) {
	doc := ParseYAML(`schemaVersion: 1
kind: xlflow.userform
controls:
  - id: label_😀
    parentId: frame_main
form:
  observed:
    caption: あ
`)
	if doc.ParseError != nil {
		t.Fatal(doc.ParseError)
	}
	parent, ok := doc.Field("controls[0].parentId")
	if !ok {
		t.Fatal("controls[0].parentId was not indexed")
	}
	if parent.KeyRange.Start != (Position{Line: 4, Character: 4}) || parent.ValueRange.Start != (Position{Line: 4, Character: 14}) {
		t.Fatalf("parentId ranges = %#v, want key/value starts on line 4", parent)
	}
	if path, ok := doc.FieldPathForNode(parent.Value); !ok || path != "controls[0].parentId" {
		t.Fatalf("FieldPathForNode(parentId value) = %q, %v", path, ok)
	}
	if _, ok := doc.Field("form.observed.caption"); !ok {
		t.Fatal("nested mapping field was not indexed")
	}
	id, ok := doc.Field("controls[0].id")
	if !ok {
		t.Fatal("controls[0].id was not indexed")
	}
	if id.ValueRange.End != (Position{Line: 3, Character: 16}) {
		t.Fatalf("UTF-16 value end = %#v, want line 3 character 16", id.ValueRange.End)
	}
}

func TestCursorContextFallsBackForIncompleteYAML(t *testing.T) {
	doc := ParseYAML("controls:\n  - type: TextBox\n    ca\n")
	if doc.ParseError == nil {
		t.Fatal("incomplete YAML should not parse as a complete document")
	}
	context := doc.CursorContextAt(Position{Line: 2, Character: 6})
	if context.ParentPath != "controls[0]" {
		t.Fatalf("ParentPath = %q, want controls[0]", context.ParentPath)
	}
	if context.SequenceIndex != 0 {
		t.Fatalf("SequenceIndex = %d, want 0", context.SequenceIndex)
	}
	if context.LinePrefix != "    ca" {
		t.Fatalf("LinePrefix = %q, want %q", context.LinePrefix, "    ca")
	}
	if len(context.ExistingKeys) != 1 || context.ExistingKeys[0] != "type" {
		t.Fatalf("ExistingKeys = %#v, want [type]", context.ExistingKeys)
	}
}

func TestCursorContextBlankLineRespectsItsIndentation(t *testing.T) {
	doc := ParseYAML("controls:\n  - type: Label\n\n")
	context := doc.CursorContextAt(Position{Line: 2, Character: 0})
	if context.ParentPath != "" {
		t.Fatalf("unindented blank line ParentPath = %q, want root", context.ParentPath)
	}

	doc = ParseYAML("controls:\n  - type: Label\n    \n")
	context = doc.CursorContextAt(Position{Line: 2, Character: 4})
	if context.ParentPath != "controls[0]" {
		t.Fatalf("indented blank line ParentPath = %q, want controls[0]", context.ParentPath)
	}
}
