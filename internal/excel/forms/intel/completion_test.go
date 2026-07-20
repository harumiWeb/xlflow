package intel

import (
	"strings"
	"testing"
)

func TestCompleteYAMLRootFormAndControlProperties(t *testing.T) {
	t.Run("root omits present fields", func(t *testing.T) {
		items := CompleteYAML("schemaVersion: 1\n", Position{Line: 1})
		if hasLabel(items, "schemaVersion") || !hasLabel(items, "kind") || !hasLabel(items, "controls") {
			t.Fatalf("root items = %#v", labels(items))
		}
	})
	t.Run("form includes build but hides observed by default", func(t *testing.T) {
		items := CompleteYAML("form:\n  \n", Position{Line: 1, Character: 2})
		if !hasLabel(items, "name") || !hasLabel(items, "build") || hasLabel(items, "observed") {
			t.Fatalf("form items = %#v", labels(items))
		}
	})
	t.Run("known control only offers applicable properties", func(t *testing.T) {
		items := CompleteYAML("controls:\n  - type: Label\n    \n", Position{Line: 2, Character: 4})
		if !hasLabel(items, "caption") || hasLabel(items, "text") || hasLabel(items, "list") || hasLabel(items, "selectedIndex") {
			t.Fatalf("Label items = %#v", labels(items))
		}
		caption, _ := completionByLabel(items, "caption")
		if !strings.Contains(caption.Detail, "string") || !strings.Contains(caption.Documentation, "Applies to: Label.") || !strings.Contains(caption.Documentation, "Support: **supported**") {
			t.Fatalf("caption documentation = %#v", caption)
		}
	})
	t.Run("snapshot field appears only after it is typed", func(t *testing.T) {
		items := CompleteYAML("form:\n  ob", Position{Line: 1, Character: 4})
		observed, ok := completionByLabel(items, "observed")
		if !ok || observed.SortText != "z" {
			t.Fatalf("snapshot completion = %#v", observed)
		}
	})
	t.Run("custom ProgID is restricted to common properties", func(t *testing.T) {
		items := CompleteYAML("controls:\n  - type: Label\n    progId: Vendor.Custom.1\n    \n", Position{Line: 3, Character: 4})
		if !hasLabel(items, "id") || hasLabel(items, "caption") {
			t.Fatalf("custom items = %#v", labels(items))
		}
	})
}

func TestCompleteYAMLFixedAndScalarValues(t *testing.T) {
	t.Run("fixed root value", func(t *testing.T) {
		items := CompleteYAML("kind: \n", Position{Line: 0, Character: len("kind: ")})
		if !hasLabel(items, "xlflow.userform") {
			t.Fatalf("kind values = %#v", labels(items))
		}
	})
	t.Run("control type and matching ProgID", func(t *testing.T) {
		typeItems := CompleteYAML("controls:\n  - type: \n", Position{Line: 1, Character: len("  - type: ")})
		if !hasLabel(typeItems, "TextBox") || !hasLabel(typeItems, "Frame") {
			t.Fatalf("type values = %#v", labels(typeItems))
		}
		progIDItems := CompleteYAML("controls:\n  - type: Label\n    progId: \n", Position{Line: 2, Character: len("    progId: ")})
		if len(progIDItems) != 1 || progIDItems[0].Label != "Forms.Label.1" {
			t.Fatalf("ProgID values = %#v", labels(progIDItems))
		}
	})
	t.Run("boolean", func(t *testing.T) {
		items := CompleteYAML("controls:\n  - type: Label\n    enabled: \n", Position{Line: 2, Character: len("    enabled: ")})
		if !hasLabel(items, "true") || !hasLabel(items, "false") {
			t.Fatalf("boolean values = %#v", labels(items))
		}
	})
}

func TestCompleteYAMLParentReferencesAreSafeAndCasePreserving(t *testing.T) {
	source := `controls:
  - id: Frame_Main
    name: Main frame
    type: Frame
  - id: Label_One
    name: Label one
    type: Label
  - id: Custom_Container
    name: Custom host
    type: VendorControl
    progId: Vendor.Container.1
  - id: Unknown_No_ProgID
    name: Invalid unknown parent
    type: UnknownControl
  - id: Duplicate
    name: duplicate A
    type: Frame
  - id: Duplicate
    name: duplicate B
    type: Frame
  - id: Child
    name: Nested child
    type: Label
    parentId: Current
  - id: Current
    name: Current control
    type: Label
    parentId: 
`
	items := CompleteYAML(source, Position{Line: 27, Character: len("    parentId: ")})
	if !hasLabel(items, "Frame_Main") || !hasLabel(items, "Custom_Container") {
		t.Fatalf("parent values = %#v", labels(items))
	}
	if hasLabel(items, "Label_One") || hasLabel(items, "Current") || hasLabel(items, "Child") || hasLabel(items, "Duplicate") || hasLabel(items, "Unknown_No_ProgID") {
		t.Fatalf("unsafe parent values = %#v", labels(items))
	}
	frame, ok := completionByLabel(items, "Frame_Main")
	if !ok || frame.InsertText != "Frame_Main" || !strings.Contains(frame.Detail, "Main frame (Frame)") {
		t.Fatalf("Frame completion = %#v", frame)
	}
	if items[0].Label != "Frame_Main" {
		t.Fatalf("container candidate should rank first: %#v", labels(items))
	}
}

func TestCompleteYAMLIncompleteAndLegacyNestedControls(t *testing.T) {
	t.Run("incomplete property prefix", func(t *testing.T) {
		items := CompleteYAML("controls:\n  - type: TextBox\n    te", Position{Line: 2, Character: 6})
		if !hasLabel(items, "text") {
			t.Fatalf("incomplete items = %#v", labels(items))
		}
	})
	t.Run("indentationless sequence", func(t *testing.T) {
		items := CompleteYAML("controls:\n- type: Label\n  ca", Position{Line: 2, Character: 4})
		if !hasLabel(items, "caption") {
			t.Fatalf("indentationless items = %#v", labels(items))
		}
	})
	t.Run("legacy nested control is visible to parent completion", func(t *testing.T) {
		source := "controls:\n  - id: FrameMain\n    name: Main\n    type: Frame\n    controls:\n      - id: NestedFrame\n        name: Nested\n        type: Frame\n  - id: Child\n    name: Child\n    type: Label\n    parentId: \n"
		items := CompleteYAML(source, Position{Line: 11, Character: len("    parentId: ")})
		if !hasLabel(items, "FrameMain") || !hasLabel(items, "NestedFrame") {
			t.Fatalf("legacy values = %#v", labels(items))
		}
	})
}

func TestCompleteYAMLSnippets(t *testing.T) {
	skeleton := CompleteYAML("", Position{})
	if len(skeleton) != 1 || !skeleton[0].Snippet || !strings.Contains(skeleton[0].InsertText, "schemaVersion: 1") {
		t.Fatalf("skeleton = %#v", skeleton)
	}

	items := CompleteYAML("controls:\n  - ", Position{Line: 1, Character: len("  - ")})
	basic, ok := completionByLabel(items, "Basic control")
	if !ok || strings.HasPrefix(basic.InsertText, "- ") || !basic.Snippet {
		t.Fatalf("existing dash basic snippet = %#v", basic)
	}

	items = CompleteYAML("controls:\n- ", Position{Line: 1, Character: len("- ")})
	basic, ok = completionByLabel(items, "Basic control")
	if !ok || strings.HasPrefix(basic.InsertText, "- ") {
		t.Fatalf("zero-indent existing dash basic snippet = %#v", basic)
	}

	items = CompleteYAML("controls:\n  \n", Position{Line: 1, Character: 2})
	basic, ok = completionByLabel(items, "Basic control")
	if !ok || !strings.HasPrefix(basic.InsertText, "- ") {
		t.Fatalf("new item basic snippet = %#v", basic)
	}
	for _, label := range []string{"Frame", "Label", "TextBox", "CommandButton"} {
		if !hasLabel(items, label) {
			t.Fatalf("missing %s snippet: %#v", label, labels(items))
		}
	}
}

func labels(items []CompletionItem) []string {
	labels := make([]string, len(items))
	for i, item := range items {
		labels[i] = item.Label
	}
	return labels
}

func hasLabel(items []CompletionItem, label string) bool {
	_, ok := completionByLabel(items, label)
	return ok
}

func completionByLabel(items []CompletionItem, label string) (CompletionItem, bool) {
	for _, item := range items {
		if item.Label == label {
			return item, true
		}
	}
	return CompletionItem{}, false
}
