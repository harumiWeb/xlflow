package edit

import (
	"cmp"
	"fmt"
	"math"
	"slices"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func TestOperationErrorAfterTemporaryInvalidState(t *testing.T) {
	diagnostics := assertFailure(t, base, []Operation{
		{Type: SetParent, ControlID: "submit", ParentID: "new"},
		{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: []string{"invalid"}},
	}, "UFE002")
	d := diagnostics[0]
	if d.OperationIndex != 1 || d.ControlID != "submit" || d.Field != "caption" || !strings.Contains(d.Message, "string, number or boolean") {
		t.Fatalf("operation cause lost: %+v", diagnostics)
	}
}

func TestAddDuplicateUsesCanonicalDiagnostic(t *testing.T) {
	control := &spec.FormSpecControl{ID: "other", Name: "Added", Type: "Label"}
	diagnostics := assertFailure(t, base, []Operation{{Type: AddControl, Control: control}}, "UFV007")
	// Inspect the canonical parser's actual source diagnostic for comparison.
	e, err := newEngine([]byte(base))
	if err != nil {
		t.Fatal(err)
	}
	if err := e.apply(Operation{Type: AddControl, Control: control}); err != nil {
		t.Fatalf("duplicate rejected before canonical validation: %v", err)
	}
	_, canonicalErr := spec.ParseFormSpec(spec.SpecInput{Format: "yaml"}, e.source)
	canonical := validationError(0, Operation{}, canonicalErr)
	for _, d := range diagnostics {
		if d.Code == "UFV007" {
			if d.ControlID != "other" || d.Operation != AddControl || d.OperationIndex != 0 {
				t.Fatalf("add target context lost: %+v", d)
			}
			if !slices.ContainsFunc(canonical.Diagnostics, func(c Diagnostic) bool {
				return c.Code == d.Code && c.Message == d.Message && c.Suggestion == d.Suggestion && c.Field == d.Field
			}) {
				t.Fatalf("noncanonical diagnostic: %+v", d)
			}
		}
	}
}

func TestAddPayloadFailureRetainsControlID(t *testing.T) {
	diagnostics := assertFailure(t, base, []Operation{{Type: AddControl, Control: &spec.FormSpecControl{ID: "new", Type: "Label"}}}, "UFE003")
	if diagnostics[0].ControlID != "new" || diagnostics[0].Operation != AddControl {
		t.Fatalf("missing add target: %+v", diagnostics)
	}
}

func TestMultilineQuotedCommentStaysOutsideReplacement(t *testing.T) {
	for _, newline := range []string{"\n", "\r\n"} {
		for _, quote := range []string{"'", `"`} {
			for _, tag := range []string{"", "!!str "} {
				// The same comment text also occurs inside the scalar: substring
				// searching cannot establish that the actual comment was consumed.
				token := tag + quote + "first # inline\n      second" + quote
				source := strings.Replace(base, "'Submit'", token, 1)
				source = strings.ReplaceAll(source, "\n", newline)
				result := applyOK(t, source, Operation{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: "changed"})
				want := strings.Replace(source, strings.ReplaceAll(token, "\n", newline), quote+"changed"+quote, 1)
				if string(result.Source) != want {
					t.Fatalf("comment changed:\n%s\nwant:\n%s", result.Source, want)
				}
			}
		}
	}
}

func TestRemoveLastIndentlessSequenceItem(t *testing.T) {
	for _, newline := range []string{"\n", "\r\n"} {
		for _, nested := range []bool{false, true} {
			source := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform: {name: Main}\ncontrols:\n- id: child\n  name: Child\n  type: Label\nwarnings: [] # keep\n"
			want := strings.Replace(source, "- id: child\n  name: Child\n  type: Label\n", "  []\n", 1)
			if nested {
				source = strings.Replace(hierarchy, "      - id: child\n        name: Child\n        type: TextBox\n        left: 5\n        top: 8", "    - id: child\n      name: Child\n      type: TextBox", 1)
				want = strings.Replace(source, "    - id: child\n      name: Child\n      type: TextBox\n", "      []\n", 1)
			}
			result := applyOK(t, strings.ReplaceAll(source, "\n", newline), Operation{Type: RemoveControl, ControlID: "child"})
			if string(result.Source) != strings.ReplaceAll(want, "\n", newline) {
				t.Fatalf("unexpected removal:\n%s", result.Source)
			}
		}
	}
}

func TestReorderPreservesSparseSiblings(t *testing.T) {
	for _, zs := range [][]int{{0, 10, 20}, {math.MinInt, 0, math.MaxInt}, {0, 1, 2, 100}} {
		source := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform: {name: Main}\ncontrols:\n"
		for i, z := range zs {
			source += fmt.Sprintf("  - {id: c%d, name: C%d, type: Label, zIndex: %d, tabIndex: %d}\n", i, i, z, i)
		}
		id := fmt.Sprintf("c%d", len(zs)-1)
		result := applyOK(t, source, Operation{Type: ReorderControl, ControlID: id, Index: new(1)})
		ordered := slices.Clone(result.Document.Controls)
		slices.SortStableFunc(ordered, func(a, b spec.FormSpecControl) int {
			return cmp.Compare(*a.ZIndex, *b.ZIndex)
		})
		if ordered[1].ID != id {
			t.Fatalf("incorrect order: %+v", ordered)
		}
		if len(zs) == 3 && *result.Document.Controls[0].ZIndex != zs[0] {
			t.Fatal("untouched prefix rewritten")
		}
		if len(zs) == 3 && *result.Document.Controls[1].ZIndex != zs[1] {
			t.Fatal("available gap ignored")
		}
		if len(zs) == 4 && *result.Document.Controls[2].ZIndex != zs[2] {
			t.Fatal("smallest contiguous range expanded unnecessarily")
		}
		for i, c := range result.Document.Controls {
			if *c.TabIndex != i {
				t.Fatal("tabIndex changed")
			}
		}
		again := applyOK(t, string(result.Source), Operation{Type: ReorderControl, ControlID: id, Index: new(1)})
		if len(again.Edits) != 0 {
			t.Fatal("repeat reorder changed source")
		}
	}
}

func TestReorderIntegerLimitsAndTies(t *testing.T) {
	for _, zs := range [][]int{{0, 0, 0}, {math.MinInt, math.MinInt, math.MinInt + 1}, {math.MaxInt - 1, math.MaxInt, math.MaxInt}, {0, 1, 2}, {0, 10, 20}} {
		for from := range len(zs) {
			for to := range len(zs) {
				source := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform: {name: Main}\ncontrols:\n"
				var expected []string
				for i, z := range zs {
					source += fmt.Sprintf("  - {id: c%d, name: C%d, type: Label, zIndex: %d}\n", i, i, z)
					expected = append(expected, fmt.Sprintf("c%d", i))
				}
				id := expected[from]
				expected = slices.Delete(expected, from, from+1)
				expected = slices.Insert(expected, to, id)
				result := applyOK(t, source, Operation{Type: ReorderControl, ControlID: id, Index: new(to)})
				ordered := slices.Clone(result.Document.Controls)
				slices.SortStableFunc(ordered, func(a, b spec.FormSpecControl) int {
					return cmp.Compare(*a.ZIndex, *b.ZIndex)
				})
				for i, c := range ordered {
					if c.ID != expected[i] {
						t.Fatalf("%v from=%d to=%d: got %+v, want %v", zs, from, to, ordered, expected)
					}
				}
			}
		}
	}
}
