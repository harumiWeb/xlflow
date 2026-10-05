package edit

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"math"
	"reflect"
	"slices"
	"strings"
	"testing"
	"unicode/utf8"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/compiler"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

const base = "# document\nschemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: Main\ncontrols:\n  # button\n  - id: submit\n    name: Submit\n    type: CommandButton\n    caption: 'Submit' # inline\n    left: 100\n    top: 20\n  # unrelated\n  - id: other\n    name: Other\n    type: TextBox\n    text: keep\nwarnings: [] # tail\n"

func applyOK(t *testing.T, source string, ops ...Operation) Result {
	t.Helper()
	input := []byte(source)
	copyInput := bytes.Clone(input)
	result, err := Apply(spec.SpecInput{Format: "yaml"}, input, ops)
	if err != nil {
		t.Fatalf("Apply: %v (%#v)", err, err)
	}
	if !bytes.Equal(input, copyInput) {
		t.Fatal("input mutated")
	}
	rebuilt := bytes.Clone(input)
	for i := len(result.Edits) - 1; i >= 0; i-- {
		ed := result.Edits[i]
		if i > 0 && result.Edits[i-1].End > ed.Start {
			t.Fatal("overlapping edits")
		}
		next := append(bytes.Clone(rebuilt[:ed.Start]), []byte(ed.Text)...)
		rebuilt = append(next, rebuilt[ed.End:]...)
	}
	if !bytes.Equal(rebuilt, result.Source) {
		t.Fatalf("edits do not rebuild result:\n%s\nwant\n%s", rebuilt, result.Source)
	}
	parsed, err := spec.ParseFormSpec(spec.SpecInput{Format: "yaml"}, result.Source)
	if err != nil || !reflect.DeepEqual(parsed, result.Document) {
		t.Fatalf("result not canonical: %v", err)
	}
	return result
}

func assertFailure(t *testing.T, source string, ops []Operation, code string) []Diagnostic {
	t.Helper()
	input := []byte(source)
	original := bytes.Clone(input)
	result, err := Apply(spec.SpecInput{Format: "yaml"}, input, ops)
	if err == nil {
		t.Fatalf("expected failure, got:\n%s", result.Source)
	}
	editErr, ok := errors.AsType[*Error](err)
	if !ok {
		t.Fatalf("unstructured error: %T", err)
	}
	if !reflect.DeepEqual(result, Result{}) || !bytes.Equal(input, original) {
		t.Fatal("partial result or input mutation")
	}
	if code != "" && !slices.ContainsFunc(editErr.Diagnostics, func(d Diagnostic) bool { return d.Code == code }) {
		t.Fatalf("missing %s: %+v", code, editErr.Diagnostics)
	}
	return editErr.Diagnostics
}

func TestScalarEditsPreserveSource(t *testing.T) {
	for _, newline := range []string{"\n", "\r\n"} {
		t.Run(strings.ReplaceAll(newline, "\n", "LF"), func(t *testing.T) {
			source := strings.ReplaceAll(base, "\n", newline)
			result := applyOK(t, source, Operation{Type: MoveControl, ControlID: "submit", Left: new(112.25), Top: new(20.0)})
			want := strings.Replace(source, "left: 100", "left: 112.25", 1)
			if string(result.Source) != want {
				t.Fatalf("nonlocal edit:\n%s", result.Source)
			}
			if len(result.Edits) != 1 {
				t.Fatalf("edits: %+v", result.Edits)
			}
			again := applyOK(t, string(result.Source), Operation{Type: MoveControl, ControlID: "submit", Left: new(112.25), Top: new(20.0)})
			if len(again.Edits) != 0 {
				t.Fatalf("format churn: %+v", again.Edits)
			}
		})
	}
}

func TestTaggedScalarEditsPreserveComments(t *testing.T) {
	for _, newline := range []string{"\n", "\r\n"} {
		for _, token := range []string{
			`!!str "alpha # omega"`,
			`!!str 'alpha '' # omega'`,
			`!<tag:yaml.org,2002:str> "alpha \" # omega"`,
			`!!str "alpha , ] } # omega"`,
		} {
			for _, flow := range []bool{false, true} {
				t.Run(fmt.Sprintf("%q/%s/flow=%t", newline, token, flow), func(t *testing.T) {
					source := strings.Replace(base, "'Submit'", token, 1)
					if flow {
						source = "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform: {name: Main}\ncontrols: [{id: submit, name: Submit, type: CommandButton, caption: " + token + ", left: 100}] # inline\n"
					}
					source = strings.ReplaceAll(source, "\n", newline)
					result := applyOK(t, source, Operation{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: "changed"})
					quote := `"`
					if strings.Contains(token, "'alpha") {
						quote = "'"
					}
					want := strings.Replace(source, token, quote+"changed"+quote, 1)
					if string(result.Source) != want {
						t.Fatalf("tagged scalar changed unrelated bytes:\n%s\nwant:\n%s", result.Source, want)
					}
				})
			}
		}
	}
}

func TestPropertyInsertionAndExplicitValues(t *testing.T) {
	result := applyOK(t, base,
		Operation{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: ""},
		Operation{Type: SetControlProperty, ControlID: "submit", Field: "enabled", Value: false},
		Operation{Type: ResizeControl, ControlID: "submit", Width: new(72.5), Height: new(24.0)},
		Operation{Type: SetControlProperty, ControlID: "other", Field: "tabIndex", Value: 0},
	)
	if !strings.Contains(string(result.Source), "caption: '' # inline") || !strings.Contains(string(result.Source), "enabled: false") || !strings.Contains(string(result.Source), "tabIndex: 0") {
		t.Fatalf("explicit values lost:\n%s", result.Source)
	}
	if strings.Contains(string(result.Source), "observed:") {
		t.Fatal("normalization leaked into source")
	}
	if !strings.Contains(string(result.Source), "  # unrelated\n  - id: other\n    name: Other\n    type: TextBox\n    text: keep\n") {
		t.Fatal("unrelated source changed")
	}
}

func TestQuotesUnicodeAndMultiline(t *testing.T) {
	cases := []struct {
		value       string
		replacement string
	}{
		{"'日😀 ''本'' # 字'", "日本😀"},
		{"\"a\\\"b\"", "true"},
		{"plain#value", "a: b # c"},
		{"|-\n      first\n      second", "line1\nline2"},
		{"first\n      second", "new"},
	}
	for _, tc := range cases {
		t.Run(tc.value, func(t *testing.T) {
			source := strings.Replace(base, "'Submit'", tc.value, 1)
			result := applyOK(t, source, Operation{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: tc.replacement})
			if *result.Document.Controls[0].Caption != tc.replacement {
				t.Fatal("wrong string")
			}
			if !strings.Contains(string(result.Source), "    left: 100\n    top: 20\n  # unrelated") {
				t.Fatalf("neighbor consumed:\n%s", result.Source)
			}
		})
	}
	// yaml.v3 columns count runes; a flow scalar after emoji must map to bytes.
	source := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform: {name: Main}\ncontrols: [{id: a, name: A, type: Label, caption: '😀日本', left: 1, top: 2}]\n"
	result := applyOK(t, source, Operation{Type: MoveControl, ControlID: "a", Left: new(3.0), Top: new(2.0)})
	if string(result.Source) != strings.Replace(source, "left: 1", "left: 3", 1) {
		t.Fatalf("Unicode offset:\n%s", result.Source)
	}
}

func TestFormExplicitPaths(t *testing.T) {
	result := applyOK(t, base,
		Operation{Type: SetFormProperty, Field: "caption", Value: "legacy"},
		Operation{Type: SetFormProperty, Field: "build.caption", Value: "authoritative"},
		Operation{Type: SetFormProperty, Field: "build.clientWidth", Value: 320},
		Operation{Type: SetFormProperty, Field: "build.clientHeight", Value: 240},
	)
	if *result.Document.Form.Caption != "legacy" || *result.Document.Form.Build.Caption != "authoritative" || *result.Document.Form.Build.ClientWidth != 320 {
		t.Fatal("explicit paths remapped")
	}
	if !strings.Contains(string(result.Source), "controls:\n  # button") {
		t.Fatal("controls changed")
	}
}

func TestAddRemoveAndEmptySequences(t *testing.T) {
	for _, source := range []string{base, strings.TrimSuffix(base, "\n"), "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform: {name: Main}\ncontrols: [] # empty\n", strings.ReplaceAll(base, "  -", "-")} {
		t.Run(source, func(t *testing.T) {
			// The indentless case must also shift the item's property lines.
			if strings.Contains(source, "\n- id:") {
				source = strings.ReplaceAll(source, "\n    ", "\n  ")
			}
			control := spec.FormSpecControl{ID: "new", Name: "New", Type: "Label", Caption: new("New caption")}
			result := applyOK(t, source, Operation{Type: AddControl, Control: &control})
			if len(result.Document.Controls) == 0 || result.Document.Controls[len(result.Document.Controls)-1].ID != "new" {
				t.Fatal("not added")
			}
			removed := applyOK(t, string(result.Source), Operation{Type: RemoveControl, ControlID: "new"})
			if slices.ContainsFunc(removed.Document.Controls, func(c spec.FormSpecControl) bool { return c.ID == "new" }) {
				t.Fatal("not removed")
			}
		})
	}
	result := applyOK(t, base, Operation{Type: RemoveControl, ControlID: "submit"}, Operation{Type: RemoveControl, ControlID: "other"})
	if len(result.Document.Controls) != 0 || !strings.Contains(string(result.Source), "warnings: [] # tail") {
		t.Fatalf("empty sequence:\n%s", result.Source)
	}
}

const hierarchy = "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: Main\ncontrols:\n  - id: frame\n    name: Frame1\n    type: Frame\n    controls:\n      - id: child\n        name: Child\n        type: TextBox\n        left: 5\n        top: 8\n  - id: other\n    name: Other\n    type: Frame\n"

func TestParentAndCascade(t *testing.T) {
	assertFailure(t, hierarchy, []Operation{{Type: RemoveControl, ControlID: "frame"}}, "UFE007")
	removed := applyOK(t, hierarchy, Operation{Type: RemoveControl, ControlID: "frame", Cascade: true})
	if len(removed.Document.Controls) != 1 || removed.Document.Controls[0].ID != "other" {
		t.Fatalf("cascade: %+v", removed.Document.Controls)
	}
	result := applyOK(t, hierarchy, Operation{Type: SetParent, ControlID: "child", ParentID: "other", Left: new(12.5)})
	child := result.Document.Controls[slices.IndexFunc(result.Document.Controls, func(c spec.FormSpecControl) bool { return c.ID == "child" })]
	if child.ParentID != "other" || *child.Left != 12.5 || *child.Top != 8 || *child.ZIndex != 0 {
		t.Fatalf("reparent: %+v", child)
	}
	root := applyOK(t, string(result.Source), Operation{Type: SetParent, ControlID: "child"})
	if root.Document.Controls[len(root.Document.Controls)-1].ParentID != "" {
		t.Fatal("not unparented")
	}
	assertFailure(t, hierarchy, []Operation{{Type: SetParent, ControlID: "frame", ParentID: "child"}}, "UFV010")
}

func TestReorderUsesSiblingZOrder(t *testing.T) {
	source := strings.Replace(base, "    left: 100", "    zIndex: 5\n    tabIndex: 9\n    left: 100", 1)
	result := applyOK(t, source, Operation{Type: ReorderControl, ControlID: "submit", Index: new(0)})
	if *result.Document.Controls[0].ZIndex != 0 || *result.Document.Controls[1].ZIndex != 1 || *result.Document.Controls[0].TabIndex != 9 {
		t.Fatal("wrong reorder semantics")
	}
	again := applyOK(t, string(result.Source), Operation{Type: ReorderControl, ControlID: "submit", Index: new(0)})
	if len(again.Edits) > 0 {
		t.Fatal("reorder churn")
	}
}

func TestExplicitIDsAndNoop(t *testing.T) {
	missingID := strings.Replace(base, "  - id: submit\n    name: Submit", "  - name: Submit", 1)
	assertFailure(t, missingID, []Operation{{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: "changed"}}, "UFV004")
	noop := applyOK(t, base, Operation{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: "Submit"})
	if len(noop.Edits) > 0 {
		t.Fatal("no-op changed source")
	}
	result := applyOK(t, base, Operation{Type: SetControlProperty, ControlID: "submit", Field: "name", Value: "Renamed"})
	if result.Document.Controls[0].ID != "submit" {
		t.Fatal("rename lost identity")
	}
	assertFailure(t, base, []Operation{{Type: AddControl, Control: &spec.FormSpecControl{ID: "other", Name: "Added", Type: "Label"}}}, "UFV007")
}

func TestValidationAndAtomicity(t *testing.T) {
	cases := []struct {
		op   Operation
		code string
	}{
		{Operation{Type: SetParent, ControlID: "submit", ParentID: "missing"}, "UFV008"},
		{Operation{Type: SetControlProperty, ControlID: "submit", Field: "enabled", Value: "wrong-type"}, "UFV002"},
		{Operation{Type: SetControlProperty, ControlID: "submit", Field: "observed.caption", Value: "x"}, "UFE005"},
		{Operation{Type: MoveControl, ControlID: "missing", Left: new(1.0), Top: new(2.0)}, "UFE004"},
		{Operation{Type: ReorderControl, ControlID: "submit", Index: new(9)}, "UFE003"},
		{Operation{Type: "unknown"}, "UFE003"},
		{Operation{Type: SetControlProperty, ControlID: "submit", Field: "Caption", Value: "wrong-key"}, "UFE005"},
		{Operation{Type: SetFormProperty, Field: "build.ClientWidth", Value: 200}, "UFE005"},
		{Operation{Type: MoveControl, ControlID: "submit", Left: new(1.0)}, "UFE003"},
		{Operation{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: "new", Left: new(1.0)}, "UFE003"},
	}
	for _, tc := range cases {
		t.Run(tc.code+string(tc.op.Type), func(t *testing.T) {
			err := assertFailure(t, base, []Operation{{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: "changed"}, tc.op}, tc.code)
			if err[0].OperationIndex != 1 {
				t.Fatal("incorrect operation attribution")
			}
		})
	}
	assertFailure(t, base, []Operation{{Type: SetFormProperty, Field: "build.clientWidth", Value: 200}, {Type: SetFormProperty, Field: "width", Value: 300}}, "UFV017")
	// A temporary missing parent is allowed when the same transaction creates it.
	result := applyOK(t, base, Operation{Type: SetParent, ControlID: "submit", ParentID: "new"}, Operation{Type: AddControl, Control: &spec.FormSpecControl{ID: "new", Name: "New", Type: "Frame"}})
	if result.Document.Controls[0].ParentID != "new" {
		t.Fatal("transaction did not validate final state")
	}
	structured := assertFailure(t, base, []Operation{{Type: SetParent, ControlID: "submit", ParentID: "missing"}, {Type: SetControlProperty, ControlID: "other", Field: "text", Value: "updated"}}, "UFV008")
	if structured[0].OperationIndex != 0 || structured[0].ControlID != "submit" {
		t.Fatalf("wrong final-validation attribution: %+v", structured)
	}
}

func TestFlowStructuralAndAncillarySyntax(t *testing.T) {
	source := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform: {name: Main} # form\ncontrols: [{id: a, name: A, type: Label, caption: 'x'}] # controls\n"
	result := applyOK(t, source, Operation{Type: SetControlProperty, ControlID: "a", Field: "visible", Value: false}, Operation{Type: AddControl, Control: &spec.FormSpecControl{ID: "b", Name: "B", Type: "TextBox"}})
	if !strings.Contains(string(result.Source), "form: {name: Main} # form") || !strings.Contains(string(result.Source), "# controls") {
		t.Fatal("unrelated source changed")
	}
	if strings.Count(string(result.Source), "# controls") != 1 {
		t.Fatalf("duplicated comment:\n%s", result.Source)
	}
	applyOK(t, string(result.Source), Operation{Type: RemoveControl, ControlID: "a"})
	anchored := strings.Replace(base, "caption: 'Submit'", "caption: &caption 'Submit'", 1)
	assertFailure(t, anchored, []Operation{{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: "changed"}}, "UFE006")
	applyOK(t, anchored, Operation{Type: SetControlProperty, ControlID: "other", Field: "text", Value: "changed"})
	neighborAnchor := strings.Replace(base, "name: Submit", "name: &keep Submit", 1)
	neighborAnchor = strings.Replace(neighborAnchor, "text: keep", "text: *keep", 1)
	result = applyOK(t, neighborAnchor, Operation{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: "updated"})
	if string(result.Source) != strings.Replace(neighborAnchor, "caption: 'Submit'", "caption: 'updated'", 1) {
		t.Fatalf("unrelated anchor changed:\n%s", result.Source)
	}
	inserted := applyOK(t, neighborAnchor, Operation{Type: SetControlProperty, ControlID: "submit", Field: "visible", Value: false})
	if !strings.Contains(string(inserted.Source), "name: &keep Submit") || !strings.Contains(string(inserted.Source), "text: *keep") {
		t.Fatal("property insertion changed unrelated anchors")
	}
	assertFailure(t, neighborAnchor, []Operation{{Type: SetControlProperty, ControlID: "submit", Field: "name", Value: "Renamed"}}, "UFE006")
}

func TestFormattingBoundaries(t *testing.T) {
	for _, number := range []string{"100.0", "1e2", "0x64"} {
		source := strings.Replace(base, "left: 100", "left: "+number, 1)
		result := applyOK(t, source, Operation{Type: MoveControl, ControlID: "submit", Left: new(100.0), Top: new(20.0)})
		if len(result.Edits) != 0 || string(result.Source) != source {
			t.Fatalf("numeric no-op churn: %+v", result.Edits)
		}
	}
	result := applyOK(t, base, Operation{Type: RemoveControl, ControlID: "submit"})
	if strings.Contains(string(result.Source), "# button") || !strings.Contains(string(result.Source), "# unrelated") {
		t.Fatalf("item comment ownership:\n%s", result.Source)
	}
	anchored := strings.Replace(base, "controls:", "controls: &items", 1)
	assertFailure(t, anchored, []Operation{{Type: MoveControl, ControlID: "submit", Left: new(3.0), Top: new(20.0)}}, "UFE006")
	for _, value := range []string{"\n", "\n\n", " ", "\r\n", "\u0085", "\u2028", "\u2029"} {
		result := applyOK(t, base, Operation{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: value})
		if *result.Document.Controls[0].Caption != value {
			t.Fatalf("lost whitespace string %q", value)
		}
	}
	source := strings.Replace(base, "caption: 'Submit' # inline", "caption: |- # inline\n      first\n      second", 1)
	result = applyOK(t, source, Operation{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: "changed"})
	if !strings.Contains(string(result.Source), "caption: changed # inline") {
		t.Fatalf("lost scalar comment:\n%s", result.Source)
	}
}

func TestEditedFormSpecCompilesAndProjects(t *testing.T) {
	result := applyOK(t, hierarchy,
		Operation{Type: SetFormProperty, Field: "build.caption", Value: "Edited"},
		Operation{Type: MoveControl, ControlID: "child", Left: new(24.0), Top: new(12.0)},
		Operation{Type: SetControlProperty, ControlID: "child", Field: "text", Value: "Sentinel"},
		Operation{Type: AddControl, Control: &spec.FormSpecControl{ID: "label", Name: "Label1", Type: "Label", ParentID: "frame", Caption: new("Label")}},
	)
	generated, err := compiler.CompileNew(result.Document, 1252)
	if err != nil {
		t.Fatal(err)
	}
	projected, err := projection.Project(generated)
	if err != nil {
		t.Fatal(err)
	}
	index := slices.IndexFunc(projected.Controls, func(c spec.FormSpecControl) bool { return c.Name == "Child" })
	if index < 0 || projected.Controls[index].Text == nil || *projected.Controls[index].Text != "Sentinel" {
		t.Fatalf("generation lost edited value: %+v", projected.Controls)
	}
	if projected.Form.Build == nil || *projected.Form.Build.Caption != "Edited" {
		t.Fatal("generation lost caption")
	}
}

func TestPageAndTabStripAuthority(t *testing.T) {
	source := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform: {name: Main}\ncontrols:\n  - {id: multi, name: Multi, type: MultiPage, selectedIndex: 0}\n  - {id: page, name: Page1, type: Page, parentId: multi}\n  - {id: tabs, name: Tabs, type: TabStrip, tabs: [{name: Tab1}]}\n"
	assertFailure(t, source, []Operation{{Type: MoveControl, ControlID: "page", Left: new(1.0), Top: new(2.0)}}, "")
	assertFailure(t, source, []Operation{{Type: SetParent, ControlID: "multi", ParentID: "tabs"}}, "UFV011")
	twoPages := strings.Replace(source, "selectedIndex: 0", "selectedIndex: 1", 1) + "  - {id: page2, name: Page2, type: Page, parentId: multi}\n"
	assertFailure(t, twoPages, []Operation{{Type: RemoveControl, ControlID: "page"}}, "UFV020")
	result := applyOK(t, twoPages, Operation{Type: RemoveControl, ControlID: "page"}, Operation{Type: SetControlProperty, ControlID: "multi", Field: "selectedIndex", Value: 0})
	if len(result.Document.Controls) != 3 {
		t.Fatal("atomic Page deletion failed")
	}
}

func TestWriterValidationRemainsSeparate(t *testing.T) {
	// FormSpec accepts names that a binary writer may reject. The edit engine
	// follows that same boundary instead of adding a second validation ruleset.
	result := applyOK(t, base, Operation{Type: SetControlProperty, ControlID: "submit", Field: "name", Value: "Other"})
	if _, err := compiler.CompileNew(result.Document, 1252); err == nil {
		t.Fatal("writer unexpectedly accepted duplicate names")
	}
}

func TestSerializableOperationsAndJSONGeometry(t *testing.T) {
	var op Operation
	if err := json.Unmarshal([]byte(`{"type":"moveControl","controlId":"submit","left":120,"top":180}`), &op); err != nil {
		t.Fatal(err)
	}
	result := applyOK(t, base, op)
	if *result.Document.Controls[0].Left != 120 {
		t.Fatal("protocol operation failed")
	}
	jsonBody, _ := spec.MarshalSnapshot("json", result.Document)
	op.Left, op.Top = new(140.0), new(190.0)
	res, err := Apply(spec.SpecInput{Format: "json"}, jsonBody, []Operation{op})
	if err != nil {
		t.Fatalf("JSON geometry edit: %v", err)
	}
	if *res.Document.Controls[0].Left != 140 || *res.Document.Controls[0].Top != 190 || len(res.Edits) != 2 {
		t.Fatalf("JSON geometry result: document=%+v edits=%+v", res.Document.Controls[0], res.Edits)
	}
	res, err = Apply(spec.SpecInput{Format: "json"}, jsonBody, []Operation{{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: "x"}})
	if err == nil || !reflect.DeepEqual(res, Result{}) {
		t.Fatal("unsupported JSON operation was not rejected atomically")
	}
}

func TestJSONGeometryPreservesBytesAndInsertsLocally(t *testing.T) {
	source := []byte("{\r\n  \"schemaVersion\":1,\r\n  \"kind\":\"xlflow.userform\",\r\n  \"basis\":\"designer\",\r\n  \"form\":{\"name\":\"日本😀\",\"width\":240,\"height\":180},\r\n  \"controls\":[{\"id\":\"submit\",\"type\":\"CommandButton\",\"name\":\"送信😀\",\"left\":1,\"top\":2}]\r\n}\r\n")
	result, err := Apply(spec.SpecInput{Format: "json"}, source, []Operation{{Type: ResizeControl, ControlID: "submit", Width: new(80.0), Height: new(30.0)}})
	if err != nil {
		t.Fatal(err)
	}
	want := strings.Replace(string(source), `"top":2}`, `"top":2,"width":80,"height":30}`, 1)
	if string(result.Source) != want {
		t.Fatalf("unexpected source diff\n got: %s\nwant: %s", result.Source, want)
	}
	if len(result.Edits) != 1 || result.Edits[0].Start != result.Edits[0].End {
		t.Fatalf("expected one local insertion, got %+v", result.Edits)
	}
	if !bytes.Equal(source, []byte(strings.Replace(want, `,"width":80,"height":30`, "", 1))) {
		t.Fatal("source bytes were mutated")
	}
}

func TestJSONGeometryInsertionPreservesPrettyJSONStyle(t *testing.T) {
	source := []byte("{\r\n  \"schemaVersion\": 1,\r\n  \"kind\": \"xlflow.userform\",\r\n  \"basis\": \"designer\",\r\n  \"form\": {\"name\": \"\u65e5\u672c\U0001F600\", \"width\": 240, \"height\": 180},\r\n  \"controls\": [\r\n    {\r\n      \"id\": \"submit\",\r\n      \"type\": \"CommandButton\",\r\n      \"name\": \"\u9001\u4fe1\U0001F600\",\r\n      \"left\": 1,\r\n      \"top\": 2\r\n    }\r\n  ]\r\n}\r\n")
	result, err := Apply(spec.SpecInput{Format: "json"}, source, []Operation{{Type: ResizeControl, ControlID: "submit", Width: new(80.0), Height: new(30.0)}})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Edits) != 1 {
		t.Fatalf("expected one grouped insertion, got %+v", result.Edits)
	}
	edit := result.Edits[0]
	wantInsertion := ",\r\n      \"width\": 80,\r\n      \"height\": 30"
	if edit.Text != wantInsertion {
		t.Fatalf("insertion formatting = %q, want %q", edit.Text, wantInsertion)
	}
	want := string(source[:edit.Start]) + edit.Text + string(source[edit.End:])
	if string(result.Source) != want {
		t.Fatalf("source changed outside the insertion\n got: %s\nwant: %s", result.Source, want)
	}
}

func TestJSONGeometryRejectsAmbiguousAndInvalidInputAtomically(t *testing.T) {
	source := []byte(`{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"submit","type":"CommandButton","name":"Submit","left":1,"left":2,"top":3}]}`)
	op := Operation{Type: MoveControl, ControlID: "submit", Left: new(10.0), Top: new(20.0)}
	if result, err := Apply(spec.SpecInput{Format: "json"}, source, []Operation{op}); err == nil || !reflect.DeepEqual(result, Result{}) {
		t.Fatal("duplicate target key was accepted")
	}
	valid := []byte(strings.Replace(string(source), `,"left":2`, "", 1))
	op.Left = new(math.Inf(1))
	if result, err := Apply(spec.SpecInput{Format: "json"}, valid, []Operation{op}); err == nil || !reflect.DeepEqual(result, Result{}) {
		t.Fatal("non-finite geometry was accepted")
	}
	op.Left = new(10.0)
	missing := Operation{Type: MoveControl, ControlID: "submit", Left: new(10.0)}
	if result, err := Apply(spec.SpecInput{Format: "json"}, valid, []Operation{op, missing}); err == nil || !reflect.DeepEqual(result, Result{}) {
		t.Fatal("partial geometry batch was accepted")
	}
}

func TestJSONGeometryNumericNoopAndMoveResizeBatch(t *testing.T) {
	source := []byte(`{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"submit","type":"CommandButton","name":"Submit","left":1e1,"top":2,"width":10,"height":8}]}`)
	noop, err := Apply(spec.SpecInput{Format: "json"}, source, []Operation{{Type: MoveControl, ControlID: "submit", Left: new(10.0), Top: new(2.0)}})
	if err != nil {
		t.Fatal(err)
	}
	if len(noop.Edits) != 0 || !bytes.Equal(noop.Source, source) {
		t.Fatalf("numeric semantic no-op rewrote source: edits=%+v source=%s", noop.Edits, noop.Source)
	}
	result, err := Apply(spec.SpecInput{Format: "json"}, source, []Operation{
		{Type: MoveControl, ControlID: "submit", Left: new(12.0), Top: new(4.0)},
		{Type: ResizeControl, ControlID: "submit", Width: new(30.0), Height: new(16.0)},
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Edits) != 4 || *result.Document.Controls[0].Left != 12 || *result.Document.Controls[0].Top != 4 || *result.Document.Controls[0].Width != 30 || *result.Document.Controls[0].Height != 16 {
		t.Fatalf("move+resize batch was not applied atomically: edits=%+v control=%+v", result.Edits, result.Document.Controls[0])
	}
	if rebuilt := applySourceEdits(source, result.Edits); !bytes.Equal(rebuilt, result.Source) {
		t.Fatalf("reported JSON edits do not produce result source: %s", rebuilt)
	}
}

func TestJSONGeometryResolvesNestedControlByExplicitID(t *testing.T) {
	source := []byte(`{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"frame","type":"Frame","name":"Frame1","controls":[{"id":"child","type":"TextBox","name":"Child","left":1,"top":2}]}]}`)
	result, err := Apply(spec.SpecInput{Format: "json"}, source, []Operation{{Type: MoveControl, ControlID: "child", Left: new(5.0), Top: new(6.0)}})
	if err != nil {
		t.Fatal(err)
	}
	var child *spec.FormSpecControl
	for i := range result.Document.Controls {
		if result.Document.Controls[i].ID == "child" {
			child = &result.Document.Controls[i]
		}
	}
	if child == nil || *child.Left != 5 || *child.Top != 6 || len(result.Edits) != 2 {
		t.Fatalf("nested explicit ID was not edited: control=%+v edits=%+v", child, result.Edits)
	}
}

func TestJSONGeometryReportsEscapedDuplicateKey(t *testing.T) {
	source := []byte(`{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"submit","\u0069d":"other","type":"CommandButton","name":"Submit","left":1,"top":2}]}`)
	result, err := Apply(spec.SpecInput{Format: "json"}, source, []Operation{{Type: MoveControl, ControlID: "other", Left: new(5.0), Top: new(6.0)}})
	editError, ok := errors.AsType[*Error](err)
	if err == nil || !ok || !reflect.DeepEqual(result, Result{}) || len(editError.Diagnostics) == 0 || editError.Diagnostics[0].Code != "UFE008" {
		t.Fatalf("escaped duplicate key result=%+v error=%#v", result, err)
	}
}

func TestJSONGeometryUsesFinalCanonicalPageValidation(t *testing.T) {
	source := []byte(`{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"multi","type":"MultiPage","name":"Pages","controls":[{"id":"page","type":"Page","name":"Page1"}]}]}`)
	result, err := Apply(spec.SpecInput{Format: "json"}, source, []Operation{{Type: MoveControl, ControlID: "page", Left: new(5.0), Top: new(6.0)}})
	editError, ok := errors.AsType[*Error](err)
	if err == nil || !ok || !reflect.DeepEqual(result, Result{}) {
		t.Fatalf("Page geometry result=%+v error=%#v", result, err)
	}
	for _, diagnostic := range editError.Diagnostics {
		if diagnostic.Code == "UFV005" {
			return
		}
	}
	t.Fatalf("canonical Page geometry diagnostic missing: %+v", editError.Diagnostics)
}

func FuzzScalarSourceEdits(f *testing.F) {
	for _, seed := range []string{"Submit", "true", "a: b # c", "日本😀", "\"quoted\"", "line1\nline2"} {
		f.Add(seed)
	}
	f.Fuzz(func(t *testing.T, value string) {
		if !utf8.ValidString(value) {
			t.Skip()
		}
		result, err := Apply(spec.SpecInput{Format: "yaml"}, []byte(base), []Operation{{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: value}})
		if err != nil {
			t.Fatalf("valid UTF-8 scalar edit rejected: %q: %v", value, err)
		}
		if *result.Document.Controls[0].Caption != value {
			t.Fatalf("value changed: %q -> %q\n%s", value, *result.Document.Controls[0].Caption, result.Source)
		}
		if !strings.Contains(string(result.Source), "  # unrelated\n  - id: other") {
			t.Fatal("unrelated source changed")
		}
	})
}
