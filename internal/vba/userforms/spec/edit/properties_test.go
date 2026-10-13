package edit

import (
	"bytes"
	"encoding/json"
	"errors"
	"math"
	"reflect"
	"slices"
	"strconv"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

const propertyJSON = "{\r\n  \"schemaVersion\":1,\r\n  \"kind\":\"xlflow.userform\",\r\n  \"basis\":\"designer\",\r\n  \"form\":{\"name\":\"Main\",\"caption\":null},\r\n  \"controls\":[{\"id\":\"submit\",\"name\":\"Submit\",\"type\":\"CommandButton\",\"caption\":\"\\u65e5本😀\",\"left\":1e2,\"enabled\":false},{\"id\":\"other\",\"name\":\"Other\",\"type\":\"TextBox\",\"text\":\"\"}],\r\n  \"warnings\":[]\r\n}\r\n"

func applyFormatOK(t *testing.T, format, source string, operations ...Operation) Result {
	t.Helper()
	input := []byte(source)
	original := bytes.Clone(input)
	result, err := Apply(spec.SpecInput{Format: format}, input, operations)
	if err != nil {
		t.Fatalf("Apply %s: %v (%#v)", format, err, err)
	}
	if !bytes.Equal(input, original) {
		t.Fatal("input mutated")
	}
	if !bytes.Equal(applySourceEdits(input, result.Edits), result.Source) {
		t.Fatalf("edits do not rebuild result: %+v", result.Edits)
	}
	parsed, err := spec.ParseFormSpec(spec.SpecInput{Format: format}, result.Source)
	if err != nil || !reflect.DeepEqual(parsed, result.Document) {
		t.Fatalf("noncanonical result: %v", err)
	}
	return result
}

func TestPropertyEditsRejectNumericConversionLoss(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		source := base
		if format == "json" {
			source = propertyJSON
		}
		for _, number := range []string{"0.1000000000000000001", "18446744073709551616", "1e-400"} {
			input := []byte(source)
			original := bytes.Clone(input)
			result, err := Apply(spec.SpecInput{Format: format}, input, []Operation{
				{Type: SetFormProperty, Field: "caption", Value: "Temporary"},
				{Type: SetControlProperty, ControlID: "other", Field: "value", Value: json.Number(number)},
			})
			structured, ok := errors.AsType[*Error](err)
			if !ok || len(structured.Diagnostics) == 0 || structured.Diagnostics[0].Code != "UFE003" || structured.Diagnostics[0].OperationIndex != 1 || structured.Diagnostics[0].ControlID != "other" || structured.Diagnostics[0].Field != "value" || !reflect.DeepEqual(result, Result{}) || !bytes.Equal(input, original) {
				t.Fatalf("rounded payload %s/%s: %+v %v", format, number, result, err)
			}
		}
		for _, number := range []string{"0.1", "1.00", "1e2", "9007199254740993", "9223372036854775807"} {
			result := applyFormatOK(t, format, source, Operation{Type: SetControlProperty, ControlID: "other", Field: "value", Value: json.Number(number)})
			var value any
			if format == "yaml" {
				engine, err := newEngine(result.Source)
				if err != nil {
					t.Fatal(err)
				}
				if err := field(engine.controls["other"].node, "value").Decode(&value); err != nil {
					t.Fatal(err)
				}
				// Compare shortest decimal, rather than exact binary float, for fractional values.
				encoded, err := json.Marshal(value)
				if err != nil {
					t.Fatal(err)
				}
				value = json.Number(encoded)
			} else {
				root, err := parseJSONSource(result.Source)
				if err != nil {
					t.Fatal(err)
				}
				nodes := map[string][]*jsonSourceNode{}
				indexJSONControls(jsonSourceMemberValue(root, "controls"), nodes)
				node := jsonSourceMemberValue(nodes["other"][0], "value")
				value = json.Number(result.Source[node.start:node.end])
			}
			if rational(value).Cmp(rational(json.Number(number))) != 0 {
				t.Fatalf("accepted decimal changed: %s/%s %v", format, number, value)
			}
		}
	}
}

func TestCustomControlCommonPropertyEdits(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		source := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform: {name: Main}\ncontrols:\n  - {id: widget, name: Widget1, type: VendorWidget, progId: Vendor.Widget.1}\n"
		if format == "json" {
			source = `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"widget","name":"Widget1","type":"VendorWidget","progId":"Vendor.Widget.1"}]}`
		}
		result := applyFormatOK(t, format, source,
			Operation{Type: SetControlProperty, ControlID: "widget", Field: "name", Value: "Renamed"},
			Operation{Type: SetControlProperty, ControlID: "widget", Field: "enabled", Value: false},
		)
		grid, err := PropertyGrid(spec.SpecInput{Format: format}, result.Source)
		if err != nil || grid.Controls["widget"].Values["enabled"] != (PropertyState{Present: true, Value: false}) || grid.Controls["widget"].Values["name"] != (PropertyState{Present: true, Value: "Renamed"}) {
			t.Fatalf("custom common metadata: %+v %v", grid, err)
		}
		for _, name := range []string{"text", "selectedIndex", "caption", "properties", "progId", "observed", "id", "type"} {
			if _, ok := grid.Controls["widget"].Values[name]; ok {
				t.Fatalf("custom type-specific/non-authoring field exposed: %s", name)
			}
		}
		_, err = Apply(spec.SpecInput{Format: format}, []byte(source), []Operation{{Type: SetControlProperty, ControlID: "widget", Field: "text", Value: "Unsupported"}})
		if err == nil {
			t.Fatal("unsupported custom field accepted")
		}
	}
}

func TestPropertyGridNumericPrecision(t *testing.T) {
	for _, format := range []string{"json", "yaml"} {
		for _, number := range []string{"9007199254740993", "0.100000000000000000001", "1e-400", "0.1", "1.00", "1e2", "9007199254740992"} {
			t.Run(format+"/"+number, func(t *testing.T) {
				source := strings.Replace(propertyJSON, `"text":""`, `"text":"","value":`+number, 1)
				if format == "yaml" {
					source = strings.Replace(base, "text: keep", "text: keep\n    value: "+number, 1)
				}
				input := spec.SpecInput{Format: format}
				if _, err := spec.ParseFormSpec(input, []byte(source)); err != nil {
					t.Fatalf("canonical fixture: %v", err)
				}
				grid, err := PropertyGrid(input, []byte(source))
				unsafe := number == "9007199254740993" || number == "0.100000000000000000001" || number == "1e-400"
				if unsafe {
					structured, ok := errors.AsType[*Error](err)
					if !ok || len(structured.Diagnostics) == 0 || structured.Diagnostics[0].Field != "value" || !reflect.DeepEqual(grid, PropertyGridData{}) {
						t.Fatalf("precision loss not refused: %+v %v", grid, err)
					}
				} else {
					if err != nil {
						t.Fatal(err)
					}
					encoded, err := json.Marshal(grid.Controls["other"].Values["value"])
					if err != nil {
						t.Fatal(err)
					}
					var wire struct {
						Present bool
						Value   float64
					}
					if err := json.Unmarshal(encoded, &wire); err != nil {
						t.Fatal(err)
					}
					if !wire.Present || rational(json.Number(number)).Cmp(rational(json.Number(strconv.FormatFloat(wire.Value, 'g', -1, 64)))) != 0 {
						t.Fatalf("wire value changed: %s", encoded)
					}
				}
				// Metadata restrictions must not restrict canonical source edits.
				result := applyFormatOK(t, format, source, Operation{Type: SetFormProperty, Field: "caption", Value: "New"})
				if !bytes.Contains(result.Source, []byte(number)) {
					t.Fatal("unrelated numeric token changed")
				}
			})
		}
	}
}

func TestPropertyGridRejectsNonScalarAny(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		for _, value := range []string{"[1, 2]", "{nested: value}"} {
			source := strings.Replace(base, "text: keep", "text: keep\n    value: "+value, 1)
			if format == "json" {
				if strings.HasPrefix(value, "{") {
					value = `{"nested":"value"}`
				}
				source = strings.Replace(propertyJSON, "\"text\":\"\"", "\"text\":\"\",\"value\":"+value, 1)
			}
			input := spec.SpecInput{Format: format}
			if _, err := spec.ParseFormSpec(input, []byte(source)); err != nil {
				t.Fatalf("fixture must be canonical valid: %v", err)
			}
			grid, err := PropertyGrid(input, []byte(source))
			structured, ok := errors.AsType[*Error](err)
			if !ok || len(structured.Diagnostics) == 0 || structured.Diagnostics[0].Field != "value" || !reflect.DeepEqual(grid, PropertyGridData{}) {
				t.Fatalf("complex any leaked to grid %s: %+v %v", format, grid, err)
			}
		}
	}
}

func TestPropertyGridRejectsAliasAndMergeDependentYAML(t *testing.T) {
	fixtures := []string{
		strings.Replace(base, "caption: 'Submit'", "caption: &caption 'Submit'", 1) + "# anchored authoring\n",
		strings.Replace(strings.Replace(base, "caption: 'Submit'", "caption: &caption 'Submit'", 1), "text: keep", "text: *caption", 1),
		strings.Replace(base, "name: Main", "name: Main\n  width: &width 100\n  height: *width", 1),
		strings.Replace(strings.Replace(base, "type: CommandButton", "type: CommandButton\n    properties: &custom {vendor: value}", 1), "type: TextBox", "type: TextBox\n    properties:\n      <<: *custom", 1),
	}
	for _, source := range fixtures {
		input := spec.SpecInput{Format: "yaml"}
		if _, err := spec.ParseFormSpec(input, []byte(source)); err != nil {
			t.Fatalf("fixture must be canonical valid: %v", err)
		}
		grid, err := PropertyGrid(input, []byte(source))
		structured, ok := errors.AsType[*Error](err)
		if !ok || len(structured.Diagnostics) == 0 || structured.Diagnostics[0].Code != "UFE006" || !reflect.DeepEqual(grid, PropertyGridData{}) {
			t.Fatalf("unsafe YAML shown as unset: %+v %v", grid, err)
		}
	}
}

func TestNullContainersAreNotGuessedDuringInsertion(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		source := strings.Replace(base, "name: Main", "name: Main\n  width: null", 1)
		if format == "json" {
			source = strings.Replace(propertyJSON, "\"caption\":null", "\"caption\":null,\"width\":null", 1)
		}
		input := spec.SpecInput{Format: format}
		if _, err := spec.ParseFormSpec(input, []byte(source)); err != nil {
			t.Fatal(err)
		}
		grid, err := PropertyGrid(input, []byte(source))
		if err != nil || !grid.Form.Values["width"].Present || grid.Form.Values["width"].Value != nil {
			t.Fatalf("null canonical width was not preserved: %+v %v", grid, err)
		}
		body := []byte(source)
		original := bytes.Clone(body)
		result, err := Apply(input, body, []Operation{{Type: SetFormProperty, Field: "width", Value: 200}})
		if err != nil || result.Document.Form.Width == nil || *result.Document.Form.Width != 200 || !bytes.Equal(body, original) {
			t.Fatalf("null canonical width update failed: %s %v", result.Source, err)
		}
		source = strings.Replace(base, "form:\n  name: Main", "form: null", 1)
		if format == "json" {
			source = strings.Replace(propertyJSON, "\"form\":{\"name\":\"Main\",\"caption\":null}", "\"form\":null", 1)
		}
		result, err = Apply(input, []byte(source), []Operation{{Type: SetFormProperty, Field: "caption", Value: "new"}})
		if err == nil || !reflect.DeepEqual(result, Result{}) {
			t.Fatalf("invalid null form overwritten: %s %v", result.Source, err)
		}
	}
}

func TestPropertyGridTabsUseOnlyAuthoredSource(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		t.Run(format, func(t *testing.T) {
			source := pageTabsYAML
			if format == "json" {
				source = pageTabsJSON
			}
			source = strings.ReplaceAll(source, "\n", "\r\n")
			grid, err := PropertyGrid(spec.SpecInput{Format: format}, []byte(source))
			if err != nil {
				t.Fatal(err)
			}
			absent, ok := grid.Tabs["absent"]
			if !ok || absent == nil || len(absent) != 0 {
				t.Fatalf("absent authored tabs must be an empty list despite observed tabs: %#v", grid.Tabs)
			}
			authored := grid.Tabs["strip"]
			if len(authored) != 2 {
				t.Fatalf("authored tabs missing: %#v", authored)
			}
			wantCaption := any("Alpha # keep this comment")
			if format == "json" {
				wantCaption = "Bravo"
			}
			if authored[0].Values["caption"] != (PropertyState{Present: true, Value: wantCaption}) {
				t.Fatalf("authored tab value was not read from source: %+v", authored[0])
			}
			if !slices.ContainsFunc(authored[0].Descriptors, func(descriptor PropertyDescriptor) bool { return descriptor.Field == "caption" }) {
				t.Fatalf("tab metadata omitted caption descriptor: %+v", authored[0].Descriptors)
			}
			encoded, err := json.Marshal(grid.Tabs)
			if err != nil || !bytes.Contains(encoded, []byte(`"absent":[]`)) || bytes.Contains(encoded, []byte("ObservedOnly")) {
				t.Fatalf("wire tabs should contain authored arrays only: %s (%v)", encoded, err)
			}
		})
	}
}

func TestSetTabPropertyPreservesExplicitNull(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		t.Run(format, func(t *testing.T) {
			source := pageTabsYAML
			if format == "json" {
				source = pageTabsJSON
			}
			result := applyFormatOK(t, format, source, Operation{
				Type: SetTabProperty, ControlID: "strip", Index: new(0), Field: "caption", Value: nil, ValuePresent: true,
			})
			grid, err := PropertyGrid(spec.SpecInput{Format: format}, result.Source)
			if err != nil {
				t.Fatal(err)
			}
			if grid.Tabs["strip"][0].Values["caption"] != (PropertyState{Present: true, Value: nil}) {
				t.Fatalf("explicit null was not retained: %+v", grid.Tabs["strip"][0].Values["caption"])
			}
			if format == "yaml" && !bytes.Contains(result.Source, []byte("caption: null")) {
				t.Fatalf("YAML null scalar was not authored: %s", result.Source)
			}
			if format == "json" && !bytes.Contains(result.Source, []byte(`"caption":null`)) {
				t.Fatalf("JSON null scalar was not authored: %s", result.Source)
			}
		})
	}
}

func TestJSONScalarPropertyPreservesTokenSpans(t *testing.T) {
	replacement := "新😀 \"quoted\"\n\\"
	result := applyFormatOK(t, "json", propertyJSON, Operation{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: replacement})
	encoded, _ := json.Marshal(replacement)
	expected := strings.Replace(propertyJSON, "\"\\u65e5本😀\"", string(encoded), 1)
	if string(result.Source) != expected || len(result.Edits) != 1 {
		t.Fatalf("nonlocal replacement: %s %+v", result.Source, result.Edits)
	}
	result = applyFormatOK(t, "json", propertyJSON,
		Operation{Type: SetFormProperty, Field: "width", Value: 200.0},
		Operation{Type: SetFormProperty, Field: "caption", Value: ""},
		Operation{Type: SetControlProperty, ControlID: "submit", Field: "tabIndex", Value: 0},
		Operation{Type: SetControlProperty, ControlID: "submit", Field: "visible", Value: false},
	)
	if result.Document.Form.Width == nil || *result.Document.Form.Width != 200 || result.Document.Form.Caption == nil || *result.Document.Form.Caption != "" {
		t.Fatalf("missing canonical form fields: %+v", result.Document.Form)
	}
	if !bytes.Contains(result.Source, []byte("\"caption\":\"\\u65e5本😀\"")) || !bytes.Contains(result.Source, []byte("\"left\":1e2")) || bytes.Count(result.Source, []byte("\r\n")) != strings.Count(propertyJSON, "\r\n") {
		t.Fatalf("unrelated bytes changed: %s", result.Source)
	}
}

func TestJSONScalarNoopAndRepeatedEdits(t *testing.T) {
	restored := applyFormatOK(t, "json", propertyJSON,
		Operation{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: "temporary"},
		Operation{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: "日本😀"},
		Operation{Type: SetControlProperty, ControlID: "submit", Field: "left", Value: 101},
		Operation{Type: SetControlProperty, ControlID: "submit", Field: "left", Value: 100},
	)
	if string(restored.Source) != propertyJSON || len(restored.Edits) != 0 {
		t.Fatalf("restored values lost original tokens: %s %+v", restored.Source, restored.Edits)
	}
	for _, op := range []Operation{
		{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: "日本😀"},
		{Type: SetControlProperty, ControlID: "submit", Field: "left", Value: 100.0},
		{Type: SetControlProperty, ControlID: "submit", Field: "enabled", Value: false},
		{Type: SetFormProperty, Field: "caption", ValuePresent: true},
	} {
		result := applyFormatOK(t, "json", propertyJSON, op)
		if len(result.Edits) != 0 || string(result.Source) != propertyJSON {
			t.Fatalf("semantic no-op rewrote source: %+v", result.Edits)
		}
	}
	result := applyFormatOK(t, "json", propertyJSON,
		Operation{Type: SetControlProperty, ControlID: "submit", Field: "tabIndex", Value: 1},
		Operation{Type: SetControlProperty, ControlID: "submit", Field: "tabIndex", Value: 2},
		Operation{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: "temp"},
		Operation{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: "日本😀"},
	)
	if *result.Document.Controls[0].TabIndex != 2 {
		t.Fatal("last edit did not win")
	}
}

func TestJSONCanonicalFormWidthInsertionPreservesLocalFormatting(t *testing.T) {
	cases := []struct {
		name      string
		source    string
		insertion bool
	}{
		{name: "insert", source: propertyJSON, insertion: true},
		{name: "replace", source: strings.Replace(propertyJSON, `"caption":null`, `"caption":null,"width":2.40e2`, 1)},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			result := applyFormatOK(t, "json", tc.source, Operation{Type: SetFormProperty, Field: "width", Value: 200.0})
			if result.Document.Form.Width == nil || *result.Document.Form.Width != 200 || len(result.Edits) != 1 || (result.Edits[0].Start == result.Edits[0].End) != tc.insertion {
				t.Fatalf("canonical form width edit = %+v, %+v", result.Document.Form, result.Edits)
			}
			if !bytes.Contains(result.Source, []byte(`"width":200`)) {
				t.Fatalf("width was not authored locally: %s", result.Source)
			}
		})
	}
}

func TestScalarAnyValuesStayDistinct(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		source := base
		if format == "json" {
			source = propertyJSON
		}
		for _, value := range []any{"false", false, "", 0, nil} {
			result := applyFormatOK(t, format, source, Operation{Type: SetControlProperty, ControlID: "other", Field: "value", Value: value, ValuePresent: true})
			grid, err := PropertyGrid(spec.SpecInput{Format: format}, result.Source)
			if err != nil {
				t.Fatal(err)
			}
			state := grid.Controls["other"].Values["value"]
			if !state.Present {
				t.Fatal("value presence lost")
			}
			if value == 0 {
				if rational(state.Value) == nil || rational(state.Value).Sign() != 0 {
					t.Fatalf("zero value lost: %+v", state)
				}
			} else if !reflect.DeepEqual(state.Value, value) {
				t.Fatalf("value type changed: %T %v -> %T %v", value, value, state.Value, state.Value)
			}
		}
	}
}

func TestPropertyGridNestedControlsUseOriginalSyntax(t *testing.T) {
	source := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"frame","name":"Frame1","type":"Frame","controls":[{"id":"child","name":"Text1","type":"TextBox","value":false}]}]}`
	result := applyFormatOK(t, "json", source, Operation{Type: SetControlProperty, ControlID: "child", Field: "text", Value: "nested"})
	if !bytes.Contains(result.Source, []byte(`"value":false,"text":"nested"`)) {
		t.Fatalf("nested source edit failed: %s", result.Source)
	}
	grid, err := PropertyGrid(spec.SpecInput{Format: "json"}, result.Source)
	if err != nil || len(grid.Controls) != 2 || grid.Controls["child"].Values["value"] != (PropertyState{Present: true, Value: false}) || grid.Controls["child"].Values["left"].Present {
		t.Fatalf("normalized values leaked: %+v %v", grid, err)
	}
}

func TestExplicitNullPropertyAndWirePresence(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		source := base
		if format == "json" {
			source = propertyJSON
		}
		result := applyFormatOK(t, format, source, Operation{Type: SetControlProperty, ControlID: "submit", Field: "caption", ValuePresent: true})
		grid, err := PropertyGrid(spec.SpecInput{Format: format}, result.Source)
		if err != nil {
			t.Fatal(err)
		}
		if state := grid.Controls["submit"].Values["caption"]; !state.Present || state.Value != nil {
			t.Fatalf("null lost: %+v", state)
		}
		result = applyFormatOK(t, format, source, Operation{Type: SetFormProperty, Field: "caption", ValuePresent: true})
		grid, err = PropertyGrid(spec.SpecInput{Format: format}, result.Source)
		if err != nil || !grid.Form.Values["caption"].Present || grid.Form.Values["caption"].Value != nil {
			t.Fatalf("form null lost: %+v %v", grid, err)
		}
	}
	for _, test := range []struct {
		source  string
		present bool
		value   any
	}{
		{`{"type":"setFormProperty","field":"caption"}`, false, nil},
		{`{"type":"setFormProperty","field":"caption","value":null}`, true, nil},
		{`{"type":"setFormProperty","field":"caption","value":""}`, true, ""},
		{`{"type":"setControlProperty","field":"enabled","value":false}`, true, false},
		{`{"type":"setControlProperty","field":"left","value":0}`, true, json.Number("0")},
	} {
		var op Operation
		if err := json.Unmarshal([]byte(test.source), &op); err != nil {
			t.Fatal(err)
		}
		if op.ValuePresent != test.present || !reflect.DeepEqual(op.Value, test.value) {
			t.Fatalf("presence/value lost: %+v", op)
		}
		encoded, err := json.Marshal(op)
		if err != nil {
			t.Fatal(err)
		}
		var roundtrip Operation
		if err := json.Unmarshal(encoded, &roundtrip); err != nil || !reflect.DeepEqual(op, roundtrip) {
			t.Fatalf("wire roundtrip: %s %+v %v", encoded, roundtrip, err)
		}
	}
	var op Operation
	_ = json.Unmarshal([]byte(`{"type":"setFormProperty","field":"caption","value":null}`), &op)
	_ = json.Unmarshal([]byte(`{"type":"setFormProperty","field":"caption"}`), &op)
	if op.ValuePresent {
		t.Fatal("unmarshal retained previous presence")
	}
}

func TestPropertyFailuresAreAtomic(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		source := base
		if format == "json" {
			source = propertyJSON
		}
		for _, op := range []Operation{
			{Type: SetFormProperty, Field: "caption"},
			{Type: SetFormProperty, Field: "name", ValuePresent: true},
			{Type: SetControlProperty, ControlID: "submit", Field: "name", ValuePresent: true},
			{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: []string{"x"}},
			{Type: SetControlProperty, ControlID: "submit", Field: "enabled", Value: "false"},
			{Type: SetControlProperty, ControlID: "submit", Field: "tabIndex", Value: 1.5},
			{Type: SetControlProperty, ControlID: "submit", Field: "width", Value: "invalid"},
			{Type: SetControlProperty, ControlID: "submit", Field: "left", Value: math.Inf(1)},
			{Type: SetControlProperty, ControlID: "submit", Field: "parentId", Value: "other"},
			{Type: SetControlProperty, ControlID: "submit", Field: "text", Value: "unsupported"},
			{Type: SetControlProperty, ControlID: "missing", Field: "caption", Value: "x"},
			{Type: SetFormProperty, Field: "build.Width", Value: 100},
			{Type: SetFormProperty, Field: "observed.caption", Value: "x"},
			{Type: MoveControl, ControlID: "submit", Left: new(1.0), Top: new(2.0), ValuePresent: true},
		} {
			input := []byte(source)
			original := bytes.Clone(input)
			result, err := Apply(spec.SpecInput{Format: format}, input, []Operation{{Type: SetControlProperty, ControlID: "submit", Field: "caption", Value: "changed"}, op})
			if err == nil || !reflect.DeepEqual(result, Result{}) || !bytes.Equal(input, original) {
				t.Fatalf("non-atomic failure %s %+v: %v", format, op, err)
			}
		}
	}
	for _, source := range []string{
		strings.Replace(propertyJSON, "\"caption\":null", "\"caption\":null,\"caption\":\"duplicate\"", 1),
		strings.Replace(propertyJSON, "\"name\":\"Main\"", "\"name\":\"Main\",\"width\":1,\"\\u0077idth\":2", 1),
	} {
		result, err := Apply(spec.SpecInput{Format: "json"}, []byte(source), []Operation{{Type: SetFormProperty, Field: "caption", Value: "x"}})
		if err == nil || !reflect.DeepEqual(result, Result{}) {
			t.Fatal("duplicate keys accepted")
		}
		if _, err := PropertyGrid(spec.SpecInput{Format: "json"}, []byte(source)); err == nil {
			t.Fatal("duplicate metadata accepted")
		}
	}
}

func TestEditablePropertiesMatchTypedAuthoring(t *testing.T) {
	form := EditableFormProperties()
	for _, name := range []string{"name", "caption", "width", "height"} {
		if _, ok := form[name]; !ok {
			t.Fatalf("missing form field %s", name)
		}
	}
	if len(form) != 4 {
		t.Fatalf("unexpected form properties: %+v", form)
	}
	for _, name := range []string{"id", "type", "parentId", "zIndex", "picture", "pictureAlignment", "pictureSizeMode", "controls", "properties", "observed", "list", "tabs"} {
		for _, typeName := range []string{"Image", "TextBox", "MultiPage", "Page", "TabStrip"} {
			if _, ok := EditableControlProperties(typeName)[name]; ok {
				t.Fatalf("noneditable %s.%s", typeName, name)
			}
		}
	}
	for _, name := range []string{"left", "top", "width", "height"} {
		if _, ok := EditableControlProperties("Page")[name]; ok {
			t.Fatalf("Page geometry exposed: %s", name)
		}
	}
	if len(EditableControlProperties("Unknown")) == 0 {
		t.Fatal("custom control lost common scalar properties")
	}
	props := EditableControlProperties("TextBox")
	delete(props, "text")
	if _, ok := EditableControlProperties("TextBox")["text"]; !ok {
		t.Fatal("contract map shared")
	}
}

func TestPropertyGridPreservesAuthoredStates(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		source := propertyJSON
		if format == "yaml" {
			source = strings.Replace(base, "name: Main", "name: Main\n  caption: null\n  observed: {caption: snapshot, width: 999}", 1)
			source = strings.Replace(source, "left: 100", "left: 0\n    enabled: false", 1)
			source = strings.Replace(source, "text: keep", "text: ''", 1)
		}
		grid, err := PropertyGrid(spec.SpecInput{Format: format}, []byte(source))
		if err != nil {
			t.Fatal(err)
		}
		if grid.Form.Values["caption"] != (PropertyState{Present: true}) || grid.Form.Values["width"].Present {
			t.Fatalf("normalization/snapshot leaked: %+v", grid.Form.Values)
		}
		if grid.Controls["submit"].Values["enabled"] != (PropertyState{Present: true, Value: false}) || grid.Controls["other"].Values["text"] != (PropertyState{Present: true, Value: ""}) {
			t.Fatalf("zero scalar lost: %+v", grid.Controls)
		}
		for _, target := range []PropertyTarget{grid.Form, grid.Controls["submit"], grid.Controls["other"]} {
			fields := []string{}
			for _, descriptor := range target.Descriptors {
				fields = append(fields, descriptor.Field)
				if descriptor.Nullable == descriptor.Required {
					t.Fatalf("invalid nullable: %+v", descriptor)
				}
			}
			if !slices.IsSorted(fields) {
				t.Fatal("descriptor order unstable")
			}
		}
		encoded, err := json.Marshal(grid)
		if err != nil || !bytes.Contains(encoded, []byte(`"present":false,"value":null`)) || !bytes.Contains(encoded, []byte(`"present":true,"value":null`)) {
			t.Fatalf("missing/null wire state: %s %v", encoded, err)
		}
	}
}
