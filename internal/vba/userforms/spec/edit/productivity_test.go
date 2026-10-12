package edit

import (
	"bytes"
	"encoding/json"
	"errors"
	"math"
	"reflect"
	"slices"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	"gopkg.in/yaml.v3"
)

const clipboardFixtureYAML = `schemaVersion: 1
kind: xlflow.userform
basis: designer
form:
  name: Main
  build: {clientWidth: 360, clientHeight: 240}
controls:
  - id: outer
    name: Outer
    type: Frame
    left: 0
    top: 0
    width: 320
    height: 200
  - id: frame
    name: Inner
    type: Frame
    parentId: outer
    left: null
    top: null
    width: 140
    height: 150
    observed: {left: 11.125, top: 7.25}
    controls:
      - id: picture
        name: Picture
        type: Image
        left: 1.123456789012345678901
        top: 2
        width: 48
        height: 24
        visible: false
        picture: {path: assets/picture.png}
      - id: multi
        name: Multi
        type: MultiPage
        left: 4
        top: 5
        width: 100
        height: 100
        controls:
          - id: page
            name: MainPage
            type: Page
            controls:
              - id: pageLabel
                name: PageLabel
                type: Label
                left: 3
                top: 4
                width: 30
                height: 12
                enabled: false
                tabIndex: 0
                caption: ""
                observed: {top: 4.5}
      - id: strip
        name: Strip
        type: TabStrip
        left: 8
        top: 9
        width: 90
        height: 24
        tabs:
          - name: Alpha
            caption: null
            enabled: false
      - id: caption
        name: Caption
        type: Label
        left: 10
        top: null
        width: 40
        height: 18
        observed: {top: 12.75}
warnings: []
`

func productivitySource(t *testing.T, format string) []byte {
	t.Helper()
	return sourceInFormat(t, format, clipboardFixtureYAML)
}

func sourceInFormat(t *testing.T, format, yamlSource string) []byte {
	t.Helper()
	if format != "json" {
		return []byte(yamlSource)
	}
	var document yaml.Node
	if err := yaml.Unmarshal([]byte(yamlSource), &document); err != nil {
		t.Fatal(err)
	}
	compact, err := yamlNodeJSON(document.Content[0])
	if err != nil {
		t.Fatal(err)
	}
	var pretty bytes.Buffer
	if err := json.Indent(&pretty, compact, "", "  "); err != nil {
		t.Fatal(err)
	}
	return pretty.Bytes()
}

func assertProductivityFailure(t *testing.T, format string, source []byte, operation Operation, code string) {
	t.Helper()
	original := bytes.Clone(source)
	result, err := Apply(spec.SpecInput{Format: format}, source, []Operation{operation})
	if err == nil {
		t.Fatalf("Apply %s %s unexpectedly succeeded: %+v", format, operation.Type, result)
	}
	editErr, ok := errors.AsType[*Error](err)
	if !ok || !slicesContainsDiagnostic(editErr.Diagnostics, code) {
		t.Fatalf("Apply %s %s error = %#v, want diagnostic %s", format, operation.Type, err, code)
	}
	if !reflect.DeepEqual(result, Result{}) || !bytes.Equal(source, original) {
		t.Fatalf("failed %s was not atomic: result=%+v", operation.Type, result)
	}
}

func slicesContainsDiagnostic(diagnostics []Diagnostic, code string) bool {
	for _, diagnostic := range diagnostics {
		if diagnostic.Code == code {
			return true
		}
	}
	return false
}

func clipboardFixtureAllocations() []ControlAllocation {
	return []ControlAllocation{
		{SourceID: "frame", ID: "frameCopy", Name: "FrameCopy"},
		{SourceID: "picture", ID: "pictureCopy", Name: "PictureCopy"},
		{SourceID: "multi", ID: "multiCopy", Name: "MultiCopy"},
		{SourceID: "page", ID: "pageCopy", Name: "PageCopy"},
		{SourceID: "pageLabel", ID: "pageLabelCopy", Name: "PageLabelCopy"},
		{SourceID: "strip", ID: "stripCopy", Name: "StripCopy"},
		{SourceID: "caption", ID: "captionCopy", Name: "CaptionCopy"},
	}
}

func TestProductivityMovePreservesObservedNullAndFractionalGeometry(t *testing.T) {
	const sourceYAML = `schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Main, build: {clientWidth: 220, clientHeight: 120}}
controls:
  - id: first
    name: First
    type: Label
    left: 1.23456789
    top: null
    width: 20
    height: 10
    observed: {top: 4.25}
  - id: second
    name: Second
    type: Label
    left: 10.5
    top: null
    width: 20
    height: 10
    observed: {top: 8.75}
`
	for _, format := range []string{"yaml", "yml", "json"} {
		t.Run(format, func(t *testing.T) {
			source := sourceInFormat(t, format, sourceYAML)
			result, err := Apply(spec.SpecInput{Format: format}, source, []Operation{{
				Type: MoveControls,
				Moves: []ControlMove{
					{ControlID: "first", Left: 2.23456789, Top: 4.25},
					{ControlID: "second", Left: 11.5, Top: 8.75},
				},
			}})
			if err != nil {
				t.Fatalf("Apply: %v", err)
			}
			if result.Document.Controls[0].Left == nil || *result.Document.Controls[0].Left != 2.23456789 {
				t.Fatalf("first left = %v", result.Document.Controls[0].Left)
			}
			if result.Document.Controls[0].Top != nil || result.Document.Controls[0].Observed == nil || result.Document.Controls[0].Observed.Top == nil || *result.Document.Controls[0].Observed.Top != 4.25 {
				t.Fatalf("untouched authored null/observed top changed: %+v", result.Document.Controls[0])
			}
			if result.Document.Controls[1].Top != nil || result.Document.Controls[1].Observed == nil || result.Document.Controls[1].Observed.Top == nil || *result.Document.Controls[1].Observed.Top != 8.75 {
				t.Fatalf("second authored null/observed top changed: %+v", result.Document.Controls[1])
			}
			if !bytes.Contains(result.Source, []byte("top: null")) && format != "json" {
				t.Fatalf("untouched null was materialized:\n%s", result.Source)
			}
			if !bytes.Contains(result.Source, []byte(`"top": null`)) && format == "json" {
				t.Fatalf("untouched null was materialized:\n%s", result.Source)
			}

			unchanged, err := Apply(spec.SpecInput{Format: format}, source, []Operation{{
				Type: MoveControls,
				Moves: []ControlMove{
					{ControlID: "first", Left: 1.23456829, Top: 4.25},
					{ControlID: "second", Left: 10.5000004, Top: 8.75},
				},
			}})
			if err != nil {
				t.Fatalf("no-op move: %v", err)
			}
			if !bytes.Equal(unchanged.Source, source) || len(unchanged.Edits) != 0 {
				t.Fatalf("fractional no-op changed source:\n%s", unchanged.Source)
			}
		})
	}
}

func TestDistributeUsesSpatialOrderAndKeepsEndpoints(t *testing.T) {
	source := `schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Main, build: {clientWidth: 200, clientHeight: 100}}
controls:
  - {id: zFront, name: ZFront, type: Label, zIndex: 10, left: 80, top: 10, width: 10, height: 10}
  - {id: spatialFirst, name: SpatialFirst, type: Label, zIndex: 20, left: 10, top: 10, width: 10, height: 10}
  - {id: middle, name: Middle, type: Label, zIndex: 0, left: 40, top: 10, width: 10, height: 10}
`
	result := applyFormatOK(t, "yaml", source, Operation{Type: DistributeControls, ControlIDs: []string{"zFront", "spatialFirst", "middle"}, Axis: "horizontal"})
	if *result.Document.Controls[0].Left != 80 || *result.Document.Controls[1].Left != 10 || *result.Document.Controls[2].Left != 45 {
		t.Fatalf("positions after distribution: %v, %v, %v", *result.Document.Controls[0].Left, *result.Document.Controls[1].Left, *result.Document.Controls[2].Left)
	}
	if !strings.Contains(string(result.Source), "zIndex: 10, left: 80") || !strings.Contains(string(result.Source), "zIndex: 20, left: 10") {
		t.Fatalf("distribution rewrote an endpoint:\n%s", result.Source)
	}
}

func TestSingleControlReorderAndEmptyTabOrder(t *testing.T) {
	const sourceYAML = `schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Main}
controls:
  - {id: first, name: First, type: Label, zIndex: 0, left: 1, top: 1, width: 10, height: 10}
  - {id: second, name: Second, type: Label, zIndex: 1, left: 20, top: 1, width: 10, height: 10}
  - {id: empty, name: Empty, type: Frame, zIndex: 2, left: 40, top: 1, width: 80, height: 60}
`
	for _, format := range []string{"yaml", "json"} {
		t.Run(format, func(t *testing.T) {
			source := sourceInFormat(t, format, sourceYAML)
			result := applyFormatOK(t, format, string(source), Operation{Type: ReorderControls, ControlIDs: []string{"first"}, Command: "front"})
			ordered := orderedChildren(result.Document.Controls, "", "")
			if ordered[len(ordered)-1].ID != "first" {
				t.Fatalf("single selected control was not moved to front: %+v", ordered)
			}
			noOp := applyFormatOK(t, format, string(source), Operation{Type: SetTabOrder, ParentID: "empty", ControlIDs: []string{}})
			if !bytes.Equal(noOp.Source, source) || len(noOp.Edits) != 0 {
				t.Fatalf("empty complete tab order should be a no-op: %+v", noOp.Edits)
			}
		})
	}
}

func TestAlignmentPreservesUntouchedFractionalAxis(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		t.Run(format, func(t *testing.T) {
			source := sourceInFormat(t, format, `schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Main, build: {clientWidth: 200, clientHeight: 120}}
controls:
  - {id: anchor, name: Anchor, type: Label, left: 1.23456789, top: 2.34567891, width: 20, height: 10}
  - {id: moving, name: Moving, type: Label, left: 10.55555, top: 3.78901234, width: 20, height: 10}
`)
			result, err := Apply(spec.SpecInput{Format: format}, source, []Operation{{
				Type: AlignControls, ControlIDs: []string{"anchor", "moving"}, AnchorID: "anchor", Alignment: "left",
			}})
			if err != nil {
				t.Fatal(err)
			}
			if math.Abs(*result.Document.Controls[1].Left-1.234568) > 1e-12 || *result.Document.Controls[1].Top != 3.78901234 {
				t.Fatalf("alignment changed wrong geometry: left=%v top=%v", *result.Document.Controls[1].Left, *result.Document.Controls[1].Top)
			}
			if format == "yaml" && !bytes.Contains(result.Source, []byte("top: 3.78901234")) {
				t.Fatalf("untouched fractional top was rewritten:\n%s", result.Source)
			}
			if format == "json" && !bytes.Contains(result.Source, []byte(`"top": 3.78901234`)) {
				t.Fatalf("untouched fractional top was rewritten:\n%s", result.Source)
			}
		})
	}
}

func TestSetTabOrderIncludesMetadataEligibleCustomControls(t *testing.T) {
	const sourceYAML = `schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Main}
controls:
  - {id: ordinary, name: Ordinary, type: Label, tabIndex: 7}
  - {id: custom, name: Custom, type: VendorWidget, progId: Vendor.Widget.1, tabIndex: null, visible: false, enabled: false}
  - {id: frame, name: Frame, type: Frame, tabIndex: 9}
  - {id: child, name: Child, type: VendorWidget, progId: Vendor.Widget.1, parentId: frame, tabIndex: 4}
`
	for _, format := range []string{"yaml", "json"} {
		t.Run(format, func(t *testing.T) {
			source := sourceInFormat(t, format, sourceYAML)
			input := spec.SpecInput{Format: format}
			metadata, err := PropertyGrid(input, source)
			if err != nil {
				t.Fatal(err)
			}
			for _, id := range []string{"custom", "child"} {
				if !slices.ContainsFunc(metadata.Controls[id].Descriptors, func(property PropertyDescriptor) bool { return property.Field == "tabIndex" }) {
					t.Fatalf("custom control %s lacks editable tabIndex metadata", id)
				}
			}
			assertProductivityFailure(t, format, source, Operation{Type: SetTabOrder, ControlIDs: []string{"ordinary", "frame"}}, "UFE003")
			assertProductivityFailure(t, format, source, Operation{Type: SetTabOrder, ControlIDs: []string{"custom", "ordinary", "child"}}, "UFE003")
			result := applyFormatOK(t, format, string(source), Operation{Type: SetTabOrder, ControlIDs: []string{"custom", "ordinary", "frame"}})
			for index, control := range result.Document.Controls {
				want := map[string]int{"custom": 0, "ordinary": 1, "frame": 2, "child": 4}[control.ID]
				if control.TabIndex == nil || *control.TabIndex != want {
					t.Fatalf("control %d (%s) tabIndex=%v, want %d", index, control.ID, control.TabIndex, want)
				}
			}
			childResult := applyFormatOK(t, format, string(source), Operation{Type: SetTabOrder, ParentID: "frame", ControlIDs: []string{"child"}})
			for _, control := range childResult.Document.Controls {
				if control.ID == "child" && (control.TabIndex == nil || *control.TabIndex != 0) {
					t.Fatalf("custom Frame child tab order was not applied: %+v", control)
				}
			}
		})
	}
}

func TestSetTabOrderIncludesHiddenDisabledAndRejectsMultiPageParent(t *testing.T) {
	source := `schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Main}
controls:
  - id: frame
    name: Frame
    type: Frame
    left: 0
    top: 0
    width: 180
    height: 100
    controls:
      - {id: first, name: First, type: CommandButton, left: 1, top: 1, width: 50, height: 20, tabIndex: 9, enabled: false, visible: false}
      - {id: second, name: Second, type: TextBox, left: 1, top: 30, width: 80, height: 18, tabIndex: 4}
  - id: multi
    name: Multi
    type: MultiPage
    left: 0
    top: 110
    width: 180
    height: 80
    selectedIndex: 0
    controls:
      - {id: page, name: Page, type: Page}
`
	result := applyFormatOK(t, "yaml", source, Operation{Type: SetTabOrder, ParentID: "frame", ControlIDs: []string{"second", "first"}})
	byID := make(map[string]spec.FormSpecControl, len(result.Document.Controls))
	for _, control := range result.Document.Controls {
		byID[control.ID] = control
	}
	if *byID["second"].TabIndex != 0 || *byID["first"].TabIndex != 1 || byID["first"].Visible == nil || *byID["first"].Visible {
		t.Fatalf("hidden/disabled eligible tab order not applied: %+v %+v", byID["first"], byID["second"])
	}
	assertProductivityFailure(t, "yaml", []byte(source), Operation{Type: SetTabOrder, ParentID: "multi", ControlIDs: []string{}}, "UFE005")
}

func TestAlignControlsAllDirectionsAndSameSizeDimensions(t *testing.T) {
	const sourceYAML = `schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Main, build: {clientWidth: 240, clientHeight: 180}}
controls:
  - {id: anchor, name: Anchor, type: Label, left: 40, top: 40, width: 50, height: 30}
  - {id: moving, name: Moving, type: Label, left: 10, top: 10, width: 20, height: 10}
`
	for _, format := range []string{"yaml", "json"} {
		t.Run(format, func(t *testing.T) {
			source := sourceInFormat(t, format, sourceYAML)
			alignments := []struct {
				name      string
				left, top float64
			}{
				{"left", 40, 10}, {"center", 55, 10}, {"right", 70, 10},
				{"top", 10, 40}, {"middle", 10, 50}, {"bottom", 10, 60},
			}
			for _, test := range alignments {
				t.Run("align_"+test.name, func(t *testing.T) {
					result := applyFormatOK(t, format, string(source), Operation{
						Type: AlignControls, ControlIDs: []string{"moving", "anchor"}, AnchorID: "anchor", Alignment: test.name,
					})
					moving := result.Document.Controls[1]
					if moving.Left == nil || moving.Top == nil || *moving.Left != test.left || *moving.Top != test.top {
						t.Fatalf("%s alignment geometry = (%v, %v), want (%v, %v)", test.name, moving.Left, moving.Top, test.left, test.top)
					}
				})
			}

			dimensions := []struct {
				name          string
				width, height float64
			}{
				{"width", 50, 10}, {"height", 20, 30}, {"size", 50, 30},
			}
			for _, test := range dimensions {
				t.Run("same_size_"+test.name, func(t *testing.T) {
					result := applyFormatOK(t, format, string(source), Operation{
						Type: SameSizeControls, ControlIDs: []string{"moving", "anchor"}, AnchorID: "anchor", Dimension: test.name,
					})
					moving := result.Document.Controls[1]
					if moving.Width == nil || moving.Height == nil || *moving.Width != test.width || *moving.Height != test.height {
						t.Fatalf("same-size %s geometry = (%v, %v), want (%v, %v)", test.name, moving.Width, moving.Height, test.width, test.height)
					}
				})
			}
		})
	}
}

func TestReorderControlsCommandsPreserveSelectedRelativeOrderInYAMLAndJSON(t *testing.T) {
	const sourceYAML = `schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Main}
controls:
  - {id: a, name: A, type: Label, zIndex: 0}
  - {id: b, name: B, type: Label, zIndex: 1}
  - {id: c, name: C, type: Label, zIndex: 2}
  - {id: d, name: D, type: Label, zIndex: 3}
  - {id: e, name: E, type: Label, zIndex: 4}
`
	expected := map[string][]string{
		"front":    {"a", "c", "e", "b", "d"},
		"forward":  {"a", "c", "b", "e", "d"},
		"backward": {"b", "a", "d", "c", "e"},
		"back":     {"b", "d", "a", "c", "e"},
	}
	for _, format := range []string{"yaml", "json"} {
		t.Run(format, func(t *testing.T) {
			for _, command := range []string{"front", "forward", "backward", "back"} {
				t.Run(command, func(t *testing.T) {
					source := sourceInFormat(t, format, sourceYAML)
					result := applyFormatOK(t, format, string(source), Operation{
						Type: ReorderControls, ControlIDs: []string{"d", "b"}, Command: command,
					})
					got := make([]string, len(result.Document.Controls))
					for index, control := range orderedChildren(result.Document.Controls, "", "") {
						got[index] = control.ID
					}
					if !reflect.DeepEqual(got, expected[command]) {
						t.Fatalf("%s sibling order = %v, want %v", command, got, expected[command])
					}
					selectedOrder := make([]string, 0, 2)
					for _, id := range got {
						if id == "b" || id == "d" {
							selectedOrder = append(selectedOrder, id)
						}
					}
					if !reflect.DeepEqual(selectedOrder, []string{"b", "d"}) {
						t.Fatalf("%s changed selected relative order: %v", command, selectedOrder)
					}
				})
			}
		})
	}
}

func TestRemoveControlsValidatesScopeAndAppliesCascadeAtomically(t *testing.T) {
	const sourceYAML = `schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Main}
controls:
  - id: frame
    name: Frame
    type: Frame
    left: 0
    top: 0
    width: 140
    height: 80
    controls:
      - {id: nested, name: Nested, type: Label, left: 1, top: 1, width: 20, height: 10}
  - {id: sibling, name: Sibling, type: Label, left: 160, top: 1, width: 20, height: 10}
`
	for _, format := range []string{"yaml", "json"} {
		t.Run(format, func(t *testing.T) {
			source := sourceInFormat(t, format, sourceYAML)
			assertProductivityFailure(t, format, source, Operation{
				Type: RemoveControls, ControlIDs: []string{"frame", "sibling"}, Cascade: false,
			}, "UFE007")
			result := applyFormatOK(t, format, string(source), Operation{
				Type: RemoveControls, ControlIDs: []string{"frame", "sibling"}, Cascade: true,
			})
			if len(result.Document.Controls) != 0 {
				t.Fatalf("cascade left controls behind: %+v", result.Document.Controls)
			}
			assertProductivityFailure(t, format, source, Operation{
				Type: RemoveControls, ControlIDs: []string{"frame", "nested"}, Cascade: true,
			}, "UFE005")
		})
	}
}

func TestProductivityRejectsCustomParentsAndDirectPageOrCustomTargets(t *testing.T) {
	source := []byte(`schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Main, build: {clientWidth: 240, clientHeight: 180}}
controls:
  - id: customFrame
    name: CustomFrame
    type: Frame
    progId: Vendor.Frame.1
    left: 0
    top: 0
    width: 180
    height: 120
    controls:
      - {id: first, name: First, type: Label, left: 1, top: 1, width: 20, height: 10}
      - {id: second, name: Second, type: Label, left: 30, top: 1, width: 20, height: 10}
  - {id: multi, name: Multi, type: MultiPage, left: 0, top: 130, width: 180, height: 80, selectedIndex: 0}
  - {id: page, name: Page, type: Page, parentId: multi}
  - {id: pageChild, name: PageChild, type: Label, parentId: page, left: 1, top: 1, width: 20, height: 10}
  - {id: custom, name: Custom, type: VendorWidget, progId: Vendor.Widget.1, left: 190, top: 1, width: 30, height: 20}
  - {id: ordinary, name: Ordinary, type: Label, left: 190, top: 30, width: 30, height: 20}
`)
	assertProductivityFailure(t, "yaml", source, Operation{
		Type: ReorderControls, ControlIDs: []string{"first", "second"}, Command: "front",
	}, "UFE005")
	assertProductivityFailure(t, "yaml", source, Operation{
		Type: RemoveControls, ControlIDs: []string{"page"}, Cascade: true,
	}, "UFE005")
	assertProductivityFailure(t, "yaml", source, Operation{
		Type: RemoveControls, ControlIDs: []string{"custom"}, Cascade: true,
	}, "UFE005")
	assertProductivityFailure(t, "yaml", source, Operation{
		Type: SetTabOrder, ParentID: "customFrame", ControlIDs: []string{"first", "second"},
	}, "UFE005")
	assertProductivityFailure(t, "yaml", source, Operation{
		Type: SetTabOrder, ParentID: "missing", ControlIDs: []string{},
	}, "UFE004")
}

func TestSetTabOrderRequiresCompleteListAndSupportsPageParent(t *testing.T) {
	const sourceYAML = `schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Main}
controls:
  - id: multi
    name: Multi
    type: MultiPage
    selectedIndex: 0
  - id: page
    name: Page
    type: Page
    parentId: multi
    controls:
      - {id: first, name: First, type: Label, tabIndex: 8}
      - {id: second, name: Second, type: TextBox, tabIndex: 3}
`
	source := []byte(sourceYAML)
	for _, ids := range [][]string{{"second"}, {"first", "first"}, {"first", "outside"}} {
		assertProductivityFailure(t, "yaml", source, Operation{Type: SetTabOrder, ParentID: "page", ControlIDs: ids}, "UFE003")
	}
	result := applyFormatOK(t, "yaml", sourceYAML, Operation{Type: SetTabOrder, ParentID: "page", ControlIDs: []string{"second", "first"}})
	byID := make(map[string]spec.FormSpecControl, len(result.Document.Controls))
	for _, control := range result.Document.Controls {
		byID[control.ID] = control
	}
	if *byID["second"].TabIndex != 0 || *byID["first"].TabIndex != 1 {
		t.Fatalf("Page children did not receive complete zero-based tab order: first=%v second=%v", byID["first"].TabIndex, byID["second"].TabIndex)
	}
}

func TestLayoutProductivityRejectsPageAndCustomDirectTargets(t *testing.T) {
	source := `schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Main}
controls:
  - {id: multi, name: Multi, type: MultiPage, left: 0, top: 0, width: 180, height: 100, selectedIndex: 0}
  - {id: pageA, name: PageA, type: Page, parentId: multi}
  - {id: pageB, name: PageB, type: Page, parentId: multi}
  - {id: custom, name: Custom, type: VendorWidget, progId: Vendor.Widget.1, left: 1, top: 1, width: 30, height: 20}
  - {id: ordinary, name: Ordinary, type: Label, left: 40, top: 1, width: 30, height: 20}
`
	assertFailure(t, source, []Operation{{Type: MoveControls, Moves: []ControlMove{{ControlID: "pageA", Left: 1, Top: 1}, {ControlID: "pageB", Left: 2, Top: 2}}}}, "UFE005")
	assertFailure(t, source, []Operation{{Type: AlignControls, ControlIDs: []string{"custom", "ordinary"}, AnchorID: "ordinary", Alignment: "left"}}, "UFE005")
}

func TestAlignUsesDesignerDefaultsAndRequiredContainerSize(t *testing.T) {
	t.Run("checkbox default width and command button height", func(t *testing.T) {
		source := `schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Main, build: {clientWidth: 240, clientHeight: 160}}
controls:
  - {id: option, name: Option, type: CheckBox, left: 1, top: 4}
  - {id: anchor, name: Anchor, type: CommandButton, left: 40, top: 40}
`
		result := applyFormatOK(t, "yaml", source, Operation{Type: AlignControls, ControlIDs: []string{"option", "anchor"}, AnchorID: "anchor", Alignment: "right"})
		if *result.Document.Controls[0].Left != 40 {
			t.Fatalf("CheckBox default width was not 72: left=%v", *result.Document.Controls[0].Left)
		}
		middle := applyFormatOK(t, "yaml", source, Operation{Type: AlignControls, ControlIDs: []string{"option", "anchor"}, AnchorID: "anchor", Alignment: "middle"})
		if *middle.Document.Controls[0].Top != 43 {
			t.Fatalf("CheckBox/CommandButton default heights did not yield centered top 43: %v", *middle.Document.Controls[0].Top)
		}
	})

	t.Run("hidden Frame child protects minimum during same-size", func(t *testing.T) {
		source := `schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Main, build: {clientWidth: 300, clientHeight: 200}}
controls:
  - id: anchor
    name: Anchor
    type: Frame
    left: 0
    top: 0
    width: 100
    height: 100
  - id: frame
    name: Frame
    type: Frame
    left: 130
    top: 0
    width: 140
    height: 100
    controls:
      - {id: hidden, name: Hidden, type: Label, left: 80, top: 10, width: 30, height: 15, visible: false}
`
		assertFailure(t, source, []Operation{{Type: SameSizeControls, ControlIDs: []string{"anchor", "frame"}, AnchorID: "anchor", Dimension: "width"}}, "UFE003")
	})

	t.Run("hidden Page child protects MultiPage minimum during alignment", func(t *testing.T) {
		source := `schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Main, build: {clientWidth: 320, clientHeight: 220}}
controls:
  - id: anchor
    name: Anchor
    type: MultiPage
    left: 0
    top: 0
    width: 160
    height: 120
    controls:
      - {id: anchorPage, name: AnchorPage, type: Page}
  - id: multi
    name: Multi
    type: MultiPage
    left: 80
    top: 10
    width: 100
    height: 80
    controls:
      - id: page
        name: Page
        type: Page
        controls:
          - {id: hidden, name: Hidden, type: Label, left: 80, top: 30, width: 30, height: 15, visible: false}
`
		assertFailure(t, source, []Operation{{Type: AlignControls, ControlIDs: []string{"anchor", "multi"}, AnchorID: "anchor", Alignment: "left"}}, "UFE003")
	})
}

func TestCopyControlsPreservesSourceSubtreeAndObservationState(t *testing.T) {
	for _, format := range []string{"yaml", "yml", "json"} {
		t.Run(format, func(t *testing.T) {
			snapshot, document, err := CopyControls(spec.SpecInput{Format: format}, productivitySource(t, format), []string{"frame"})
			if err != nil {
				t.Fatalf("CopyControls: %v", err)
			}
			if snapshot.Format != map[bool]string{true: "json", false: "yaml"}[format == "json"] || !reflect.DeepEqual(snapshot.ControlIDs, []string{"frame"}) {
				t.Fatalf("snapshot metadata: %+v", snapshot)
			}
			if len(document.Controls) != 7 {
				t.Fatalf("copied %d controls, want all 7 subtree controls: %+v", len(document.Controls), document.Controls)
			}
			byID := make(map[string]spec.FormSpecControl, len(document.Controls))
			for _, control := range document.Controls {
				byID[control.ID] = control
			}
			if _, exists := byID["outer"]; exists {
				t.Fatal("copy document included the external parent")
			}
			if byID["frame"].ParentID != "" || byID["page"].ParentID != "multi" || byID["pageLabel"].ParentID != "page" {
				t.Fatalf("copied topology is not subtree-local: %+v", byID)
			}
			if byID["frame"].Observed == nil || byID["frame"].Observed.Left == nil || *byID["frame"].Observed.Left != 11.125 || byID["frame"].Left != nil {
				t.Fatalf("authored null/observation was not preserved: %+v", byID["frame"])
			}
			if byID["pageLabel"].Caption == nil || *byID["pageLabel"].Caption != "" || byID["pageLabel"].Enabled == nil || *byID["pageLabel"].Enabled || byID["pageLabel"].TabIndex == nil || *byID["pageLabel"].TabIndex != 0 {
				t.Fatalf("empty/false/zero source values lost: %+v", byID["pageLabel"])
			}
			if byID["picture"].Picture == nil || byID["picture"].Picture.Path != "assets/picture.png" || byID["strip"].Tabs == nil || len(byID["strip"].Tabs) != 1 {
				t.Fatalf("picture or TabStrip data lost: picture=%+v tabs=%+v", byID["picture"].Picture, byID["strip"].Tabs)
			}
			if format == "json" && !strings.Contains(snapshot.Text, "1.123456789012345678901") {
				t.Fatalf("copy rounded an authored JSON number:\n%s", snapshot.Text)
			}
			copiedSource := []byte(snapshot.Text)
			if bytes.Contains(copiedSource, []byte("outer")) || bytes.Contains(copiedSource, []byte("Outer")) {
				t.Fatalf("snapshot text retained external parent data:\n%s", snapshot.Text)
			}
			parsed, err := spec.ParseFormSpec(spec.SpecInput{Format: snapshot.Format}, copiedSource)
			if err != nil || len(parsed.Controls) != len(document.Controls) {
				t.Fatalf("snapshot is not self-contained: controls=%d err=%v", len(parsed.Controls), err)
			}
		})
	}
}

func TestPasteControlsRewritesIDsParentsAndCoordinatesAtomically(t *testing.T) {
	for _, format := range []string{"yaml", "yml", "json"} {
		t.Run(format, func(t *testing.T) {
			input := spec.SpecInput{Format: format}
			source := productivitySource(t, format)
			parseInput := input
			if parseInput.Format == "yml" {
				parseInput.Format = "yaml"
			}
			before, err := spec.ParseFormSpec(parseInput, source)
			if err != nil {
				t.Fatal(err)
			}
			snapshot, copied, err := CopyControls(input, source, []string{"frame"})
			if err != nil {
				t.Fatal(err)
			}
			allocations := clipboardFixtureAllocations()
			result, err := Apply(input, source, []Operation{{
				Type:        PasteControls,
				Snapshot:    &snapshot,
				Allocations: allocations,
				ParentID:    "outer",
				Offset:      &GeometryOffset{Left: 10, Top: 5},
			}})
			if err != nil {
				t.Fatalf("paste: %v", err)
			}
			byID := make(map[string]spec.FormSpecControl, len(result.Document.Controls))
			for _, control := range result.Document.Controls {
				byID[control.ID] = control
			}
			if byID["frameCopy"].ParentID != "outer" || byID["frameCopy"].Left == nil || *byID["frameCopy"].Left != 21.125 || byID["frameCopy"].Top == nil || *byID["frameCopy"].Top != 12.25 {
				t.Fatalf("root parent/offset incorrect: %+v", byID["frameCopy"])
			}
			if byID["multiCopy"].ParentID != "frameCopy" || byID["pageCopy"].ParentID != "multiCopy" || byID["pageLabelCopy"].ParentID != "pageCopy" || byID["pictureCopy"].ParentID != "frameCopy" {
				t.Fatalf("internal parent references were not rewritten: %+v", byID)
			}
			if byID["pictureCopy"].Picture == nil || byID["pictureCopy"].Picture.Path != "assets/picture.png" || byID["stripCopy"].Tabs == nil || len(byID["stripCopy"].Tabs) != 1 {
				t.Fatalf("subtree payload lost: picture=%+v tabs=%+v", byID["pictureCopy"].Picture, byID["stripCopy"].Tabs)
			}
			if format == "json" && !bytes.Contains(result.Source, []byte("1.123456789012345678901")) {
				t.Fatalf("pasted JSON rounded an authored source number:\n%s", result.Source)
			}
			if len(result.Document.Controls) != len(before.Controls)+len(copied.Controls) {
				t.Fatalf("paste has %d controls, want %d", len(result.Document.Controls), len(before.Controls)+len(copied.Controls))
			}
			if !bytes.Equal(applySourceEdits(source, result.Edits), result.Source) {
				t.Fatal("paste source edits do not reproduce the result")
			}

			colliding := slicesClone(allocations)
			colliding[1].Name = "fRaMeCoPy"
			failed, err := Apply(input, source, []Operation{{
				Type: PasteControls, Snapshot: &snapshot, Allocations: colliding, ParentID: "outer", Offset: &GeometryOffset{},
			}})
			if err == nil || !reflect.DeepEqual(failed, Result{}) {
				t.Fatalf("case-insensitive in-paste name collision was not atomic: result=%+v err=%v", failed, err)
			}
		})
	}
}

func TestPasteControlsAcceptsCrossFormatSnapshotsAndClampsGroupBounds(t *testing.T) {
	for _, formats := range [][2]string{{"yaml", "json"}, {"json", "yaml"}} {
		t.Run(formats[0]+"_to_"+formats[1], func(t *testing.T) {
			sourceFormat, targetFormat := formats[0], formats[1]
			snapshot, _, err := CopyControls(spec.SpecInput{Format: sourceFormat}, productivitySource(t, sourceFormat), []string{"frame"})
			if err != nil {
				t.Fatal(err)
			}
			target := productivitySource(t, targetFormat)
			result, err := Apply(spec.SpecInput{Format: targetFormat}, target, []Operation{{
				Type: PasteControls, Snapshot: &snapshot, Allocations: clipboardFixtureAllocations(),
				Offset: &GeometryOffset{Left: 10000, Top: -10000},
			}})
			if err != nil {
				t.Fatalf("cross-format paste: %v", err)
			}
			frame := controlFromResult(t, result, "frameCopy")
			if frame.Left == nil || *frame.Left != 220 || frame.Top == nil || *frame.Top != 0 {
				t.Fatalf("paste did not clamp the whole copied root to parent bounds: left=%v top=%v", frame.Left, frame.Top)
			}
			if controlFromResult(t, result, "pageLabelCopy").ParentID != "pageCopy" || controlFromResult(t, result, "pictureCopy").Picture == nil {
				t.Fatalf("cross-format paste lost copied subtree: %+v", result.Document.Controls)
			}
			if !bytes.Equal(applySourceEdits(target, result.Edits), result.Source) {
				t.Fatal("cross-format paste source edits do not reproduce result")
			}
		})
	}
}

func TestPasteBoundaryClampPreservesUnchangedNullGeometryAxes(t *testing.T) {
	const sourceYAML = `schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Source, build: {clientWidth: 360, clientHeight: 240}}
controls:
  - {id: frame, name: Frame, type: Frame, left: null, top: null, width: 140, height: 150, observed: {left: 220, top: 7.25}}
`
	const targetYAML = `schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Target, build: {clientWidth: 360, clientHeight: 240}}
controls: []
`
	for _, format := range []string{"yaml", "json"} {
		t.Run(format, func(t *testing.T) {
			input := spec.SpecInput{Format: format}
			snapshot, _, err := CopyControls(input, sourceInFormat(t, format, sourceYAML), []string{"frame"})
			if err != nil {
				t.Fatal(err)
			}
			result, err := Apply(input, sourceInFormat(t, format, targetYAML), []Operation{{
				Type:        PasteControls,
				Snapshot:    &snapshot,
				Allocations: []ControlAllocation{{SourceID: "frame", ID: "pastedFrame", Name: "PastedFrame"}},
				Offset:      &GeometryOffset{Left: 10},
			}})
			if err != nil {
				t.Fatal(err)
			}
			pasted := controlFromResult(t, result, "pastedFrame")
			if pasted.Left != nil || pasted.Top != nil || pasted.Observed == nil || pasted.Observed.Left == nil || *pasted.Observed.Left != 220 || pasted.Observed.Top == nil || *pasted.Observed.Top != 7.25 {
				t.Fatalf("paste materialized an unchanged authored axis or lost observations: %+v", pasted)
			}
			nullToken := "left: null"
			if format == "json" {
				nullToken = `"left": null`
			}
			if !strings.Contains(string(result.Source), nullToken) {
				t.Fatalf("paste did not preserve authored null geometry in source:\n%s", result.Source)
			}
		})
	}
}

func TestPasteControlsRejectsCollidingOrIncompleteAllocationsAtomically(t *testing.T) {
	input := spec.SpecInput{Format: "yaml"}
	source := productivitySource(t, "yaml")
	snapshot, _, err := CopyControls(input, source, []string{"frame"})
	if err != nil {
		t.Fatal(err)
	}
	base := clipboardFixtureAllocations()
	for _, test := range []struct {
		name        string
		allocations []ControlAllocation
		code        string
	}{
		{name: "target ID collision", allocations: func() []ControlAllocation { a := slicesClone(base); a[0].ID = "outer"; return a }(), code: "UFE008"},
		{name: "target name collision ignoring case", allocations: func() []ControlAllocation { a := slicesClone(base); a[0].Name = "oUtEr"; return a }(), code: "UFE008"},
		{name: "duplicate allocated ID", allocations: func() []ControlAllocation { a := slicesClone(base); a[1].ID = a[0].ID; return a }(), code: "UFE008"},
		{name: "missing subtree allocation", allocations: slicesClone(base[:len(base)-1]), code: "UFE003"},
		{name: "unknown source allocation", allocations: func() []ControlAllocation { a := slicesClone(base); a[0].SourceID = "outside"; return a }(), code: "UFE008"},
	} {
		t.Run(test.name, func(t *testing.T) {
			assertProductivityFailure(t, "yaml", source, Operation{
				Type: PasteControls, Snapshot: &snapshot, Allocations: test.allocations,
				ParentID: "outer", Offset: &GeometryOffset{},
			}, test.code)
		})
	}
}

func TestPasteControlsRejectsUnfitDestinationParentAtomically(t *testing.T) {
	snapshot, _, err := CopyControls(spec.SpecInput{Format: "yaml"}, productivitySource(t, "yaml"), []string{"frame"})
	if err != nil {
		t.Fatal(err)
	}
	target := []byte(`schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Main, build: {clientWidth: 240, clientHeight: 180}}
controls:
  - {id: tiny, name: Tiny, type: Frame, left: 0, top: 0, width: 100, height: 100}
`)
	assertProductivityFailure(t, "yaml", target, Operation{
		Type: PasteControls, Snapshot: &snapshot, Allocations: clipboardFixtureAllocations(),
		ParentID: "tiny", Offset: &GeometryOffset{},
	}, "UFE003")
}

func controlFromResult(t *testing.T, result Result, id string) spec.FormSpecControl {
	t.Helper()
	for _, control := range result.Document.Controls {
		if control.ID == id {
			return control
		}
	}
	t.Fatalf("result has no control %q", id)
	return spec.FormSpecControl{}
}

func TestCopyAndPasteRootsFollowSiblingStackingOrder(t *testing.T) {
	source := []byte(`schemaVersion: 1
kind: xlflow.userform
basis: designer
form: {name: Main, build: {clientWidth: 200, clientHeight: 120}}
controls:
  - {id: a, name: A, type: Label, zIndex: 2, left: 10, top: 10, width: 20, height: 10}
  - {id: b, name: B, type: Label, zIndex: 8, left: 40, top: 10, width: 20, height: 10}
`)
	input := spec.SpecInput{Format: "yaml"}
	snapshot, copied, err := CopyControls(input, source, []string{"b", "a"})
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(snapshot.ControlIDs, []string{"a", "b"}) || len(copied.Controls) != 2 || copied.Controls[0].ID != "a" || copied.Controls[1].ID != "b" {
		t.Fatalf("copy order did not follow source sibling stacking: ids=%v controls=%+v", snapshot.ControlIDs, copied.Controls)
	}

	// The RPC boundary treats clipboard data as untrusted. Reversing root IDs
	// must not reverse the copied source stacking order during paste.
	snapshot.ControlIDs = []string{"b", "a"}
	result, err := Apply(input, source, []Operation{{
		Type:     PasteControls,
		Snapshot: &snapshot,
		Allocations: []ControlAllocation{
			{SourceID: "a", ID: "copyA", Name: "CopyA"},
			{SourceID: "b", ID: "copyB", Name: "CopyB"},
		},
		Offset: &GeometryOffset{Left: 0, Top: 0},
	}})
	if err != nil {
		t.Fatal(err)
	}
	ordered := orderedChildren(result.Document.Controls, "", "")
	if ordered[len(ordered)-2].ID != "copyA" || ordered[len(ordered)-1].ID != "copyB" {
		t.Fatalf("paste trusted untrusted controlIds order: %+v", ordered)
	}
}

func slicesClone[T any](values []T) []T {
	return append([]T(nil), values...)
}
