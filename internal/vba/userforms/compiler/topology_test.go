package compiler

import (
	"errors"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func TestCompileNewGeneratesNestedFramesAndEmptySibling(t *testing.T) {
	input := newSpec()
	input.CoordinateSystem = "parent-relative"
	input.Controls = []spec.FormSpecControl{
		{ID: "outer", Name: "Outer", Type: "Frame", ZIndex: new(0)},
		{ID: "inner", ParentID: "outer", Name: "Inner", Type: "Frame", Enabled: new(false), ZIndex: new(0)},
		{ID: "nested-label", ParentID: "inner", Name: "NestedLabel", Type: "Label", Left: new(72.0), Top: new(18.0), ZIndex: new(0)},
		{ID: "empty", Name: "Empty", Type: "Frame", ZIndex: new(1)},
	}
	form, err := CompileNew(input, 932)
	if err != nil {
		t.Fatal(err)
	}
	projected, err := projection.Project(form)
	if err != nil {
		t.Fatal(err)
	}
	if len(projected.Controls) != 4 || projected.Controls[0].Name != "Outer" || projected.Controls[1].Name != "Inner" || projected.Controls[2].Name != "NestedLabel" || projected.Controls[3].Name != "Empty" {
		t.Fatalf("generated preorder = %#v", projected.Controls)
	}
	if projected.Controls[1].ParentID != projected.Controls[0].ID || projected.Controls[2].ParentID != projected.Controls[1].ID || projected.Controls[3].ParentID != "" {
		t.Fatalf("generated parent relationships = %#v", projected.Controls)
	}
	if projected.Controls[2].Left == nil || *projected.Controls[2].Left != 72 || projected.Controls[2].Top == nil || *projected.Controls[2].Top != 18 {
		t.Fatalf("nested points were not stored parent-relative: %#v", projected.Controls[2])
	}

	outer := findControl(form.Controls, "Outer")
	inner := findControl(form.Controls, "Inner")
	empty := findControl(form.Controls, "Empty")
	if outer == nil || inner == nil || empty == nil || len(outer.Children) != 1 || len(inner.Children) != 1 || len(empty.Children) != 0 {
		t.Fatal("nested/empty Frame ownership differs")
	}
	if got := outer.Record.Sizes["DisplayedSize"]; got != (oforms.Size{Width: 5080, Height: 3810}) {
		t.Fatalf("default Frame size = %+v", got)
	}
	if outer.Record.Values["BooleanProperties"] != 0x8004 || inner.Record.Values["BooleanProperties"] != 0x8000 {
		t.Fatalf("Frame enabled flags = %#x, %#x", outer.Record.Values["BooleanProperties"], inner.Record.Values["BooleanProperties"])
	}
	for _, frame := range []*oforms.Control{outer, inner, empty} {
		if _, exists := frame.Record.Values["VariousPropertyBits"]; exists {
			t.Fatalf("Frame %q persisted VariousPropertyBits", frame.Name)
		}
	}
}

func TestCompileNewNestedFrameTypeMatchingFollowsContract(t *testing.T) {
	input := newSpec()
	input.Controls = []spec.FormSpecControl{
		{ID: "frame", Name: "FrameMain", Type: "fRaMe"},
		{ID: "child", ParentID: "frame", Name: "Child", Type: "Label"},
	}
	form, err := CompileNew(input, 932)
	if err != nil {
		t.Fatal(err)
	}
	if frame := findControl(form.Controls, "FrameMain"); frame == nil || len(frame.Children) != 1 {
		t.Fatal("case-insensitive Frame contract did not generate its child")
	}
}

func topologyBase(t *testing.T) (*oforms.Form, spec.FormSpec) {
	t.Helper()
	input := newSpec()
	input.CoordinateSystem = "parent-relative"
	input.Controls = []spec.FormSpecControl{
		{ID: "left", Name: "LeftFrame", Type: "Frame", ZIndex: new(0)},
		{ID: "input", ParentID: "left", Name: "Input", Type: "TextBox", Value: "keep", ZIndex: new(0)},
		{ID: "right", Name: "RightFrame", Type: "Frame", ZIndex: new(1)},
		{ID: "replace", Name: "ReplaceMe", Type: "Label", Caption: new("old"), ZIndex: new(2)},
		{ID: "remove", Name: "RemoveMe", Type: "Label", ZIndex: new(3)},
		{ID: "keep", Name: "Keep", Type: "Label", ZIndex: new(4)},
	}
	form, err := CompileNew(input, 932)
	if err != nil {
		t.Fatal(err)
	}
	before, err := projection.Project(form)
	if err != nil {
		t.Fatal(err)
	}
	return form, before
}

func TestCompileEditsAppliesTopologyChangesAtomically(t *testing.T) {
	base, before := topologyBase(t)
	original, err := oforms.SerializeForm(base, 932)
	if err != nil {
		t.Fatal(err)
	}
	beforeInput := snapshotCopy(t, before)
	after := snapshotCopy(t, before)
	controls := make([]spec.FormSpecControl, 0, len(after.Controls))
	for _, control := range after.Controls {
		if control.Name != "LeftFrame" && control.Name != "RemoveMe" {
			controls = append(controls, control)
		}
	}
	after.Controls = controls
	byName := make(map[string]*spec.FormSpecControl, len(after.Controls))
	for i := range after.Controls {
		byName[after.Controls[i].Name] = &after.Controls[i]
	}
	byName["Input"].ParentID = byName["RightFrame"].ID
	byName["ReplaceMe"].Type = "TextBox"
	byName["ReplaceMe"].ProgID = ""
	byName["ReplaceMe"].Caption = nil
	byName["ReplaceMe"].Properties = nil
	byName["ReplaceMe"].Value = "replacement"
	byName["Keep"].ZIndex = new(0)
	byName["RightFrame"].ZIndex = new(1)
	byName["ReplaceMe"].ZIndex = new(2)
	after.Controls = append(after.Controls,
		spec.FormSpecControl{ID: "added-frame", Name: "AddedFrame", Type: "Frame", ZIndex: new(3)},
		spec.FormSpecControl{ID: "added-label", ParentID: "added-frame", Name: "AddedLabel", Type: "Label", Caption: new("new"), ZIndex: new(0)},
	)
	result, err := CompileEdits(base, before, after, 932)
	if err != nil {
		t.Fatal(err)
	}
	if got := []string{result.Controls[0].Name, result.Controls[1].Name, result.Controls[2].Name, result.Controls[3].Name}; !reflect.DeepEqual(got, []string{"Keep", "RightFrame", "ReplaceMe", "AddedFrame"}) {
		t.Fatalf("root sibling order = %v", got)
	}
	if findControl(result.Controls, "LeftFrame") != nil || findControl(result.Controls, "RemoveMe") != nil {
		t.Fatal("deleted controls remain")
	}
	right := findControl(result.Controls, "RightFrame")
	input := findControl(result.Controls, "Input")
	if right == nil || len(right.Children) != 1 || right.Children[0] != input {
		t.Fatal("Input was not reparented under RightFrame")
	}
	addedFrame := findControl(result.Controls, "AddedFrame")
	if addedFrame == nil || len(addedFrame.Children) != 1 || addedFrame.Children[0].Name != "AddedLabel" {
		t.Fatal("new nested Frame subtree was not created")
	}
	oldInput, oldReplacement := findControl(base.Controls, "Input"), findControl(base.Controls, "ReplaceMe")
	newReplacement := findControl(result.Controls, "ReplaceMe")
	if input.ID != oldInput.ID || !reflect.DeepEqual(input.Record, oldInput.Record) || !reflect.DeepEqual(input.Site.Raw, oldInput.Site.Raw) {
		t.Fatal("reparented retained control lost its identity or raw payload")
	}
	if oldReplacement == nil || newReplacement == nil || newReplacement.ID == oldReplacement.ID || newReplacement.CLSIDCacheIndex == oldReplacement.CLSIDCacheIndex {
		t.Fatal("type replacement retained the old control identity")
	}
	maxOldID := int32(0)
	for _, control := range flattenControls(base.Controls) {
		maxOldID = max(maxOldID, control.ID)
	}
	for _, name := range []string{"ReplaceMe", "AddedFrame", "AddedLabel"} {
		control := findControl(result.Controls, name)
		if control == nil || control.ID <= maxOldID {
			t.Fatalf("new control %q did not receive a globally monotonic ID", name)
		}
	}
	unchanged, err := oforms.SerializeForm(base, 932)
	if err != nil || !reflect.DeepEqual(unchanged, original) {
		t.Fatal("successful topology compilation mutated its input Form")
	}
	if !reflect.DeepEqual(snapshotCopy(t, before), beforeInput) {
		t.Fatal("successful topology compilation mutated before snapshot")
	}
}

func TestCompileEditsRejectsStaleAndInvalidTopologyWithoutMutation(t *testing.T) {
	for _, tc := range []struct {
		name   string
		code   string
		change func(*spec.FormSpec, *spec.FormSpec)
	}{
		{
			name: "stale-before-tree",
			code: Stale,
			change: func(before, after *spec.FormSpec) {
				before.Controls[0].Name = "NotThePersistedFrame"
				after.Controls[0].Caption = new("requested")
			},
		},
		{
			name: "unknown-parent",
			code: Invalid,
			change: func(_, after *spec.FormSpec) {
				after.Controls[1].ParentID = "missing-parent"
			},
		},
	} {
		t.Run(tc.name, func(t *testing.T) {
			base, before := topologyBase(t)
			original, err := oforms.SerializeForm(base, 932)
			if err != nil {
				t.Fatal(err)
			}
			beforeOriginal := snapshotCopy(t, before)
			after := snapshotCopy(t, before)
			tc.change(&before, &after)
			beforeInput, afterInput := snapshotCopy(t, before), snapshotCopy(t, after)
			result, err := CompileEdits(base, before, after, 932)
			detail, ok := errors.AsType[*Error](err)
			if result != nil || !ok || detail.Code != tc.code {
				t.Fatalf("result=%v error=%v; want %s", result, err, tc.code)
			}
			unchanged, err := oforms.SerializeForm(base, 932)
			if err != nil || !reflect.DeepEqual(unchanged, original) {
				t.Fatal("failed topology compilation mutated its input Form")
			}
			if !reflect.DeepEqual(snapshotCopy(t, before), beforeInput) || !reflect.DeepEqual(snapshotCopy(t, after), afterInput) {
				t.Fatal("failed topology compilation mutated its snapshots")
			}
			if tc.name == "unknown-parent" && !reflect.DeepEqual(beforeOriginal, beforeInput) {
				t.Fatal("invalid after topology changed before snapshot")
			}
		})
	}
}

func TestCompileTemplatePreservesAndRejectsMultiPageTopologyChanges(t *testing.T) {
	base, snapshot := fixture(t, "p6_nested_form.bin")
	original, err := oforms.SerializeForm(base, 932)
	if err != nil {
		t.Fatal(err)
	}
	unchanged, err := CompileTemplate(base, templateInput(t, base, snapshot), 932)
	if err != nil {
		t.Fatal(err)
	}
	unchangedBytes, err := oforms.SerializeForm(unchanged, 932)
	if err != nil {
		t.Fatal(err)
	}
	assertDesignerBytes(t, original, unchangedBytes)

	desired := templateInput(t, base, snapshot)
	multiPageIndex := -1
	for i := range desired.Controls {
		if desired.Controls[i].Type == "MultiPage" {
			multiPageIndex = i
			break
		}
	}
	if multiPageIndex < 0 {
		t.Fatal("nested fixture has no MultiPage")
	}
	byID := make(map[string]spec.FormSpecControl, len(desired.Controls))
	for _, control := range desired.Controls {
		byID[control.ID] = control
	}
	frameID := ""
	for _, candidate := range desired.Controls {
		if candidate.Type != "Frame" || candidate.ID == desired.Controls[multiPageIndex].ID || candidate.ID == desired.Controls[multiPageIndex].ParentID {
			continue
		}
		insideMultiPage := false
		for parentID := candidate.ParentID; parentID != ""; parentID = byID[parentID].ParentID {
			if parentID == desired.Controls[multiPageIndex].ID {
				insideMultiPage = true
				break
			}
		}
		if !insideMultiPage {
			frameID = candidate.ID
			break
		}
	}
	if frameID == "" {
		t.Fatal("nested fixture has no Frame outside the MultiPage subtree")
	}
	desired.Controls[multiPageIndex].ParentID = frameID
	desiredInput := snapshotCopy(t, desired)
	result, err := CompileTemplate(base, desired, 932)
	detail, ok := errors.AsType[*Error](err)
	if result != nil || !ok || detail.Code != Unsupported {
		t.Fatalf("MultiPage reparent result=%v error=%v; want %s", result, err, Unsupported)
	}
	if !reflect.DeepEqual(desiredInput, snapshotCopy(t, desired)) {
		t.Fatal("rejected MultiPage mutation changed template input")
	}
	current, err := oforms.SerializeForm(base, 932)
	if err != nil || !reflect.DeepEqual(current, original) {
		t.Fatal("rejected MultiPage mutation changed the base Form")
	}
}
