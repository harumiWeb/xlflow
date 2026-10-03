package oforms

import (
	"bytes"
	"errors"
	"fmt"
	"maps"
	"reflect"
	"testing"
)

func frameDefinition(name string, children ...ControlDefinition) ControlDefinition {
	return ControlDefinition{Name: name, Class: 14, Size: Size{5080, 3810}, Visible: true, Properties: map[string]any{"Caption": name, "BooleanProperties": int64(0x8004)}, Controls: children}
}

func multiPageDefinition(name string, pageNames ...string) ControlDefinition {
	pages := make([]ControlDefinition, len(pageNames))
	tabs := make([]Tab, len(pageNames))
	for i, pageName := range pageNames {
		pages[i] = ControlDefinition{Name: pageName, Class: 7, Visible: true}
		tabs[i] = Tab{Name: "Tab" + pageName, Caption: pageName, Enabled: true, Visible: true}
	}
	selected := int32(0)
	if len(pages) == 0 {
		selected = -1
	}
	return ControlDefinition{Name: name, Class: 57, Size: Size{5000, 3000}, Visible: true, Controls: pages, Tabs: &TabStrip{SelectedIndex: selected, Tabs: tabs}}
}

func TestNewFormNestedFrameOwnership(t *testing.T) {
	form, err := NewForm(Definition{Name: "Nested", Size: Size{8467, 6350}, Controls: []ControlDefinition{
		frameDefinition("Outer", frameDefinition("Inner", ControlDefinition{Name: "Text", Class: 23, Size: Size{1000, 600}, Position: Position{10, 20}, Visible: true, Properties: map[string]any{"Value": "日本語"}})), frameDefinition("Empty"),
	}}, 1252)
	if err != nil {
		t.Fatal(err)
	}
	if len(form.Levels) != 4 {
		t.Fatalf("levels=%d", len(form.Levels))
	}
	outer := form.Controls[0]
	inner := outer.Children[0]
	if outer.Level.Path != "Nested/i01" || inner.Level.Path != "Nested/i01/i02" || inner.Children[0].ID != 3 || form.Controls[1].ID != 4 {
		t.Fatal("incorrect storage/site identities")
	}
	if inner.Children[0].Site.Position.Left != 10 || inner.Children[0].Record.Strings["Value"].Text != "日本語" {
		t.Fatal("child persistence changed")
	}
	for _, level := range form.Levels {
		var count int
		for _, c := range level.Controls {
			count += int(c.ObjectStreamSize)
			if c.Level != nil && c.ObjectStreamSize != 0 {
				t.Fatal("container has inline object")
			}
		}
		if count != len(level.ORaw) {
			t.Fatal("object accounting")
		}
	}
	if _, err := SerializeForm(form, 1252); err != nil {
		t.Fatal(err)
	}
}

func topologyOf(form *Form) []TopologyControl {
	var out []TopologyControl
	unnamed := 0
	var walk func([]*Control, string)
	walk = func(cs []*Control, parent string) {
		for _, c := range cs {
			name := c.Name
			if name == "" {
				unnamed++
				name = fmt.Sprintf("<unnamed_%d>", unnamed)
			}
			out = append(out, TopologyControl{Name: name, Parent: parent})
			children := c.Children
			if c.MultiPage != nil {
				children = c.MultiPage.Pages
			}
			walk(children, name)
		}
	}
	walk(form.Controls, "")
	return out
}

func TestApplyTopologyNoOpPreservesNestedFixtures(t *testing.T) {
	for _, fixture := range []string{"p4_form.bin", "p6_nested_form.bin"} {
		t.Run(fixture, func(t *testing.T) {
			form, err := ReadForm(openFixture(t, fixture), "UserForm1", 932)
			if err != nil {
				t.Fatal(err)
			}
			before, err := SerializeForm(form, 932)
			if err != nil {
				t.Fatal(err)
			}
			after, err := ApplyTopology(form, topologyOf(form), 932)
			if err != nil {
				t.Fatal(err)
			}
			got, err := SerializeForm(after, 932)
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(before, got) {
				t.Fatal("no-op changed binary")
			}
		})
	}
}

func TestApplyTopologyMovesAddsDeletesAndReplacesAtomically(t *testing.T) {
	form, err := NewForm(Definition{Name: "Nested", Size: Size{8467, 6350}, Controls: []ControlDefinition{
		frameDefinition("Left", ControlDefinition{Name: "Text", Class: 23, Size: Size{1000, 600}, Visible: true, Properties: map[string]any{"Value": "keep"}}), frameDefinition("Right"), frameDefinition("Delete", frameDefinition("Gone")),
	}}, 1252)
	if err != nil {
		t.Fatal(err)
	}
	before, err := SerializeForm(form, 1252)
	if err != nil {
		t.Fatal(err)
	}
	text := form.Controls[0].Children[0]
	newFrame := frameDefinition("New")
	newLabel := ControlDefinition{Name: "Label", Class: 21, Size: Size{1000, 600}, Visible: true, Properties: map[string]any{"Caption": "new"}}
	got, err := ApplyTopology(form, []TopologyControl{{Name: "Right"}, {Name: "Text", Parent: "Right"}, {Name: "Left"}, {Name: "New", Parent: "Left", Definition: &newFrame}, {Name: "Label", Parent: "New", Definition: &newLabel}}, 1252)
	if err != nil {
		t.Fatal(err)
	}
	if got.Controls[0].Name != "Right" || got.Controls[0].Children[0].Name != "Text" || !bytes.Equal(text.Record.Raw, got.Controls[0].Children[0].Record.Raw) {
		t.Fatal("move changed payload or order")
	}
	stored, err := SerializeForm(got, 1252)
	if err != nil {
		t.Fatal(err)
	}
	for path := range stored.Storages {
		if path == "Nested/i04" || path == "Nested/i04/i05" {
			t.Fatal("orphan storage")
		}
	}
	unchanged, err := SerializeForm(form, 1252)
	if err != nil || !reflect.DeepEqual(before, unchanged) {
		t.Fatal("input changed")
	}
	bad := topologyOf(form)
	bad[0].Parent = "Text"
	if rejected, err := ApplyTopology(form, bad, 1252); rejected != nil || err == nil {
		t.Fatal("accepted invalid ownership")
	}
	unchanged, _ = SerializeForm(form, 1252)
	if !maps.EqualFunc(before.Streams, unchanged.Streams, bytes.Equal) {
		t.Fatal("failure changed input")
	}
	replacement := ControlDefinition{Name: "Right", Class: 21, Size: Size{1000, 600}, Visible: true}
	if _, err := ApplyTopology(got, []TopologyControl{{Name: "Right", Definition: &replacement}}, 1252); err != nil {
		t.Fatal(err)
	}
}

func TestApplyTopologyRejectsCyclesAndAcceptsMultiPageRemoval(t *testing.T) {
	form, err := NewForm(Definition{Name: "Nested", Size: Size{1000, 1000}, Controls: []ControlDefinition{frameDefinition("A", frameDefinition("B"))}}, 1252)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := ApplyTopology(form, []TopologyControl{{Name: "A", Parent: "B"}, {Name: "B", Parent: "A"}}, 1252); !errors.Is(err, ErrInvalidEdit) {
		t.Fatalf("cycle=%v", err)
	}
	base, err := NewForm(Definition{Name: "TopologyMulti", Size: Size{8467, 6350}, Controls: []ControlDefinition{
		multiPageDefinition("Pages", "PageA", "PageB"),
		frameDefinition("Keep"),
	}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	before, err := SerializeForm(base, 932)
	if err != nil {
		t.Fatal(err)
	}
	removed, err := ApplyTopology(base, []TopologyControl{{Name: "Keep"}}, 932)
	if err != nil {
		t.Fatalf("MultiPage removal: %v", err)
	}
	if len(removed.Controls) != 1 || removed.Controls[0].Name != "Keep" {
		t.Fatalf("controls after removal = %#v", removed.Controls)
	}
	removedStreams, err := SerializeForm(removed, 932)
	if err != nil {
		t.Fatal(err)
	}
	removedPrefix := base.Controls[0].Level.Path + "/"
	for path := range removedStreams.Streams {
		if len(path) >= len(removedPrefix) && path[:len(removedPrefix)] == removedPrefix {
			t.Fatalf("removed MultiPage left stream %q", path)
		}
	}
	unchanged, err := SerializeForm(base, 932)
	if err != nil || !reflect.DeepEqual(before, unchanged) {
		t.Fatal("successful removal changed input")
	}
}

func TestApplyTopologyMovesPageBetweenMultiPagesAtomically(t *testing.T) {
	base, err := NewForm(Definition{Name: "TopologyMove", Size: Size{8467, 6350}, Controls: []ControlDefinition{
		multiPageDefinition("FirstPages", "PageA", "PageB"),
		multiPageDefinition("SecondPages", "PageC"),
		frameDefinition("Container"),
	}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	before, err := SerializeForm(base, 932)
	if err != nil {
		t.Fatal(err)
	}
	pageB := base.Controls[0].MultiPage.Pages[1]
	wantTopology := []TopologyControl{
		{Name: "FirstPages"}, {Name: "PageA", Parent: "FirstPages"}, {Name: "PageB", Parent: "FirstPages"},
		{Name: "SecondPages"}, {Name: "PageC", Parent: "SecondPages"}, {Name: "Container"},
	}
	if got := topologyOf(base); !reflect.DeepEqual(got, wantTopology) {
		t.Fatalf("topologyOf included hidden strips or changed Page order: %#v", got)
	}
	desired := []TopologyControl{
		{Name: "FirstPages"}, {Name: "PageA", Parent: "FirstPages"},
		{Name: "SecondPages"}, {Name: "PageB", Parent: "SecondPages"}, {Name: "PageC", Parent: "SecondPages"},
		{Name: "Container"},
	}
	moved, err := ApplyTopology(base, desired, 932)
	if err != nil {
		t.Fatalf("Page move: %v", err)
	}
	first, second := moved.Controls[0].MultiPage, moved.Controls[1].MultiPage
	if len(first.Pages) != 1 || first.Pages[0].Name != "PageA" || len(second.Pages) != 2 || second.Pages[0].Name != "PageB" || second.Pages[1].Name != "PageC" {
		t.Fatalf("pages after move: first=%#v second=%#v", first.Pages, second.Pages)
	}
	if second.Pages[0].ID != pageB.ID || second.Hidden.TabStrip.Tabs[0].Caption != "PageB" || second.Hidden.TabStrip.Tabs[1].Caption != "PageC" {
		t.Fatal("moving a Page lost its identity or associated tab")
	}
	if second.Hidden.TabStrip.SelectedIndex != 1 {
		t.Fatalf("moved MultiPage selection=%d, want retained PageC at index 1", second.Hidden.TabStrip.SelectedIndex)
	}
	if _, err := SerializeForm(moved, 932); err != nil {
		t.Fatalf("moved topology did not serialize: %v", err)
	}

	invalidBase, err := NewForm(Definition{Name: "InvalidPageParent", Size: Size{8467, 6350}, Controls: []ControlDefinition{
		multiPageDefinition("Pages", "PageA", "PageB"),
		frameDefinition("Container"),
	}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	invalidBefore, err := SerializeForm(invalidBase, 932)
	if err != nil {
		t.Fatal(err)
	}
	invalid := topologyOf(invalidBase)
	for i := range invalid {
		if invalid[i].Name == "PageA" {
			invalid[i].Parent = "Container"
		}
	}
	if rejected, err := ApplyTopology(invalidBase, invalid, 932); rejected != nil || !errors.Is(err, ErrUnsupportedEdit) {
		t.Fatalf("Page under ordinary Frame: result=%v error=%v", rejected, err)
	}
	unchanged, err := SerializeForm(base, 932)
	if err != nil || !reflect.DeepEqual(before, unchanged) {
		t.Fatal("move changed input")
	}
	unchanged, err = SerializeForm(invalidBase, 932)
	if err != nil || !reflect.DeepEqual(invalidBefore, unchanged) {
		t.Fatal("invalid-parent attempt changed input")
	}
}

func TestGeneratedFrameIdentityMatchesExcelFixture(t *testing.T) {
	fixture, err := ReadForm(openFixture(t, "p6_nested_form.bin"), "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	generated, err := NewForm(Definition{Name: "Nested", Size: Size{1000, 1000}, Controls: []ControlDefinition{frameDefinition("Frame")}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	excel := fixture.Controls[1].Level
	got := generated.Controls[0].Level
	if got.StorageMeta.CLSID != excel.StorageMeta.CLSID || !bytes.Equal(got.CompObjRaw, excel.CompObjRaw) {
		t.Fatal("Frame CLSID/CompObj differs from real Excel")
	}
	if !bytes.Equal(got.ClassTableRaw, excel.ClassTableRaw) || len(got.ClassTableRaw) != 0 {
		t.Fatal("Frame DontSaveClassTable flag must omit the class-table count field")
	}
}

func TestFrameEditPreservesUnchangedMultiPageSubtree(t *testing.T) {
	base, err := ReadForm(openFixture(t, "p6_nested_form.bin"), "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	before, err := SerializeForm(base, 932)
	if err != nil {
		t.Fatal(err)
	}
	desired := topologyOf(base)
	frame := frameDefinition("AddedFrame")
	label := ControlDefinition{Name: "NewLabel", Class: 21, Size: Size{1000, 600}, Visible: true}
	desired = append(desired, TopologyControl{Name: "AddedFrame", Definition: &frame}, TopologyControl{Name: "NewLabel", Parent: "AddedFrame", Definition: &label})
	after, err := ApplyTopology(base, desired, 932)
	if err != nil {
		t.Fatal(err)
	}
	got, err := SerializeForm(after, 932)
	if err != nil {
		t.Fatal(err)
	}
	multiPagePrefix := base.Controls[1].Level.Path + "/"
	for path, raw := range before.Streams {
		if len(path) >= len(multiPagePrefix) && path[:len(multiPagePrefix)] == multiPagePrefix && !bytes.Equal(raw, got.Streams[path]) {
			t.Fatalf("special subtree changed: %s", path)
		}
	}
	addedFrame := after.Controls[len(after.Controls)-1]
	if addedFrame.Name != "AddedFrame" || len(addedFrame.Children) != 1 || after.Levels[0].Record.Values["NextAvailableID"] <= int64(addedFrame.Children[0].ID) {
		t.Fatal("root allocation counter did not include nested addition")
	}
}

func TestNewFormRejectsExcessiveDepth(t *testing.T) {
	control := frameDefinition("Deep")
	for i := range maxNestingDepth + 1 {
		control = frameDefinition(fmt.Sprintf("Frame%d", i), control)
	}
	if _, err := NewForm(Definition{Name: "Nested", Size: Size{1000, 1000}, Controls: []ControlDefinition{control}}, 1252); !errors.Is(err, ErrInvalidEdit) {
		t.Fatalf("depth limit: %v", err)
	}
}

func TestReadFormRejectsDuplicateSiteIDs(t *testing.T) {
	base, err := NewForm(Definition{Name: "Duplicate", Size: Size{1000, 1000}, Controls: []ControlDefinition{{Name: "A", Class: 21, Visible: true}, {Name: "B", Class: 21, Visible: true}}}, 1252)
	if err != nil {
		t.Fatal(err)
	}
	stored, err := SerializeForm(base, 1252)
	if err != nil {
		t.Fatal(err)
	}
	level := base.Levels[0]
	level.Sites[1].Values["ID"] = level.Sites[0].Values["ID"]
	level.Sites[1].Raw, err = encodeEditedSite(level.Sites[1])
	if err != nil {
		t.Fatal(err)
	}
	stored.Streams["Duplicate/f"], err = encodeEditedLevel(level)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := reparseEditedForm(stored, "Duplicate", 1252); !errors.Is(err, ErrMalformed) {
		t.Fatalf("duplicate site ID accepted: %v", err)
	}
}

func TestApplyTopologyMovesFrameSubtreeWithoutChangingMetadata(t *testing.T) {
	base, err := NewForm(Definition{Name: "Nested", Size: Size{1000, 1000}, Controls: []ControlDefinition{frameDefinition("A", frameDefinition("Child", frameDefinition("Grandchild"))), frameDefinition("B")}}, 1252)
	if err != nil {
		t.Fatal(err)
	}
	old := base.Controls[0].Children[0]
	before, err := SerializeForm(base, 1252)
	if err != nil {
		t.Fatal(err)
	}
	meta := old.Level.StorageMeta
	got, err := ApplyTopology(base, []TopologyControl{{Name: "A"}, {Name: "B"}, {Name: "Child", Parent: "B"}, {Name: "Grandchild", Parent: "Child"}}, 1252)
	if err != nil {
		t.Fatal(err)
	}
	child := got.Controls[1].Children[0]
	if child.ID != old.ID || child.Level.StorageMeta != meta || child.Level.Path != "Nested/i04/i02" || child.Children[0].Level.Path != "Nested/i04/i02/i03" {
		t.Fatal("subtree relocation lost identity")
	}
	stored, err := SerializeForm(got, 1252)
	if err != nil {
		t.Fatal(err)
	}
	if _, orphan := stored.Storages["Nested/i01/i02"]; orphan {
		t.Fatal("orphan moved subtree")
	}
	if !bytes.Equal(before.Streams["Nested/i01/i02/f"], stored.Streams["Nested/i04/i02/f"]) {
		t.Fatal("unchanged child f changed")
	}
}
