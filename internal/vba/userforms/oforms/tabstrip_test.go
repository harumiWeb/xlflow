package oforms

import (
	"bytes"
	"encoding/binary"
	"os"
	"reflect"
	"testing"
	"unicode/utf16"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
)

func TestParseTabStripArrayStringsAndOmittedDefaults(t *testing.T) {
	record := newTabStripTestRecord([]Tab{
		{Caption: "A", Enabled: true, Visible: true},
		{Caption: "日", Enabled: true, Visible: true},
	}, -1, false, 4, true, nil, true, []byte{0xde, 0xad}, nil)
	strip, err := ParseTabStrip(record)
	if err != nil {
		t.Fatal(err)
	}
	if len(strip.Tabs) != 2 || strip.SelectedIndex != -1 {
		t.Fatalf("strip has %d tabs and selected index %d", len(strip.Tabs), strip.SelectedIndex)
	}
	if strip.Tabs[0].Caption != "A" || strip.Tabs[1].Caption != "日" {
		t.Fatalf("captions = %q, %q", strip.Tabs[0].Caption, strip.Tabs[1].Caption)
	}
	for index, tab := range strip.Tabs {
		if tab.Name != "" || tab.ControlTipText != "" || tab.Tag != "" || tab.Accelerator != "" {
			t.Fatalf("tab %d omitted strings = %+v", index, tab)
		}
		if !tab.Visible || !tab.Enabled {
			t.Fatalf("tab %d omitted flags = visible %t enabled %t", index, tab.Visible, tab.Enabled)
		}
	}
	if got, want := record.Arrays["Items"], []byte{
		1, 0, 0, 0x80, 'A', 0, 0, 0,
		1, 0, 0, 0, 0xe5, 0x65, 0, 0,
	}; !bytes.Equal(got, want) {
		t.Fatalf("Items encoding = %x, want %x", got, want)
	}
	if record.Values["TabsAllocated"] != 4 || record.Values["TabData"] != 0 {
		t.Fatalf("allocation/data counts = %d/%d", record.Values["TabsAllocated"], record.Values["TabData"])
	}
	if !bytes.Equal(strip.flagsTail, []byte{0xde, 0xad}) {
		t.Fatalf("opaque tail = %x", strip.flagsTail)
	}
}

func TestParseTabStripUTF16SurrogatePair(t *testing.T) {
	record := newTabStripTestRecord([]Tab{{Caption: "😀", Enabled: true, Visible: true}}, -1, false, 0, false, nil, false, nil, nil)
	strip, err := ParseTabStrip(record)
	if err != nil {
		t.Fatal(err)
	}
	if got := strip.Tabs[0].Caption; got != "😀" {
		t.Fatalf("caption = %q", got)
	}
	if got, want := record.Arrays["Items"], []byte{2, 0, 0, 0, 0x3d, 0xd8, 0, 0xde}; !bytes.Equal(got, want) {
		t.Fatalf("surrogate-pair encoding = %x, want %x", got, want)
	}
}

func TestParseTabStripRetainsEmptyStaleListIndex(t *testing.T) {
	record := newTabStripTestRecord(nil, 1, true, 4, true, nil, true, []byte{0xca, 0xfe}, nil)
	strip, err := ParseTabStrip(record)
	if err != nil {
		t.Fatal(err)
	}
	if len(strip.Tabs) != 0 || strip.SelectedIndex != 1 {
		t.Fatalf("empty strip state = tabs %d selected %d", len(strip.Tabs), strip.SelectedIndex)
	}
	before := cloneTabStripTestRecord(record)
	if err := SetTabStrip(record, strip); err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(record, before) {
		t.Fatal("no-op SetTabStrip changed the record")
	}
}

func TestParseTabStripChecksTabDataAndSelectedIndex(t *testing.T) {
	emptyTips := newTabStripTestRecord([]Tab{{Caption: "one", Visible: true, Enabled: true}}, -1, true, 0, false, nil, true, nil, map[tabStringKind]bool{tabTip: true})
	emptyTips.Arrays["TipStrings"] = nil
	emptyTips.Values["TipStringsSize"] = 0
	tests := []struct {
		name   string
		record *Record
	}{
		{
			name:   "TabData cannot exceed item count",
			record: newTabStripTestRecord([]Tab{{Caption: "one", Visible: true, Enabled: true}}, -1, true, 0, false, []uint32{3, 3}, true, nil, nil),
		},
		{
			name:   "ListIndex cannot exceed non-empty item count",
			record: newTabStripTestRecord([]Tab{{Caption: "one", Visible: true, Enabled: true}}, 1, true, 1, true, nil, true, nil, nil),
		},
		{
			name:   "flag array must be bounded by TailRaw",
			record: tabStripRecordWithoutFlags(),
		},
		{
			name:   "present empty array disagrees with Items count",
			record: emptyTips,
		},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			if _, err := ParseTabStrip(test.record); err == nil {
				t.Fatal("ParseTabStrip accepted inconsistent data")
			}
		})
	}
}

func TestParseTabStripRejectsOversizedArray(t *testing.T) {
	record := newTabStripTestRecord([]Tab{{Caption: "one", Visible: true, Enabled: true}}, -1, false, 0, false, nil, false, nil, nil)
	record.Values["ItemsSize"] = maxDesignerStreamSize + 1
	if _, err := ParseTabStrip(record); err == nil {
		t.Fatal("ParseTabStrip accepted an oversized Items array")
	}
}

func tabStripRecordWithoutFlags() *Record {
	record := newTabStripTestRecord([]Tab{{Caption: "one", Visible: true, Enabled: true}}, -1, true, 0, false, []uint32{3}, true, nil, nil)
	record.TailRaw = nil
	return record
}

func TestSetTabStripPreservesOpaqueFlagsAndTail(t *testing.T) {
	record := newTabStripTestRecord([]Tab{
		{Caption: "Alpha", Enabled: true, Visible: true},
		{Caption: "Beta", Enabled: false, Visible: false},
	}, 0, true, 4, true, []uint32{0xaabbccf3, 0x12345670}, true, []byte{0xde, 0xad}, nil)
	record.Mask |= uint64(1) << 19
	record.Values["NewVersion"] = 1
	strip, err := ParseTabStrip(record)
	if err != nil {
		t.Fatal(err)
	}
	strip.Tabs[0].Caption = "Renamed"
	strip.Tabs[0].Enabled = false
	strip.Tabs[0].Name = "PageAlpha"
	strip.Tabs[0].Tag = "alpha-tag"
	strip.Tabs[1].ControlTipText = "tip"
	strip.Tabs[1].Accelerator = "B"
	strip.SelectedIndex = 1
	if err := SetTabStrip(record, strip); err != nil {
		t.Fatal(err)
	}
	if record.Values["TabsAllocated"] != 4 {
		t.Fatalf("TabsAllocated = %d, want retained allocation count 4", record.Values["TabsAllocated"])
	}
	if record.Mask&(uint64(1)<<19) == 0 || record.Values["NewVersion"] != 1 {
		t.Fatal("SetTabStrip did not retain the NewVersion property")
	}
	if record.Values["TabData"] != 2 {
		t.Fatalf("TabData = %d, want 2", record.Values["TabData"])
	}
	if record.Values["ListIndex"] != 1 {
		t.Fatalf("ListIndex = %d, want 1", record.Values["ListIndex"])
	}
	wantFlags := []byte{0xf1, 0xcc, 0xbb, 0xaa, 0x70, 0x56, 0x34, 0x12, 0xde, 0xad}
	if !bytes.Equal(record.TailRaw, wantFlags) {
		t.Fatalf("TailRaw = %x, want %x", record.TailRaw, wantFlags)
	}
	parsed, err := ParseTabStrip(record)
	if err != nil {
		t.Fatal(err)
	}
	if parsed.Tabs[0].Caption != "Renamed" || parsed.Tabs[0].Name != "PageAlpha" || parsed.Tabs[0].Tag != "alpha-tag" || parsed.Tabs[1].ControlTipText != "tip" || parsed.Tabs[1].Accelerator != "B" {
		t.Fatalf("round-trip tabs = %+v", parsed.Tabs)
	}
	if parsed.Tabs[0].Enabled || !parsed.Tabs[0].Visible || parsed.Tabs[1].Enabled || parsed.Tabs[1].Visible {
		t.Fatalf("round-trip flags = %+v", parsed.Tabs)
	}
}

func TestSetTabStripStructuralChangesPersistCompleteFlagArray(t *testing.T) {
	opaqueTail := []byte{0xde, 0xad}
	record := newTabStripTestRecord([]Tab{
		{Caption: "Alpha", Enabled: false, Visible: false},
		{Caption: "Beta", Enabled: true, Visible: true},
	}, 0, true, 2, true, []uint32{0xaabbccf0, 0x12345673}, true, opaqueTail, nil)
	strip, err := ParseTabStrip(record)
	if err != nil {
		t.Fatal(err)
	}
	strip.Tabs = append(strip.Tabs, Tab{Caption: "Gamma", Enabled: true, Visible: true})
	if err := SetTabStrip(record, strip); err != nil {
		t.Fatal(err)
	}
	if got := record.Values["TabData"]; got != 3 {
		t.Fatalf("TabData after growth = %d, want 3", got)
	}
	wantGrowthTail := binary.LittleEndian.AppendUint32(nil, 0xaabbccf0)
	wantGrowthTail = binary.LittleEndian.AppendUint32(wantGrowthTail, 0x12345673)
	wantGrowthTail = binary.LittleEndian.AppendUint32(wantGrowthTail, 0x00000003)
	wantGrowthTail = append(wantGrowthTail, opaqueTail...)
	if !bytes.Equal(record.TailRaw, wantGrowthTail) {
		t.Fatalf("TailRaw after growth = %x, want %x", record.TailRaw, wantGrowthTail)
	}
	grown, err := ParseTabStrip(record)
	if err != nil {
		t.Fatal(err)
	}
	if len(grown.Tabs) != 3 || grown.Tabs[0].Enabled || grown.Tabs[0].Visible || !grown.Tabs[1].Enabled || !grown.Tabs[1].Visible || !grown.Tabs[2].Enabled || !grown.Tabs[2].Visible {
		t.Fatalf("flags after growth = %+v", grown.Tabs)
	}
	if grown.Tabs[0].rawFlag != 0xaabbccf0 || grown.Tabs[1].rawFlag != 0x12345673 || !bytes.Equal(grown.flagsTail, opaqueTail) {
		t.Fatalf("opaque flag bits/tail after growth = %08x, %08x, %x", grown.Tabs[0].rawFlag, grown.Tabs[1].rawFlag, grown.flagsTail)
	}

	grown.Tabs = grown.Tabs[:2]
	if err := SetTabStrip(record, grown); err != nil {
		t.Fatal(err)
	}
	if got := record.Values["TabData"]; got != 2 {
		t.Fatalf("TabData after shrink = %d, want 2", got)
	}
	wantShrinkTail := binary.LittleEndian.AppendUint32(nil, 0xaabbccf0)
	wantShrinkTail = binary.LittleEndian.AppendUint32(wantShrinkTail, 0x12345673)
	wantShrinkTail = append(wantShrinkTail, opaqueTail...)
	if !bytes.Equal(record.TailRaw, wantShrinkTail) {
		t.Fatalf("TailRaw after shrink = %x, want %x", record.TailRaw, wantShrinkTail)
	}
	if got := record.Values["TabsAllocated"]; got != 3 {
		t.Fatalf("TabsAllocated after shrink = %d, want retained allocation count 3", got)
	}
	shrunk, err := ParseTabStrip(record)
	if err != nil {
		t.Fatal(err)
	}
	if len(shrunk.Tabs) != 2 || shrunk.Tabs[0].Enabled || !shrunk.Tabs[1].Enabled || !bytes.Equal(shrunk.flagsTail, opaqueTail) {
		t.Fatalf("flags/tail after shrink = %+v / %x", shrunk.Tabs, shrunk.flagsTail)
	}
}

func TestSetTabStripFlagFallbackMatchesRetainedNames(t *testing.T) {
	tabs := []Tab{
		{Name: "Alpha", Caption: "Alpha", Enabled: false, Visible: false},
		{Name: "Beta", Caption: "Beta", Enabled: true, Visible: true},
	}
	flags := []uint32{0xaabbccf0, 0x12345673}
	present := map[tabStringKind]bool{tabName: true}

	t.Run("replacement does not inherit by index", func(t *testing.T) {
		record := newTabStripTestRecord(tabs, 0, true, 2, true, flags, true, nil, present)
		strip, err := ParseTabStrip(record)
		if err != nil {
			t.Fatal(err)
		}
		strip.Tabs[0] = Tab{Name: "Gamma", Caption: "Gamma", Enabled: true, Visible: true}
		if err := SetTabStrip(record, strip); err != nil {
			t.Fatal(err)
		}
		got := []uint32{
			binary.LittleEndian.Uint32(record.TailRaw[0:4]),
			binary.LittleEndian.Uint32(record.TailRaw[4:8]),
		}
		if want := []uint32{0x00000003, 0x12345673}; !reflect.DeepEqual(got, want) {
			t.Fatalf("replacement flags = %08x, want %08x", got, want)
		}
	})

	t.Run("retained reorder follows names", func(t *testing.T) {
		record := newTabStripTestRecord(tabs, 0, true, 2, true, flags, true, nil, present)
		strip, err := ParseTabStrip(record)
		if err != nil {
			t.Fatal(err)
		}
		strip.Tabs[0], strip.Tabs[1] = strip.Tabs[1], strip.Tabs[0]
		for index := range strip.Tabs {
			strip.Tabs[index].hasRawFlag = false
			strip.Tabs[index].rawFlag = 0
		}
		if err := SetTabStrip(record, strip); err != nil {
			t.Fatal(err)
		}
		got := []uint32{
			binary.LittleEndian.Uint32(record.TailRaw[0:4]),
			binary.LittleEndian.Uint32(record.TailRaw[4:8]),
		}
		if want := []uint32{0x12345673, 0xaabbccf0}; !reflect.DeepEqual(got, want) {
			t.Fatalf("reordered flags = %08x, want %08x", got, want)
		}
	})
}

func TestSetTabStripAllocationCountCanExceedCurrentItems(t *testing.T) {
	record := newTabStripTestRecord(nil, -1, true, 4, true, nil, true, nil, nil)
	strip := &TabStrip{
		Tabs: []Tab{
			{Caption: "One", Enabled: true, Visible: true},
			{Caption: "Two", Enabled: true, Visible: true},
		},
		SelectedIndex: 0,
	}
	if err := SetTabStrip(record, strip); err != nil {
		t.Fatal(err)
	}
	if got := record.Values["TabsAllocated"]; got != 6 {
		t.Fatalf("TabsAllocated after adding two tabs = %d, want 6", got)
	}
	if got := record.Values["TabData"]; got != 2 {
		t.Fatalf("TabData after structural growth = %d, want complete count 2", got)
	}
	if got, want := record.TailRaw, []byte{3, 0, 0, 0, 3, 0, 0, 0}; !bytes.Equal(got, want) {
		t.Fatalf("default flags after structural growth = %x, want %x", got, want)
	}
	parsed, err := ParseTabStrip(record)
	if err != nil {
		t.Fatal(err)
	}
	if len(parsed.Tabs) != 2 || !parsed.Tabs[0].Visible || !parsed.Tabs[1].Enabled {
		t.Fatalf("new tabs = %+v", parsed.Tabs)
	}
}

func TestSetTabStripRejectsInvalidIndexWithoutMutation(t *testing.T) {
	record := newTabStripTestRecord([]Tab{{Caption: "One", Enabled: true, Visible: true}}, 0, true, 1, true, nil, true, nil, nil)
	before := cloneTabStripTestRecord(record)
	strip, err := ParseTabStrip(record)
	if err != nil {
		t.Fatal(err)
	}
	strip.SelectedIndex = 1
	if err := SetTabStrip(record, strip); err == nil {
		t.Fatal("SetTabStrip accepted an out-of-range ListIndex")
	}
	if !reflect.DeepEqual(record, before) {
		t.Fatal("rejected SetTabStrip changed the record")
	}
}

func TestParseTabStripNestedFixture(t *testing.T) {
	form, err := ReadForm(openFixture(t, "p6_nested_form.bin"), "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	seen := 0
	var walk func([]*Control)
	walk = func(controls []*Control) {
		for _, control := range controls {
			if control.Record != nil && control.Record.Type == "TabStrip" {
				seen++
				strip, err := ParseTabStrip(control.Record)
				if err != nil {
					t.Errorf("ParseTabStrip(%s): %v", control.Name, err)
				} else if len(strip.Tabs) == 0 {
					t.Errorf("fixture TabStrip %s unexpectedly has no tabs", control.Name)
				}
			}
			walk(control.Children)
		}
	}
	walk(form.Controls)
	if seen == 0 {
		t.Skip("p6_nested_form.bin contains no TabStrip record")
	}
}

func TestParseTabStripCapturedFixture(t *testing.T) {
	path := os.Getenv("XLFLOW_TABSTRIP_EVIDENCE")
	if path == "" {
		t.Skip("developer-only saved TabStrip evidence inspection")
	}
	body, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	container, err := cfb.Open(body)
	if err != nil {
		t.Fatal(err)
	}
	forms := DiscoverForms(container)
	if len(forms) == 0 {
		t.Fatal("captured binary has no forms")
	}
	seen := 0
	for _, name := range forms {
		form, err := ReadForm(container, name, 932)
		if err != nil {
			t.Fatal(err)
		}
		for _, level := range form.Levels {
			for _, control := range level.Controls {
				if control.Record == nil || control.Record.Type != "TabStrip" {
					continue
				}
				strip, err := ParseTabStrip(control.Record)
				if err != nil {
					t.Fatalf("TabStrip at %s: %v", level.Path, err)
				}
				seen++
				for index, tab := range strip.Tabs {
					t.Logf("TabStrip at %s tab %d: name=%q caption=%q tip=%q tag=%q accelerator=%q enabled=%t visible=%t", level.Path, index, tab.Name, tab.Caption, tab.ControlTipText, tab.Tag, tab.Accelerator, tab.Enabled, tab.Visible)
				}
				t.Logf("TabStrip at %s: selected=%d allocated=%d data=%d trailing=%x", level.Path, strip.SelectedIndex, control.Record.Values["TabsAllocated"], control.Record.Values["TabData"], strip.flagsTail)
			}
		}
	}
	if seen == 0 {
		t.Fatal("captured binary has no TabStrip records")
	}
}

func newTabStripTestRecord(tabs []Tab, selected int32, selectedPresent bool, allocated uint32, allocatedPresent bool, flags []uint32, tabDataPresent bool, tail []byte, present map[tabStringKind]bool) *Record {
	record := &Record{
		Type: "TabStrip", Major: 2, MaskWidth: 4,
		Values: make(map[string]int64), Arrays: make(map[string][]byte),
	}
	if selectedPresent {
		record.Mask |= uint64(1) << tabStripSelectedBit
		record.Values["ListIndex"] = int64(selected)
	}
	for _, spec := range tabArraySpecs {
		include := spec.kind == tabCaption && len(tabs) > 0 || present[spec.kind]
		if !include {
			continue
		}
		array := make([]byte, 0)
		for _, tab := range tabs {
			array = append(array, encodeTestArrayString(tabString(tab, spec.kind))...)
		}
		record.Mask |= uint64(1) << spec.bit
		record.Values[spec.sizeField] = int64(len(array))
		record.Arrays[spec.name] = array
	}
	if allocatedPresent || allocated != 0 {
		record.Mask |= uint64(1) << tabStripTabsAllocBit
		record.Values["TabsAllocated"] = int64(allocated)
	}
	if tabDataPresent || len(flags) > 0 {
		record.Mask |= uint64(1) << tabStripTabDataBit
		record.Values["TabData"] = int64(len(flags))
	}
	for _, word := range flags {
		record.TailRaw = binary.LittleEndian.AppendUint32(record.TailRaw, word)
	}
	record.TailRaw = append(record.TailRaw, tail...)
	return record
}

func encodeTestArrayString(value string) []byte {
	units := utf16.Encode([]rune(value))
	compressed := true
	for _, unit := range units {
		if unit > 0xff {
			compressed = false
			break
		}
	}
	count := uint32(len(units))
	if compressed {
		count |= 0x80000000
	}
	encoded := binary.LittleEndian.AppendUint32(nil, count)
	if compressed {
		for _, unit := range units {
			encoded = append(encoded, byte(unit))
		}
	} else {
		for _, unit := range units {
			encoded = binary.LittleEndian.AppendUint16(encoded, unit)
		}
	}
	for len(encoded)%4 != 0 {
		encoded = append(encoded, 0)
	}
	return encoded
}

func cloneTabStripTestRecord(record *Record) *Record {
	clone := *record
	clone.Values = make(map[string]int64, len(record.Values))
	for key, value := range record.Values {
		clone.Values[key] = value
	}
	clone.Arrays = make(map[string][]byte, len(record.Arrays))
	for key, value := range record.Arrays {
		clone.Arrays[key] = bytes.Clone(value)
	}
	clone.TailRaw = bytes.Clone(record.TailRaw)
	clone.Raw = bytes.Clone(record.Raw)
	return &clone
}
