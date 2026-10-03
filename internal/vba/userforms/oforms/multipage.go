package oforms

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"math"
	"slices"
)

// New top-tab layouts use fixed 96 DPI and an explicit Tahoma 8.25pt font.
// The 2px border and 19px band do not query the build host's display/font APIs.
func generatedPageBox(size Size) (Position, Size) {
	wide := max(1, int64(math.Round(float64(size.Width)*96/2540)))
	tall := max(1, int64(math.Round(float64(size.Height)*96/2540)))
	scale := func(pixels, value, extent int64) int32 { return int32((2*pixels*value + extent) / (2 * extent)) }
	left, top := scale(2, int64(size.Width), wide), scale(21, int64(size.Height), tall)
	right, bottom := scale(max(2, wide-2), int64(size.Width), wide), scale(max(21, tall-2), int64(size.Height), tall)
	return Position{left, top}, Size{max(0, right-left), max(0, bottom-top)}
}

// MultiPage binds the hidden tab strip, page storage ownership and x records.
// Raw record bytes remain authoritative for lossless replay.
type MultiPage struct {
	Hidden         *Control
	Pages          []*Control
	Reserved       []byte
	PageProperties map[int32][]byte
	Properties     *Record
}

var pagePropertiesSpec = recordSpec{typeName: "PageProperties", major: 2, stopAfterExtra: true,
	data: []dataField{{1, "TransitionEffect", 4, fieldUnsigned}, {2, "TransitionPeriod", 4, fieldUnsigned}}}
var multiPagePropertiesSpec = recordSpec{typeName: "MultiPageProperties", major: 2, stopAfterExtra: true,
	data: []dataField{{1, "PageCount", 4, fieldSigned}, {2, "ID", 4, fieldSigned}}, flags: map[uint8]flagField{3: {name: "Flags", value: 0}}}

func bindMultiPage(control *Control, codePage uint16) error {
	level := control.Level
	if level == nil || !level.HasXStream {
		return fmt.Errorf("%w: MultiPage requires x stream", ErrMalformed)
	}
	state := &MultiPage{PageProperties: map[int32][]byte{}}
	pages := map[int32]*Control{}
	for _, child := range level.Controls {
		switch child.CLSIDCacheIndex {
		case 18:
			if state.Hidden != nil || child.Name != "" {
				return fmt.Errorf("%w: MultiPage hidden TabStrip ownership", ErrMalformed)
			}
			state.Hidden = child
		case 7:
			pages[child.ID] = child
		default:
			return fmt.Errorf("%w: MultiPage contains a non-Page control", ErrMalformed)
		}
	}
	if state.Hidden == nil {
		return fmt.Errorf("%w: MultiPage hidden TabStrip missing", ErrMalformed)
	}
	raw := level.XRaw
	offset := 0
	records := make([][]byte, 0, len(pages)+1)
	for range len(pages) + 1 {
		if len(raw)-offset < 8 {
			return fmt.Errorf("%w: truncated PageProperties", ErrMalformed)
		}
		end := offset + 4 + int(binary.LittleEndian.Uint16(raw[offset+2:]))
		if end > len(raw) {
			return fmt.Errorf("%w: PageProperties boundary", ErrMalformed)
		}
		if _, _, err := parseRecord(raw[offset:end], &pagePropertiesSpec, codePage, end-offset); err != nil {
			return err
		}
		records = append(records, bytes.Clone(raw[offset:end]))
		offset = end
	}
	if len(raw)-offset < 8 {
		return fmt.Errorf("%w: missing MultiPageProperties", ErrMalformed)
	}
	end := offset + 4 + int(binary.LittleEndian.Uint16(raw[offset+2:]))
	if end > len(raw) {
		return fmt.Errorf("%w: MultiPageProperties boundary", ErrMalformed)
	}
	props, _, err := parseRecord(raw[offset:end], &multiPagePropertiesSpec, codePage, end-offset)
	if err != nil {
		return err
	}
	if props.Values["PageCount"] != int64(len(pages)) || props.Values["ID"] != int64(state.Hidden.ID) || len(raw)-end != 4*len(pages) {
		return fmt.Errorf("%w: MultiPage page count/hidden ID/extent mismatch", ErrMalformed)
	}
	state.Reserved, state.Properties = records[0], props
	for index := range len(pages) {
		id := int32(binary.LittleEndian.Uint32(raw[end+4*index:]))
		page := pages[id]
		if page == nil {
			return fmt.Errorf("%w: MultiPage duplicated or missing page ID %d", ErrMalformed, id)
		}
		delete(pages, id)
		state.Pages = append(state.Pages, page)
		state.PageProperties[id] = records[index+1]
	}
	control.MultiPage = state
	// Excel can retain cached tab arrays after removing the final Page. The
	// x stream owns the empty topology; keep the hidden bytes for no-op replay.
	if state.Hidden.TabStrip == nil || len(state.Pages) > 0 && len(state.Hidden.TabStrip.Tabs) != len(state.Pages) {
		return fmt.Errorf("%w: MultiPage TabStrip/page count", ErrMalformed)
	}
	return nil
}

func rebuildMultiPage(owner *Control, wanted []*Control, oldTabs map[*Control]Tab, originalPageProperties map[int32][]byte, payloads map[*Control][]byte) error {
	state := owner.MultiPage
	previous := state.Pages
	position, pageSize := generatedPageBox(owner.Record.Sizes["DisplayedSize"])
	if len(previous) > 0 && previous[0].Site.Position != nil {
		position = *previous[0].Site.Position
		pageSize = previous[0].Record.Sizes["DisplayedSize"]
	}
	strip := *state.Hidden.TabStrip
	selected := int32(0)
	if strip.SelectedIndex >= 0 && int(strip.SelectedIndex) < len(previous) {
		selected = previous[strip.SelectedIndex].ID
	}
	state.Pages = slices.Clone(wanted[1:])
	strip.Tabs = nil
	strip.SelectedIndex = 0
	if len(state.Pages) == 0 {
		strip.SelectedIndex = -1
	}
	for i, page := range state.Pages {
		if page.CLSIDCacheIndex != 7 {
			return fmt.Errorf("%w: non-Page child in MultiPage", ErrInvalidEdit)
		}
		tab, ok := oldTabs[page]
		if !ok {
			tab = Tab{Name: fmt.Sprintf("Tab%d", page.ID), Caption: page.Name, Enabled: true, Visible: true}
		}
		strip.Tabs = append(strip.Tabs, tab)
		if raw := originalPageProperties[page.ID]; raw != nil {
			state.PageProperties[page.ID] = raw
		}
		if page.ID == selected {
			strip.SelectedIndex = int32(i)
		}
		if !slices.Contains(previous, page) {
			if !standardTabLayout(state.Hidden.Record) {
				return fmt.Errorf("%w: adding or moving a Page requires standard tab layout", ErrUnsupportedEdit)
			}
			page.Site.Position = new(Position{})
			*page.Site.Position, page.Level.Record.Sizes["DisplayedSize"] = position, pageSize
			var err error
			page.Site.Raw, err = encodeEditedSite(page.Site)
			if err != nil {
				return err
			}
			page.Level.Record.Raw, err = encodeEditedRecord(page.Level.Record, &formSpec)
			if err != nil {
				return err
			}
			page.Level.FRaw, err = encodeEditedLevel(page.Level)
			if err != nil {
				return err
			}
		}
	}
	// Topology may remove the selected identity or add a Page whose generated
	// Site is visible. Reconcile every remaining Site after the final index is
	// known, even when the subsequent TabEdit keeps that numeric index.
	for i, page := range state.Pages {
		flags := page.Site.Values["BitFlags"]
		next := editedBit(flags, int32(i) == strip.SelectedIndex)
		if flags != next {
			page.Site.Mask |= 1 << 4
			page.Site.Values["BitFlags"] = next
			var err error
			page.Site.Raw, err = encodeEditedSite(page.Site)
			if err != nil {
				return err
			}
		}
	}
	if err := SetTabStrip(state.Hidden.Record, &strip); err != nil {
		return err
	}
	raw, err := encodeEditedRecord(state.Hidden.Record, &tabStripSpec)
	if err != nil {
		return err
	}
	state.Hidden.Record.Raw = raw
	state.Hidden.Site.Values["ObjectStreamSize"] = int64(len(raw))
	state.Hidden.ObjectStreamSize = uint32(len(raw))
	state.Hidden.Site.Mask |= 1 << 5
	state.Hidden.Site.Raw, err = encodeEditedSite(state.Hidden.Site)
	if err != nil {
		return err
	}
	payloads[state.Hidden] = raw
	owner.Level.XRaw, err = encodeMultiPageX(state)
	return err
}

func encodeMultiPageX(state *MultiPage) ([]byte, error) {
	b := bytes.Clone(state.Reserved)
	if len(b) == 0 {
		b = []byte{0, 2, 4, 0, 0, 0, 0, 0}
	}
	for _, page := range state.Pages {
		raw := state.PageProperties[page.ID]
		if len(raw) == 0 {
			raw = []byte{0, 2, 4, 0, 0, 0, 0, 0}
		}
		b = append(b, raw...)
	}
	props := state.Properties
	if props == nil {
		props = newGenerationRecord(&multiPagePropertiesSpec)
	}
	if err := setGenerationProperty(props, &multiPagePropertiesSpec, "PageCount", int64(len(state.Pages))); err != nil {
		return nil, err
	}
	if err := setGenerationProperty(props, &multiPagePropertiesSpec, "ID", int64(state.Hidden.ID)); err != nil {
		return nil, err
	}
	raw, err := encodeEditedRecord(props, &multiPagePropertiesSpec)
	if err != nil {
		return nil, err
	}
	b = append(b, raw...)
	for _, page := range state.Pages {
		b = binary.LittleEndian.AppendUint32(b, uint32(page.ID))
	}
	if len(b) > maxDesignerStreamSize {
		return nil, ErrInvalidEdit
	}
	return b, nil
}

// Existing standard top-tab layouts keep their measured border and tab band.
// Resize applies the same extent delta to the hidden strip and each Page;
// child controls retain their Page-relative coordinates.
func resizeMultiPage(control *Control, property string, value int32, records map[*Record]Edit, edit Edit) error {
	hidden := control.MultiPage.Hidden.Record
	if !standardTabLayout(hidden) {
		return fmt.Errorf("%w: resizing this tab layout requires unmodeled layout bookkeeping", ErrUnsupportedEdit)
	}
	width := property == "Width"
	extent := func(size Size) int32 {
		if width {
			return size.Width
		}
		return size.Height
	}
	old := control.Record.Sizes["DisplayedSize"]
	delta := int64(value) - int64(extent(old))
	for _, entry := range []struct {
		record *Record
		name   string
	}{{hidden, "Size"}} {
		size, ok := entry.record.Sizes[entry.name]
		if !ok {
			return ErrUnsupportedEdit
		}
		next := int64(extent(size)) + delta
		if next < 0 || next > math.MaxInt32 {
			return ErrInvalidEdit
		}
		if width {
			size.Width = int32(next)
		} else {
			size.Height = int32(next)
		}
		entry.record.Sizes[entry.name] = size
		records[entry.record] = edit
	}
	for _, page := range control.MultiPage.Pages {
		size, ok := page.Record.Sizes["DisplayedSize"]
		if !ok {
			return ErrUnsupportedEdit
		}
		next := int64(extent(size)) + delta
		if next < 0 || next > math.MaxInt32 {
			return ErrInvalidEdit
		}
		if width {
			size.Width = int32(next)
		} else {
			size.Height = int32(next)
		}
		page.Record.Sizes["DisplayedSize"] = size
		records[page.Record] = edit
	}
	return nil
}

func standardTabLayout(record *Record) bool {
	return record.Values["TabOrientation"] == 0 && record.Values["TabStyle"] == 0 && record.Mask&(1<<10) == 0 && record.Values["TabFixedHeight"] == 0 && record.Values["TabFixedWidth"] == 0
}
