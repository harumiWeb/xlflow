package oforms

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"maps"
	"math"
	"unicode/utf16"
)

const (
	tabStripItemsBit       = 5
	tabStripTipStringsBit  = 15
	tabStripNamesBit       = 17
	tabStripTabsAllocBit   = 20
	tabStripTagsBit        = 21
	tabStripTabDataBit     = 22
	tabStripAcceleratorBit = 23
	tabStripSelectedBit    = 0
	tabStripVisibleBit     = 1 << 0
	tabStripEnabledBit     = 1 << 1
	tabStripFlagBits       = tabStripVisibleBit | tabStripEnabledBit
	maxArrayStringCount    = maxDesignerStreamSize / 4
)

type tabStringKind uint8

const (
	tabCaption tabStringKind = iota
	tabTip
	tabName
	tabTag
	tabAccelerator
	tabStringKindCount
)

type tabArraySpec struct {
	kind      tabStringKind
	name      string
	sizeField string
	bit       uint
}

var tabArraySpecs = [...]tabArraySpec{
	{kind: tabCaption, name: "Items", sizeField: "ItemsSize", bit: tabStripItemsBit},
	{kind: tabTip, name: "TipStrings", sizeField: "TipStringsSize", bit: tabStripTipStringsBit},
	{kind: tabName, name: "TabNames", sizeField: "NamesSize", bit: tabStripNamesBit},
	{kind: tabTag, name: "Tags", sizeField: "TagsSize", bit: tabStripTagsBit},
	{kind: tabAccelerator, name: "Accelerators", sizeField: "AcceleratorsSize", bit: tabStripAcceleratorBit},
}

// Tab is one persisted tab in a TabStrip control.
type Tab struct {
	Name           string
	Caption        string
	ControlTipText string
	Tag            string
	Accelerator    string
	Enabled        bool
	Visible        bool

	rawStrings [tabStringKindCount][]byte
	rawValues  [tabStringKindCount]string
	rawFlag    uint32
	hasRawFlag bool
}

// TabStrip is the tab-specific portion of a TabStrip property record.
// Raw flag words and trailing bytes stay private and are retained by ParseTabStrip.
type TabStrip struct {
	Tabs          []Tab
	SelectedIndex int32

	tabDataCount uint32
	flagsTail    []byte
	listIndexRaw bool
}

// ParseTabStrip decodes a TabStrip property record. ArrayString compression
// stores the low byte of UTF-16 code units and is independent of the VBA code page.
func ParseTabStrip(record *Record) (*TabStrip, error) {
	if record == nil {
		return nil, malformedTabStrip("nil record")
	}
	if record.Type != "TabStrip" || record.Major != 2 || record.Minor != 0 {
		return nil, malformedTabStrip("unexpected record identity %q version %d.%d", record.Type, record.Major, record.Minor)
	}

	arrays := make(map[tabStringKind]decodedTabArray, len(tabArraySpecs))
	arrayPresent := make(map[tabStringKind]bool, len(tabArraySpecs))
	totalArrayBytes := 0
	for _, spec := range tabArraySpecs {
		decoded, present, err := readTabArray(record, spec)
		if err != nil {
			return nil, err
		}
		if present {
			if len(decoded.raw) > maxDesignerStreamSize-totalArrayBytes {
				return nil, malformedTabStrip("string arrays exceed %d-byte limit", maxDesignerStreamSize)
			}
			totalArrayBytes += len(decoded.raw)
		}
		arrays[spec.kind] = decoded
		arrayPresent[spec.kind] = present
	}

	items := arrays[tabCaption].values
	if len(items) > maxArrayStringCount {
		return nil, malformedTabStrip("tab count %d exceeds limit", len(items))
	}
	for _, spec := range tabArraySpecs[1:] {
		if arrayPresent[spec.kind] && len(arrays[spec.kind].values) != len(items) {
			count := len(arrays[spec.kind].values)
			return nil, malformedTabStrip("%s has %d entries for %d Items", spec.name, count, len(items))
		}
	}
	strip := &TabStrip{
		Tabs:          make([]Tab, len(items)),
		SelectedIndex: -1,
		listIndexRaw:  record.Mask&(uint64(1)<<tabStripSelectedBit) != 0,
	}
	for index := range strip.Tabs {
		tab := &strip.Tabs[index]
		tab.Caption = items[index]
		tab.Visible = true
		tab.Enabled = true
		for _, spec := range tabArraySpecs {
			decoded := arrays[spec.kind]
			if len(decoded.values) != 0 {
				setTabString(tab, spec.kind, decoded.values[index])
				tab.rawValues[spec.kind] = decoded.values[index]
				tab.rawStrings[spec.kind] = bytes.Clone(decoded.elements[index])
			}
		}
	}

	if value, present := record.Values["ListIndex"]; present {
		if value < math.MinInt32 || value > math.MaxInt32 {
			return nil, malformedTabStrip("ListIndex %d is outside int32", value)
		}
		strip.SelectedIndex = int32(value)
	}
	if err := validateParsedSelectedIndex(strip); err != nil {
		return nil, err
	}

	dataCount, err := tabStripCount(record, "TabData", tabStripTabDataBit)
	if err != nil {
		return nil, err
	}
	if dataCount > uint32(len(strip.Tabs)) {
		return nil, malformedTabStrip("TabData %d exceeds Items tab count %d", dataCount, len(strip.Tabs))
	}
	strip.tabDataCount = dataCount
	flagBytes := uint64(dataCount) * 4
	if flagBytes > uint64(len(record.TailRaw)) {
		return nil, malformedTabStrip("TabData %d requires %d flag bytes, TailRaw has %d", dataCount, flagBytes, len(record.TailRaw))
	}
	for index := range dataCount {
		word := binary.LittleEndian.Uint32(record.TailRaw[index*4:])
		tab := &strip.Tabs[index]
		tab.rawFlag = word
		tab.hasRawFlag = true
		tab.Visible = word&tabStripVisibleBit != 0
		tab.Enabled = word&tabStripEnabledBit != 0
	}
	strip.flagsTail = bytes.Clone(record.TailRaw[int(flagBytes):])
	if totalArrayBytes > maxDesignerStreamSize-len(record.TailRaw) {
		return nil, malformedTabStrip("arrays and TabData exceed %d-byte limit", maxDesignerStreamSize)
	}
	return strip, nil
}

// SetTabStrip updates only the TabStrip arrays, ListIndex, allocation count,
// per-tab flags, and their opaque tail. All validation and encoding completes
// before record is changed.
func SetTabStrip(record *Record, state *TabStrip) error {
	if record == nil || state == nil {
		return fmt.Errorf("%w: TabStrip record and state are required", ErrInvalidEdit)
	}
	current, err := ParseTabStrip(record)
	if err != nil {
		return err
	}
	if len(state.Tabs) > maxArrayStringCount {
		return fmt.Errorf("%w: TabStrip tab count %d exceeds limit", ErrInvalidEdit, len(state.Tabs))
	}
	if equalTabStripState(current, state) {
		return nil
	}
	if err := validateSetSelectedIndex(state, current); err != nil {
		return err
	}
	collectionChanged := !equalTabCollection(current.Tabs, state.Tabs)

	arrays := make(map[string][]byte, len(tabArraySpecs))
	values := maps.Clone(record.Values)
	if values == nil {
		values = make(map[string]int64)
	}
	mask := record.Mask
	arrayBytes := 0
	for _, spec := range tabArraySpecs {
		present := record.Mask&(uint64(1)<<spec.bit) != 0 || spec.kind == tabCaption && len(state.Tabs) > 0 || tabArrayHasValue(state.Tabs, spec.kind)
		if !present {
			mask &^= uint64(1) << spec.bit
			delete(values, spec.sizeField)
			continue
		}
		encoded, err := encodeTabArray(state.Tabs, spec.kind, current.Tabs)
		if err != nil {
			return err
		}
		if len(encoded) > maxDesignerStreamSize-arrayBytes {
			return fmt.Errorf("%w: TabStrip arrays exceed %d-byte limit", ErrInvalidEdit, maxDesignerStreamSize)
		}
		arrayBytes += len(encoded)
		arrays[spec.name] = encoded
		values[spec.sizeField] = int64(len(encoded))
		mask |= uint64(1) << spec.bit
	}

	oldAllocated, err := tabStripCount(record, "TabsAllocated", tabStripTabsAllocBit)
	if err != nil {
		return fmt.Errorf("%w: invalid retained TabsAllocated: %v", ErrInvalidEdit, err)
	}
	allocated := uint64(oldAllocated)
	if len(state.Tabs) > len(current.Tabs) {
		allocated += uint64(len(state.Tabs) - len(current.Tabs))
	}
	if allocated > math.MaxUint32 {
		return fmt.Errorf("%w: TabsAllocated exceeds uint32", ErrInvalidEdit)
	}
	if allocated != 0 || record.Mask&(uint64(1)<<tabStripTabsAllocBit) != 0 {
		mask |= uint64(1) << tabStripTabsAllocBit
		values["TabsAllocated"] = int64(allocated)
	} else {
		mask &^= uint64(1) << tabStripTabsAllocBit
		delete(values, "TabsAllocated")
	}

	flagCount := min(current.tabDataCount, uint32(len(state.Tabs)))
	if collectionChanged || record.Mask&(uint64(1)<<tabStripTabDataBit) != 0 {
		flagCount = uint32(len(state.Tabs))
	}
	flagWords := make([]uint32, len(state.Tabs))
	retainedFlagsByName := make(map[string][]uint32, len(current.Tabs))
	for _, tab := range current.Tabs {
		if tab.Name != "" && tab.hasRawFlag {
			retainedFlagsByName[tab.Name] = append(retainedFlagsByName[tab.Name], tab.rawFlag)
		}
	}
	for index, tab := range state.Tabs {
		word := uint32(0)
		if tab.hasRawFlag {
			word = tab.rawFlag
		} else if tab.Name != "" {
			if index < len(current.Tabs) && current.Tabs[index].Name == tab.Name && current.Tabs[index].hasRawFlag {
				word = current.Tabs[index].rawFlag
			} else if retained := retainedFlagsByName[tab.Name]; len(retained) != 0 {
				word = retained[0]
			}
		}
		word = word &^ tabStripFlagBits
		if tab.Visible {
			word |= tabStripVisibleBit
		}
		if tab.Enabled {
			word |= tabStripEnabledBit
		}
		flagWords[index] = word
		if uint32(index) >= flagCount && (word&tabStripFlagBits != tabStripFlagBits || word&^tabStripFlagBits != 0) {
			flagCount = uint32(index + 1)
		}
	}
	if flagCount > uint32(len(state.Tabs)) {
		return fmt.Errorf("%w: TabData exceeds tab count", ErrInvalidEdit)
	}
	flagTail := make([]byte, 0, int(flagCount)*4+len(current.flagsTail))
	for _, word := range flagWords[:flagCount] {
		flagTail = binary.LittleEndian.AppendUint32(flagTail, word)
	}
	flagTail = append(flagTail, current.flagsTail...)
	if arrayBytes > maxDesignerStreamSize-len(flagTail) {
		return fmt.Errorf("%w: TabStrip record data exceeds %d-byte limit", ErrInvalidEdit, maxDesignerStreamSize)
	}
	if record.Mask&(uint64(1)<<tabStripTabDataBit) != 0 || flagCount > 0 {
		mask |= uint64(1) << tabStripTabDataBit
		values["TabData"] = int64(flagCount)
	} else {
		mask &^= uint64(1) << tabStripTabDataBit
		delete(values, "TabData")
	}

	if state.SelectedIndex != current.SelectedIndex || record.Mask&(uint64(1)<<tabStripSelectedBit) != 0 {
		mask |= uint64(1) << tabStripSelectedBit
		values["ListIndex"] = int64(state.SelectedIndex)
	} else {
		mask &^= uint64(1) << tabStripSelectedBit
		delete(values, "ListIndex")
	}

	arraysMap := maps.Clone(record.Arrays)
	if arraysMap == nil {
		arraysMap = make(map[string][]byte)
	}
	for _, spec := range tabArraySpecs {
		if encoded, present := arrays[spec.name]; present {
			arraysMap[spec.name] = encoded
		} else {
			delete(arraysMap, spec.name)
		}
	}
	if mask == record.Mask && equalInt64Maps(values, record.Values) && equalByteSliceMaps(arraysMap, record.Arrays) && bytes.Equal(flagTail, record.TailRaw) {
		return nil
	}

	record.Mask = mask
	record.Values = values
	record.Arrays = arraysMap
	record.TailRaw = flagTail
	record.Raw = nil
	return nil
}

type decodedTabArray struct {
	values   []string
	elements [][]byte
	raw      []byte
}

func readTabArray(record *Record, spec tabArraySpec) (decodedTabArray, bool, error) {
	present := record.Mask&(uint64(1)<<spec.bit) != 0
	raw, hasRaw := record.Arrays[spec.name]
	if present != hasRaw {
		return decodedTabArray{}, false, malformedTabStrip("%s presence disagrees with its property mask", spec.name)
	}
	if !present {
		if value, exists := record.Values[spec.sizeField]; exists && value != 0 {
			return decodedTabArray{}, false, malformedTabStrip("%s is omitted but %s is %d", spec.name, spec.sizeField, value)
		}
		return decodedTabArray{}, false, nil
	}
	size, exists := record.Values[spec.sizeField]
	if !exists || size < 0 || size > maxDesignerStreamSize || size != int64(len(raw)) {
		return decodedTabArray{}, false, malformedTabStrip("%s size %d does not match %d bytes", spec.sizeField, size, len(raw))
	}
	decoded, err := decodeTabArray(raw)
	if err != nil {
		return decodedTabArray{}, false, malformedTabStrip("%s: %v", spec.name, err)
	}
	return decoded, true, nil
}

func decodeTabArray(raw []byte) (decodedTabArray, error) {
	if len(raw) > maxDesignerStreamSize {
		return decodedTabArray{}, fmt.Errorf("array exceeds %d-byte limit", maxDesignerStreamSize)
	}
	decoded := decodedTabArray{raw: bytes.Clone(raw)}
	for offset := 0; offset < len(raw); {
		start := offset
		if len(raw)-offset < 4 {
			return decodedTabArray{}, fmt.Errorf("truncated CountAndCompression at byte %d", offset)
		}
		countAndCompression := binary.LittleEndian.Uint32(raw[offset:])
		offset += 4
		count := uint64(countAndCompression & 0x7fffffff)
		compressed := countAndCompression&0x80000000 != 0
		payloadLength := count
		if !compressed {
			payloadLength *= 2
		}
		if payloadLength > uint64(len(raw)-offset) {
			return decodedTabArray{}, fmt.Errorf("string payload of %d bytes exceeds remaining %d", payloadLength, len(raw)-offset)
		}
		payload := raw[offset : offset+int(payloadLength)]
		offset += int(payloadLength)
		value, err := decodeArrayString(payload, int(count), compressed)
		if err != nil {
			return decodedTabArray{}, err
		}
		padding := (4 - (offset-start)%4) % 4
		if padding > len(raw)-offset {
			return decodedTabArray{}, fmt.Errorf("truncated ArrayString padding at byte %d", offset)
		}
		offset += padding
		decoded.values = append(decoded.values, value)
		decoded.elements = append(decoded.elements, bytes.Clone(raw[start:offset]))
		if len(decoded.values) > maxArrayStringCount {
			return decodedTabArray{}, fmt.Errorf("array element count exceeds limit")
		}
	}
	return decoded, nil
}

func decodeArrayString(payload []byte, codeUnits int, compressed bool) (string, error) {
	if compressed {
		if len(payload) != codeUnits {
			return "", fmt.Errorf("compressed string has %d bytes for %d UTF-16 units", len(payload), codeUnits)
		}
		runes := make([]rune, len(payload))
		for index, value := range payload {
			runes[index] = rune(value)
		}
		return string(runes), nil
	}
	if len(payload) != codeUnits*2 {
		return "", fmt.Errorf("uncompressed string has %d bytes for %d UTF-16 units", len(payload), codeUnits)
	}
	units := make([]uint16, codeUnits)
	for index := range units {
		units[index] = binary.LittleEndian.Uint16(payload[index*2:])
	}
	return string(utf16.Decode(units)), nil
}

func encodeTabArray(tabs []Tab, kind tabStringKind, fallback []Tab) ([]byte, error) {
	encoded := make([]byte, 0)
	for index := range tabs {
		value := tabString(tabs[index], kind)
		var raw []byte
		if tabs[index].rawValues[kind] == value {
			raw = tabs[index].rawStrings[kind]
		}
		if len(raw) == 0 && index < len(fallback) && tabString(fallback[index], kind) == value && fallback[index].rawValues[kind] == value {
			raw = fallback[index].rawStrings[kind]
		}
		if len(raw) == 0 {
			var err error
			raw, err = encodeArrayString(value)
			if err != nil {
				return nil, fmt.Errorf("%w: encode TabStrip string: %v", ErrInvalidEdit, err)
			}
		}
		if len(raw) > maxDesignerStreamSize-len(encoded) {
			return nil, fmt.Errorf("%w: TabStrip string array exceeds %d-byte limit", ErrInvalidEdit, maxDesignerStreamSize)
		}
		encoded = append(encoded, raw...)
	}
	return encoded, nil
}

func encodeArrayString(value string) ([]byte, error) {
	units := utf16.Encode([]rune(value))
	if uint64(len(units)) > 0x7fffffff {
		return nil, fmt.Errorf("UTF-16 string is too long")
	}
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
	if len(encoded) > maxDesignerStreamSize {
		return nil, fmt.Errorf("ArrayString exceeds %d-byte limit", maxDesignerStreamSize)
	}
	for len(encoded)%4 != 0 {
		encoded = append(encoded, 0)
	}
	return encoded, nil
}

func tabStripCount(record *Record, field string, bit uint) (uint32, error) {
	if record.Mask&(uint64(1)<<bit) == 0 {
		if value, exists := record.Values[field]; exists && value != 0 {
			return 0, malformedTabStrip("%s is omitted but has value %d", field, value)
		}
		return 0, nil
	}
	value, exists := record.Values[field]
	if !exists || value < 0 || uint64(value) > math.MaxUint32 {
		return 0, malformedTabStrip("%s has invalid unsigned value %d", field, value)
	}
	return uint32(value), nil
}

func validateParsedSelectedIndex(strip *TabStrip) error {
	if strip.SelectedIndex < -1 {
		return malformedTabStrip("ListIndex %d is less than -1", strip.SelectedIndex)
	}
	if len(strip.Tabs) > 0 && strip.SelectedIndex >= int32(len(strip.Tabs)) {
		return malformedTabStrip("ListIndex %d exceeds Items tab count %d", strip.SelectedIndex, len(strip.Tabs))
	}
	// Excel can retain a stale non-negative ListIndex when an empty TabStrip is
	// nested in a MultiPage. Preserve that persisted value for a no-op edit.
	return nil
}

func validateSetSelectedIndex(state, current *TabStrip) error {
	index := state.SelectedIndex
	if index < -1 {
		return fmt.Errorf("%w: ListIndex %d is less than -1", ErrInvalidEdit, index)
	}
	if len(state.Tabs) > 0 && index >= int32(len(state.Tabs)) {
		return fmt.Errorf("%w: ListIndex %d exceeds tab count %d", ErrInvalidEdit, index, len(state.Tabs))
	}
	if len(state.Tabs) == 0 && index >= 0 {
		if len(current.Tabs) == 0 && index == current.SelectedIndex && current.listIndexRaw {
			return nil
		}
		return fmt.Errorf("%w: non-negative ListIndex %d is only retained for an unchanged empty TabStrip", ErrInvalidEdit, index)
	}
	return nil
}

func equalTabStripState(left, right *TabStrip) bool {
	return left.SelectedIndex == right.SelectedIndex && equalTabCollection(left.Tabs, right.Tabs)
}

func equalTabCollection(left, right []Tab) bool {
	if len(left) != len(right) {
		return false
	}
	for index := range left {
		if !equalTabValues(left[index], right[index]) {
			return false
		}
	}
	return true
}

func equalTabValues(left, right Tab) bool {
	return left.Name == right.Name && left.Caption == right.Caption && left.ControlTipText == right.ControlTipText && left.Tag == right.Tag && left.Accelerator == right.Accelerator && left.Enabled == right.Enabled && left.Visible == right.Visible
}

func tabString(tab Tab, kind tabStringKind) string {
	switch kind {
	case tabCaption:
		return tab.Caption
	case tabTip:
		return tab.ControlTipText
	case tabName:
		return tab.Name
	case tabTag:
		return tab.Tag
	case tabAccelerator:
		return tab.Accelerator
	default:
		return ""
	}
}

func setTabString(tab *Tab, kind tabStringKind, value string) {
	switch kind {
	case tabCaption:
		tab.Caption = value
	case tabTip:
		tab.ControlTipText = value
	case tabName:
		tab.Name = value
	case tabTag:
		tab.Tag = value
	case tabAccelerator:
		tab.Accelerator = value
	}
}

func tabArrayHasValue(tabs []Tab, kind tabStringKind) bool {
	for _, tab := range tabs {
		if tabString(tab, kind) != "" {
			return true
		}
	}
	return false
}

func malformedTabStrip(format string, args ...any) error {
	return fmt.Errorf("%w: TabStrip: %s", ErrMalformed, fmt.Sprintf(format, args...))
}

func equalInt64Maps(left, right map[string]int64) bool {
	if len(left) != len(right) {
		return false
	}
	for key, value := range left {
		if other, exists := right[key]; !exists || other != value {
			return false
		}
	}
	return true
}

func equalByteSliceMaps(left, right map[string][]byte) bool {
	if len(left) != len(right) {
		return false
	}
	for key, value := range left {
		if !bytes.Equal(value, right[key]) {
			return false
		}
	}
	return true
}
