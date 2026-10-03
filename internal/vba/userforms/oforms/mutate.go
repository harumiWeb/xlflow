package oforms

import (
	"bytes"
	"encoding/binary"
	"errors"
	"fmt"
	"maps"
	"math"
	"reflect"
	"slices"
	"strings"
	"unicode/utf16"
	"unicode/utf8"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
)

var (
	ErrUnsupportedEdit = errors.New("oforms: unsupported edit")
	ErrInvalidEdit     = errors.New("oforms: invalid edit")
)

// Edit addresses a persistence property. An empty Control selects the root.
// Position and size values are int32 HIMETRIC; TabIndex is int16.
type Edit struct {
	Control  string
	Property string
	Value    any
}

// EditError retains the edit address and supports errors.Is/errors.As.
type EditError struct {
	Control, Property string
	Err               error
}

func (e *EditError) Error() string {
	return fmt.Sprintf("oforms edit %q.%s: %v", e.Control, e.Property, e.Err)
}
func (e *EditError) Unwrap() error { return e.Err }

// ApplyEdits validates the signed input, edits an independent clone, and returns
// a newly parsed signed model. Arbitrary model changes remain unsupported.
func ApplyEdits(base *Form, edits []Edit, codePage uint16) (*Form, error) {
	stored, err := SerializeForm(base, codePage)
	if err != nil {
		return nil, err
	}
	clone, err := reparseEditedForm(stored, base.Name, codePage)
	if err != nil {
		return nil, err
	}
	if len(edits) == 0 {
		return clone, nil
	}
	records := map[*Record]Edit{}
	sites := map[*Site]Edit{}
	var rootCaption *string
	for _, edit := range edits {
		if edit.Control == "" && edit.Property == "Caption" {
			caption, ok := edit.Value.(string)
			if !ok {
				return nil, &EditError{Control: edit.Control, Property: edit.Property, Err: ErrInvalidEdit}
			}
			rootCaption = new(caption)
			if storedCaption, found := clone.Levels[0].Record.Strings["Caption"]; found && storedCaption.Text == caption {
				continue
			}
		}
		if err := applyPersistenceEdit(clone, edit, codePage, records, sites); err != nil {
			return nil, &EditError{Control: edit.Control, Property: edit.Property, Err: err}
		}
	}
	for _, level := range clone.Levels {
		fChanged := false
		if edit, changed := records[level.Record]; changed {
			raw, err := encodeEditedRecord(level.Record, &formSpec)
			if err != nil {
				return nil, &EditError{edit.Control, edit.Property, err}
			}
			level.Record.Raw = raw
			fChanged = true
		}
		var object []byte
		oChanged := false
		offset := 0
		for _, control := range level.Controls {
			size := int(control.ObjectStreamSize)
			raw := level.ORaw[offset : offset+size]
			offset += size
			if control.Level == nil {
				if edit, changed := records[control.Record]; changed {
					oChanged = true
					raw, err = encodeEditedRecord(control.Record, specsByCacheIndex[control.CLSIDCacheIndex])
					if err != nil {
						return nil, &EditError{edit.Control, edit.Property, err}
					}
					control.Site.Mask |= 1 << 5
					control.Site.Values["ObjectStreamSize"] = int64(len(raw))
					sites[control.Site] = edit
				}
			}
			object = append(object, raw...)
		}
		if len(object) > maxDesignerStreamSize {
			return nil, fmt.Errorf("%w: object stream too large", ErrInvalidEdit)
		}
		if oChanged {
			level.ORaw = object
		}
		for _, site := range level.Sites {
			if edit, changed := sites[site]; changed {
				raw, err := encodeEditedSite(site)
				if err != nil {
					return nil, &EditError{edit.Control, edit.Property, err}
				}
				site.Raw = raw
				fChanged = true
			}
		}
		if fChanged {
			level.FRaw, err = encodeEditedLevel(level)
			if err != nil {
				return nil, err
			}
		}
		stored.Streams[level.Path+"/"+level.FStreamName] = level.FRaw
		stored.Streams[level.Path+"/"+level.OStreamName] = level.ORaw
	}
	if rootCaption != nil {
		root := clone.Levels[0]
		current, found, err := rootVBFrameCaption(root.VBFrameRaw, codePage)
		if err != nil {
			return nil, &EditError{Control: "", Property: "Caption", Err: err}
		}
		if !found || current != *rootCaption {
			updated, err := rewriteVBFrameCaption(root.VBFrameRaw, *rootCaption, codePage)
			if err != nil {
				return nil, &EditError{Control: "", Property: "Caption", Err: err}
			}
			stored.Streams[root.Path+"/"+root.VBFrameName] = updated
		}
	}
	return reparseEditedForm(stored, base.Name, codePage)
}

func reparseEditedForm(stored *SerializedForm, root string, codePage uint16) (*Form, error) {
	w := cfb.NewWriter()
	for _, path := range slices.Sorted(maps.Keys(stored.Storages)) {
		w.AddStorage(splitStoragePath(path), stored.Storages[path])
	}
	for _, path := range slices.Sorted(maps.Keys(stored.Streams)) {
		w.AddStream(strings.Split(path, "/"), stored.Streams[path])
	}
	body, err := w.Bytes()
	if err != nil {
		return nil, err
	}
	c, err := cfb.Open(body)
	if err != nil {
		return nil, err
	}
	return ReadForm(c, root, codePage)
}

func applyPersistenceEdit(form *Form, edit Edit, cp uint16, records map[*Record]Edit, sites map[*Site]Edit) error {
	r := form.Levels[0].Record
	var control *Control
	if edit.Control != "" {
		for _, level := range form.Levels {
			for _, candidate := range level.Controls {
				if candidate.Name == edit.Control {
					if control != nil {
						return fmt.Errorf("%w: ambiguous control", ErrInvalidEdit)
					}
					control = candidate
				}
			}
		}
		if control == nil {
			return fmt.Errorf("%w: control not found", ErrInvalidEdit)
		}
		r = control.Record
	} else if edit.Property != "Caption" {
		return ErrUnsupportedEdit
	}
	if control != nil {
		s := control.Site
		switch edit.Property {
		case "Left", "Top":
			v, ok := edit.Value.(int32)
			if !ok {
				return ErrInvalidEdit
			}
			if s.Position == nil {
				return fmt.Errorf("%w: omitted Position default unknown", ErrUnsupportedEdit)
			}
			if edit.Property == "Left" {
				s.Position.Left = v
			} else {
				s.Position.Top = v
			}
			sites[s] = edit
			return nil
		case "TabIndex":
			v, ok := edit.Value.(int16)
			if !ok || v < 0 {
				return ErrInvalidEdit
			}
			s.Mask |= 1 << 6
			s.Values["TabIndex"] = int64(v)
			sites[s] = edit
			return nil
		case "Visible":
			v, ok := edit.Value.(bool)
			if !ok {
				return ErrInvalidEdit
			}
			flags, found := s.Values["BitFlags"]
			if !found {
				// Match the existing site projection's class-specific defaults.
				flags = 0x33
				switch control.CLSIDCacheIndex {
				case 21:
					flags = 0x32
				case 7, 14, 57:
					flags = 0x40023
				}
			}
			s.Mask |= 1 << 4
			s.Values["BitFlags"] = editedBit(flags, v)
			sites[s] = edit
			return nil
		case "Tag", "ControlTipText":
			v, ok := edit.Value.(string)
			if !ok {
				return ErrInvalidEdit
			}
			value, err := editedString(s.Strings[edit.Property], v, cp)
			if err != nil {
				return err
			}
			s.Strings[edit.Property] = value
			name, bit := "TagData", uint32(1<<1)
			if edit.Property == "ControlTipText" {
				name, bit = "ControlTipTextData", 1<<11
			}
			s.Values[name] = packedStringLength(value)
			s.Mask |= bit
			sites[s] = edit
			return nil
		}
	}
	if r == nil {
		return ErrUnsupportedEdit
	}
	spec := &formSpec
	if control != nil && control.Level == nil {
		spec = specsByCacheIndex[control.CLSIDCacheIndex]
	}
	if spec == nil {
		return ErrUnsupportedEdit
	}
	if control != nil && control.Level != nil && control.CLSIDCacheIndex != 14 {
		return ErrUnsupportedEdit
	}
	property := edit.Property
	switch property {
	case "Width", "Height":
		v, ok := edit.Value.(int32)
		if !ok || v < 0 {
			return ErrInvalidEdit
		}
		name := "Size"
		if spec == &formSpec {
			name = "DisplayedSize"
		}
		size, found := r.Sizes[name]
		if !found {
			return fmt.Errorf("%w: omitted size default unknown", ErrUnsupportedEdit)
		}
		if property == "Width" {
			size.Width = v
		} else {
			size.Height = v
		}
		r.Sizes[name] = size
		records[r] = edit
		return nil
	case "Enabled":
		v, ok := edit.Value.(bool)
		if !ok {
			return ErrInvalidEdit
		}
		if spec == &formSpec {
			// MS-OFORMS BooleanProperties bit2 is Enabled. Excel-authored
			// Issue #882 fixtures confirm Frame 0x8004 -> 0x8000 on disable.
			flags := int64(0x8004)
			if stored, found := r.Values["BooleanProperties"]; found {
				flags = stored
			}
			if v {
				flags |= 4
			} else {
				flags &^= 4
			}
			property = "BooleanProperties"
			edit.Value = flags
			break
		}
		flags, found := r.Values["VariousPropertyBits"]
		if !found {
			flags, found = DefaultVariousPropertyBits(control.CLSIDCacheIndex)
			if !found {
				return fmt.Errorf("%w: omitted VariousPropertyBits default unknown", ErrUnsupportedEdit)
			}
		}
		property = "VariousPropertyBits"
		edit.Value = editedBit(flags, v)
	case "Caption", "Value", "GroupName", "BackColor", "ForeColor", "BorderColor", "BorderStyle", "MaxLength":
	default:
		return ErrUnsupportedEdit
	}
	for _, field := range spec.data {
		if field.name != property {
			continue
		}
		if field.kind == fieldStringLength {
			v, ok := edit.Value.(string)
			if !ok {
				return ErrInvalidEdit
			}
			value, err := editedString(r.Strings[property], v, cp)
			if err != nil {
				return err
			}
			r.Strings[property] = value
			r.Values[property] = packedStringLength(value)
		} else {
			v, ok := editInteger(edit.Value)
			if !ok || v < 0 || uint64(v) > (uint64(1)<<(8*field.size))-1 {
				return ErrInvalidEdit
			}
			if property == "BorderStyle" && v > 1 {
				return ErrInvalidEdit
			}
			if property == "MaxLength" && v > math.MaxInt32 {
				return ErrInvalidEdit
			}
			r.Values[property] = v
		}
		r.Mask |= uint64(1) << field.bit
		// Keep the public address, including Enabled, for encoder failures.
		records[r] = Edit{Control: edit.Control, Property: edit.Property}
		return nil
	}
	return ErrUnsupportedEdit
}

// DefaultVariousPropertyBits is the MS-OFORMS omission default, not a snapshot
// of flags Excel may explicitly persist for a newly created control.
// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/7a72ac4a-39d9-4e2b-829e-19e3e9a1f60d
func DefaultVariousPropertyBits(cacheIndex uint16) (int64, bool) {
	switch cacheIndex {
	case 17:
		return 0x1b, true
	case 21:
		return 0x80001b, true
	case 23, 24, 25, 26, 27, 28:
		return 0x2c80081b, true
	default:
		return 0, false
	}
}

func editedBit(flags int64, value bool) int64 {
	if value {
		return flags | 2
	}
	return flags &^ 2
}

func editInteger(value any) (int64, bool) {
	if value == nil {
		return 0, false
	}
	v := reflect.ValueOf(value)
	switch v.Kind() {
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		return v.Int(), true
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		u := v.Uint()
		return int64(u), u <= math.MaxInt64
	}
	return 0, false
}

func editedString(old StoredString, text string, _ uint16) (StoredString, error) {
	if !utf8.ValidString(text) || len(text) > 3*math.MaxUint16 {
		return StoredString{}, ErrInvalidEdit
	}
	if text == old.Text {
		return old, nil
	}
	units := utf16.Encode([]rune(text))
	if old.Compressed || old.Raw == nil {
		raw := make([]byte, len(units))
		compressed := true
		for i, unit := range units {
			if unit > 0xff {
				compressed = false
				break
			}
			raw[i] = byte(unit)
		}
		if compressed {
			return StoredString{Text: text, Raw: raw, Compressed: true}, nil
		}
	}
	raw := make([]byte, len(units)*2)
	for i, unit := range units {
		binary.LittleEndian.PutUint16(raw[i*2:], unit)
	}
	return StoredString{Text: text, Raw: raw}, nil
}

func packedStringLength(value StoredString) int64 {
	n := uint32(len(value.Raw))
	if value.Compressed {
		n |= 0x80000000
	}
	return int64(n)
}

func appendEditInteger(b []byte, value int64, size int) []byte {
	switch size {
	case 1:
		return append(b, byte(value))
	case 2:
		return binary.LittleEndian.AppendUint16(b, uint16(value))
	default:
		return binary.LittleEndian.AppendUint32(b, uint32(value))
	}
}

func appendEditPadding(b []byte, size int, padding []byte) []byte {
	n := (size - len(b)%size) % size
	if len(padding) == n {
		return append(b, padding...)
	}
	return append(b, make([]byte, n)...)
}

func finishEditHeader(b []byte) ([]byte, error) {
	if len(b)-4 > math.MaxUint16 {
		return nil, fmt.Errorf("%w: record exceeds uint16 length", ErrInvalidEdit)
	}
	binary.LittleEndian.PutUint16(b[2:4], uint16(len(b)-4))
	return b, nil
}

func encodeEditedRecord(r *Record, spec *recordSpec) ([]byte, error) {
	if spec == nil {
		return nil, ErrUnsupportedEdit
	}
	var known uint64
	for _, f := range spec.data {
		known |= uint64(1) << f.bit
	}
	for _, f := range spec.extra {
		known |= uint64(1) << f.bit
	}
	for bit := range spec.flags {
		known |= uint64(1) << bit
	}
	if spec == &morphDataSpec {
		// MS-OFORMS 2.2.5.2 bit31 is Reserved: set to 1 and ignored;
		// it has no DataBlock field. Keep its original value losslessly.
		// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/19014e19-67fa-4060-8c81-d1463809f117
		known |= 1 << 31
	}
	if r.Mask&^known != 0 {
		return nil, fmt.Errorf("%w: unknown %s mask bits %#x", ErrUnsupportedEdit, r.Type, r.Mask&^known)
	}
	b := []byte{r.Minor, r.Major, 0, 0}
	b = binary.LittleEndian.AppendUint32(b, uint32(r.Mask))
	if spec.mask64 {
		b = binary.LittleEndian.AppendUint32(b, uint32(r.Mask>>32))
	}
	for _, field := range spec.data {
		if r.Mask&(uint64(1)<<field.bit) == 0 {
			continue
		}
		b = appendEditPadding(b, field.size, r.Padding["before:"+field.name])
		b = appendEditInteger(b, r.Values[field.name], field.size)
	}
	b = appendEditPadding(b, 4, r.Padding["data:end"])
	for _, extra := range spec.extra {
		if r.Mask&(uint64(1)<<extra.bit) == 0 {
			continue
		}
		switch extra.kind {
		case extraSize, extraPosition:
			size := r.Sizes[extra.name]
			b = binary.LittleEndian.AppendUint32(b, uint32(size.Width))
			b = binary.LittleEndian.AppendUint32(b, uint32(size.Height))
		case extraString:
			b = append(b, r.Strings[extra.name].Raw...)
			b = appendEditPadding(b, 4, r.Padding["str:"+extra.name])
		case extraArray:
			b = append(b, r.Arrays[extra.name]...)
		}
	}
	b, err := finishEditHeader(b)
	if err != nil {
		return nil, err
	}
	if spec.stopAfterExtra {
		return b, nil
	}
	for _, stream := range spec.stream {
		if r.Mask&(uint64(1)<<stream.bit) != 0 {
			b = append(b, r.Pictures[stream.name]...)
		}
	}
	if spec.textProps {
		if r.TextProps == nil {
			return nil, fmt.Errorf("%w: missing TextProps", ErrInvalidEdit)
		}
		b = append(b, r.TextProps.Raw...)
	}
	b = append(b, r.TailRaw...)
	return b, nil
}

func encodeEditedSite(s *Site) ([]byte, error) {
	known := uint32(1 << 8)
	for _, f := range siteFields {
		known |= uint32(1) << f.bit
	}
	if s.Mask&^known != 0 {
		return nil, fmt.Errorf("%w: unknown site mask bits %#x", ErrUnsupportedEdit, s.Mask&^known)
	}
	b := binary.LittleEndian.AppendUint16(nil, s.Version)
	b = append(b, 0, 0)
	b = binary.LittleEndian.AppendUint32(b, s.Mask)
	for _, field := range siteFields {
		if s.Mask&(uint32(1)<<field.bit) == 0 {
			continue
		}
		b = appendEditPadding(b, field.size, s.Padding["before:"+field.name])
		b = appendEditInteger(b, s.Values[field.name], field.size)
	}
	b = appendEditPadding(b, 4, s.Padding["data:end"])
	appendStrings := func(items [][2]string) {
		for _, item := range items {
			if _, found := s.Values[item[0]]; found {
				b = append(b, s.Strings[item[1]].Raw...)
				b = appendEditPadding(b, 4, s.Padding["str:"+item[1]])
			}
		}
	}
	appendStrings(siteStringsBeforePosition)
	if s.Mask&(1<<8) != 0 {
		if s.Position == nil {
			return nil, ErrInvalidEdit
		}
		b = binary.LittleEndian.AppendUint32(b, uint32(s.Position.Left))
		b = binary.LittleEndian.AppendUint32(b, uint32(s.Position.Top))
	}
	appendStrings(siteStringsAfterPosition)
	return finishEditHeader(b)
}

func encodeEditedLevel(level *Level) ([]byte, error) {
	b := bytes.Clone(level.Record.Raw)
	b = append(b, level.MouseIconRaw...)
	b = append(b, level.FontRaw...)
	b = append(b, level.PictureRaw...)
	// An omitted class table has zero bytes; a stored empty table has two.
	b = append(b, level.ClassTableRaw...)
	b = binary.LittleEndian.AppendUint32(b, uint32(len(level.Sites)))
	countOffset := len(b)
	b = append(b, 0, 0, 0, 0)
	start := len(b)
	b = append(b, level.DepthsRaw...)
	for _, site := range level.Sites {
		b = append(b, site.Raw...)
	}
	binary.LittleEndian.PutUint32(b[countOffset:], uint32(len(b)-start))
	b = append(b, level.TrailingRaw...)
	if len(b) > maxDesignerStreamSize {
		return nil, fmt.Errorf("%w: form stream too large", ErrInvalidEdit)
	}
	return b, nil
}
