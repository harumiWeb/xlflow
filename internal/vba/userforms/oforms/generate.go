package oforms

import (
	"encoding/binary"
	"fmt"
	"maps"
	"math"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
)

// Definition is authoring input in persistence units, separate from FormSpec.
// Properties use persistence field names and strings or int64 values.
type Definition struct {
	Name, Caption string
	Size          Size
	Controls      []ControlDefinition
}

type ControlDefinition struct {
	Name       string
	Class      uint16
	Size       Size
	Position   Position
	TabIndex   int16
	Visible    bool
	Properties map[string]any
}

// NewForm authors a flat built-in Designer and returns a read-time signed model.
// It never borrows a template, starts Excel, or accepts arbitrary raw bytes.
func NewForm(input Definition, codePage uint16) (*Form, error) {
	if input.Name == "" || strings.ContainsAny(input.Name, "\\/\x00\r\n\"{}") || strings.EqualFold(input.Name, "VBA") {
		return nil, fmt.Errorf("%w: invalid form name", ErrInvalidEdit)
	}
	if input.Size.Width < 0 || input.Size.Height < 0 || len(input.Controls) > math.MaxInt16+1 {
		return nil, fmt.Errorf("%w: invalid size or site count", ErrInvalidEdit)
	}
	root := newGenerationRecord(&formSpec)
	root.Mask = 1<<3 | 1<<10 | 1<<11 | 1<<26 | 1<<27
	root.Values["NextAvailableID"] = int64(len(input.Controls) + 1)
	root.Values["ShapeCookie"] = int64(len(input.Controls))
	root.Values["DrawBuffer"] = 32000
	root.Sizes["DisplayedSize"] = input.Size
	root.Sizes["LogicalSize"] = Size{}
	if err := setGenerationProperty(root, &formSpec, "Caption", input.Caption); err != nil {
		return nil, err
	}
	var err error
	root.Raw, err = encodeEditedRecord(root, &formSpec)
	if err != nil {
		return nil, err
	}
	level := &Level{Record: root, ClassTableRaw: []byte{0, 0}}
	siteBytes := len(root.Raw) + 10
	names := map[string]bool{}
	for i, control := range input.Controls {
		if control.Name == "" || strings.ContainsAny(control.Name, "\\/\x00\r\n\"") || names[strings.ToLower(control.Name)] || control.TabIndex < 0 || control.Size.Width < 0 || control.Size.Height < 0 {
			return nil, fmt.Errorf("%w: invalid control %q", ErrInvalidEdit, control.Name)
		}
		names[strings.ToLower(control.Name)] = true
		table := specsByCacheIndex[control.Class]
		if table == nil || control.Class == 15 || control.Class == 18 {
			return nil, fmt.Errorf("%w: class %d", ErrUnsupportedEdit, control.Class)
		}
		record := newGenerationRecord(table)
		for _, field := range table.extra {
			if field.name == "Size" {
				record.Mask |= uint64(1) << field.bit
				record.Sizes["Size"] = control.Size
			}
		}
		if table == &morphDataSpec {
			record.Mask |= 1 << 31
			style := map[uint16]int64{23: 1, 24: 2, 25: 3, 26: 4, 27: 5, 28: 6}[control.Class]
			if err := setGenerationProperty(record, table, "DisplayStyle", style); err != nil {
				return nil, err
			}
		}
		if table.textProps {
			record.TextProps = newGenerationRecord(&textPropsSpec)
			record.TextProps.Raw, err = encodeEditedRecord(record.TextProps, &textPropsSpec)
			if err != nil {
				return nil, err
			}
		}
		for _, name := range slices.Sorted(maps.Keys(control.Properties)) {
			if name == "Tag" || name == "ControlTipText" {
				continue
			}
			if err := setGenerationProperty(record, table, name, control.Properties[name]); err != nil {
				return nil, fmt.Errorf("control %q: %w", control.Name, err)
			}
		}
		record.Raw, err = encodeEditedRecord(record, table)
		if err != nil {
			return nil, err
		}
		siteName, err := editedString(StoredString{}, control.Name, codePage)
		if err != nil {
			return nil, err
		}
		flags := int64(0x33)
		if control.Class == 21 {
			flags = 0x32
		}
		if !control.Visible {
			flags &^= 2
		}
		site := &Site{Mask: 1<<0 | 1<<2 | 1<<4 | 1<<5 | 1<<6 | 1<<7 | 1<<8,
			Values:  map[string]int64{"NameData": packedStringLength(siteName), "ID": int64(i + 1), "BitFlags": flags, "ObjectStreamSize": int64(len(record.Raw)), "TabIndex": int64(control.TabIndex), "ClsidCacheIndex": int64(control.Class)},
			Strings: map[string]StoredString{"Name": siteName}, Position: &control.Position}
		for _, property := range []struct {
			name, length string
			bit          uint32
		}{{"Tag", "TagData", 1 << 1}, {"ControlTipText", "ControlTipTextData", 1 << 11}} {
			if value, found := control.Properties[property.name]; found {
				text, ok := value.(string)
				if !ok {
					return nil, ErrInvalidEdit
				}
				stored, err := editedString(StoredString{}, text, codePage)
				if err != nil {
					return nil, err
				}
				site.Mask |= property.bit
				site.Values[property.length] = packedStringLength(stored)
				site.Strings[property.name] = stored
			}
		}
		site.Raw, err = encodeEditedSite(site)
		if err != nil {
			return nil, err
		}
		if len(site.Raw)+4 > maxDesignerStreamSize-siteBytes {
			return nil, fmt.Errorf("%w: site stream exceeds limit", ErrInvalidEdit)
		}
		siteBytes += len(site.Raw) + 4
		level.Sites = append(level.Sites, site)
		if len(record.Raw) > maxDesignerStreamSize-len(level.ORaw) {
			return nil, fmt.Errorf("%w: object stream exceeds limit", ErrInvalidEdit)
		}
		level.ORaw = append(level.ORaw, record.Raw...)
		level.DepthsRaw = append(level.DepthsRaw, 0, 1)
	}
	level.DepthsRaw = appendEditPadding(level.DepthsRaw, 4, nil)
	f, err := encodeEditedLevel(level)
	if err != nil {
		return nil, err
	}
	frameBytes, err := generatedVBFrame(input.Name, input.Caption, input.Size, codePage)
	if err != nil {
		return nil, &EditError{Control: "", Property: "Caption", Err: err}
	}
	stored := &SerializedForm{Storages: map[string]cfb.StorageMeta{input.Name: {}}, Streams: map[string][]byte{
		input.Name + "/f": f, input.Name + "/o": level.ORaw, input.Name + "/\x01CompObj": generationCompObj(), input.Name + "/\x03VBFrame": frameBytes,
	}}
	return reparseEditedForm(stored, input.Name, codePage)
}

func newGenerationRecord(table *recordSpec) *Record {
	return &Record{Type: table.typeName, Major: table.major, Values: map[string]int64{}, Strings: map[string]StoredString{}, Sizes: map[string]Size{}}
}

func setGenerationProperty(record *Record, table *recordSpec, name string, value any) error {
	for _, field := range table.data {
		if field.name != name {
			continue
		}
		if field.kind == fieldMarker {
			return fmt.Errorf("%w: resource %s", ErrUnsupportedEdit, name)
		}
		if field.kind == fieldStringLength {
			text, ok := value.(string)
			if !ok {
				return fmt.Errorf("%w: %s requires string", ErrInvalidEdit, name)
			}
			stored, err := editedString(StoredString{}, text, 0)
			if err != nil {
				return err
			}
			record.Strings[name] = stored
			record.Values[name] = packedStringLength(stored)
		} else {
			n, ok := value.(int64)
			if !ok {
				return fmt.Errorf("%w: %s requires int64", ErrInvalidEdit, name)
			}
			if field.kind == fieldSigned {
				if n < -(int64(1)<<(field.size*8-1)) || n > (int64(1)<<(field.size*8-1))-1 {
					return ErrInvalidEdit
				}
			} else if n < 0 || uint64(n) > (uint64(1)<<(field.size*8))-1 {
				return ErrInvalidEdit
			}
			record.Values[name] = n
		}
		record.Mask |= uint64(1) << field.bit
		return nil
	}
	return fmt.Errorf("%w: property %s", ErrUnsupportedEdit, name)
}

func generationCompObj() []byte {
	header := make([]byte, 28)
	binary.LittleEndian.PutUint32(header, 0xfffe0001)
	binary.LittleEndian.PutUint32(header[4:], 0xa03)
	binary.LittleEndian.PutUint32(header[8:], math.MaxUint32)
	for _, text := range []string{"Microsoft Forms 2.0 Form\x00", "Embedded Object\x00"} {
		header = binary.LittleEndian.AppendUint32(header, uint32(len(text)))
		header = append(header, text...)
	}
	header = binary.LittleEndian.AppendUint32(header, 0)
	header = binary.LittleEndian.AppendUint32(header, 0x71b239f4)
	return append(header, make([]byte, 12)...)
}
