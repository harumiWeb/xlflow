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
	Controls   []ControlDefinition
	Tabs       *TabStrip
	Internal   bool
}

// NewForm authors a built-in Designer and returns a read-time signed model.
// It never borrows a template, starts Excel, or accepts arbitrary raw bytes.
func NewForm(input Definition, codePage uint16) (*Form, error) {
	if input.Name == "" || strings.ContainsAny(input.Name, "\\/\x00\r\n\"{}") || strings.EqualFold(input.Name, "VBA") {
		return nil, fmt.Errorf("%w: invalid form name", ErrInvalidEdit)
	}
	if input.Size.Width < 0 || input.Size.Height < 0 || len(input.Controls) > math.MaxInt16+1 {
		return nil, fmt.Errorf("%w: invalid size or site count", ErrInvalidEdit)
	}
	frameBytes, err := generatedVBFrame(input.Name, input.Caption, input.Size, codePage)
	if err != nil {
		return nil, &EditError{Control: "", Property: "Caption", Err: err}
	}
	stored := &SerializedForm{Storages: map[string]cfb.StorageMeta{}, Streams: map[string][]byte{input.Name + "/\x03VBFrame": frameBytes}}
	state := generationState{stored: stored, codePage: codePage, names: map[string]bool{}, nextID: 1}
	if err := state.level(input.Name, input.Caption, input.Size, nil, input.Controls, 0, 0, nil); err != nil {
		return nil, err
	}
	return reparseEditedForm(stored, input.Name, codePage)
}

type generationState struct {
	stored   *SerializedForm
	codePage uint16
	names    map[string]bool
	nextID   int64
}

// MaxNestingDepth includes the root form level and matches the reader's limit.
const MaxNestingDepth = maxNestingDepth

func (g *generationState) level(path, caption string, size Size, properties map[string]any, controls []ControlDefinition, depth int, class uint16, tabs *TabStrip) error {
	if depth >= maxNestingDepth || len(controls) > math.MaxInt16+1 {
		return fmt.Errorf("%w: nesting or site limit", ErrInvalidEdit)
	}
	codePage := g.codePage
	firstID := g.nextID
	root := newGenerationRecord(&formSpec)
	root.Mask = 1<<3 | 1<<10 | 1<<11 | 1<<26 | 1<<27
	root.Values["NextAvailableID"] = 1
	root.Values["ShapeCookie"] = int64(len(controls))
	root.Values["DrawBuffer"] = 32000
	root.Sizes["DisplayedSize"] = size
	root.Sizes["LogicalSize"] = Size{}
	if depth > 0 {
		root.Mask |= 1 << 6
		root.Values["BooleanProperties"] = 0x8004
		if class == 14 {
			root.Mask |= 1 << 17
			root.Values["SpecialEffect"] = 3
		}
		if class == 57 {
			root.Values["BooleanProperties"] = 0xc004
		}
	}
	if err := setGenerationProperty(root, &formSpec, "Caption", caption); err != nil {
		return err
	}
	for _, name := range slices.Sorted(maps.Keys(properties)) {
		if name == "Tag" || name == "ControlTipText" || name == "Accelerator" || name == "MultiPageEnabled" || class == 7 && (name == "Caption" || name == "BooleanProperties") {
			continue
		}
		if err := setGenerationProperty(root, &formSpec, name, properties[name]); err != nil {
			return err
		}
	}
	var err error
	root.Raw, err = encodeEditedRecord(root, &formSpec)
	if err != nil {
		return err
	}
	level := &Level{Record: root, ClassTableRaw: []byte{0, 0}}
	if root.Values["BooleanProperties"]&0x8000 != 0 {
		level.ClassTableRaw = nil
	}
	if class == 57 {
		if tabs == nil {
			tabs = &TabStrip{SelectedIndex: 0}
			if len(controls) == 0 {
				tabs.SelectedIndex = -1
			}
		}
		if len(tabs.Tabs) != len(controls) {
			return fmt.Errorf("%w: MultiPage tabs/pages count", ErrInvalidEdit)
		}
		for i := range controls {
			if controls[i].Class != 7 {
				return fmt.Errorf("%w: MultiPage requires Page children", ErrInvalidEdit)
			}
			controls[i].Position, controls[i].Size = generatedPageBox(size)
			controls[i].Visible = int32(i) == tabs.SelectedIndex
		}
		controls = append([]ControlDefinition{{Class: 18, Internal: true, Visible: true, Size: size, Tabs: tabs}}, controls...)
		level.TrailingRaw = []byte{0, 2, 12, 0, 25, 0, 0, 0, 252, 143, 0, 0, 255, 1, 0, 0}
	}
	siteBytes := len(root.Raw) + 10
	for _, control := range controls {
		if (!control.Internal && (control.Name == "" || g.names[strings.ToLower(control.Name)])) || strings.ContainsAny(control.Name, "\\/\x00\r\n\"") || control.TabIndex < 0 || control.Size.Width < 0 || control.Size.Height < 0 {
			return fmt.Errorf("%w: invalid control %q", ErrInvalidEdit, control.Name)
		}
		if !control.Internal {
			g.names[strings.ToLower(control.Name)] = true
		}
		id := g.nextID
		g.nextID++
		table := specsByCacheIndex[control.Class]
		if control.Class == 14 || control.Class == 7 || control.Class == 57 {
			table = &formSpec
		}
		if table == nil || control.Class == 15 || table != &formSpec && len(control.Controls) > 0 {
			return fmt.Errorf("%w: class %d", ErrUnsupportedEdit, control.Class)
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
				return err
			}
		}
		if table.textProps {
			record.TextProps = newGenerationRecord(&textPropsSpec)
			record.TextProps.Raw, err = encodeEditedRecord(record.TextProps, &textPropsSpec)
			if err != nil {
				return err
			}
		}
		if control.Class == 18 {
			record.Mask |= 1 << 19
			if control.Tabs == nil {
				control.Tabs = &TabStrip{SelectedIndex: -1}
			}
			if err := SetTabStrip(record, control.Tabs); err != nil {
				return err
			}
			// A newly authored empty strip must persist -1 explicitly. Omission
			// has the wire default 0, unlike the projected empty selection.
			if len(control.Tabs.Tabs) == 0 {
				record.Mask |= 1
				record.Values["ListIndex"] = -1
			}
			font := record.TextProps
			if err := setGenerationProperty(font, &textPropsSpec, "FontName", "Tahoma"); err != nil {
				return err
			}
			if err := setGenerationProperty(font, &textPropsSpec, "FontHeight", int64(165)); err != nil {
				return err
			}
			font.Raw, err = encodeEditedRecord(font, &textPropsSpec)
			if err != nil {
				return err
			}
		}
		for _, name := range slices.Sorted(maps.Keys(control.Properties)) {
			if name == "Tag" || name == "ControlTipText" || name == "Accelerator" || name == "MultiPageEnabled" || control.Class == 7 && (name == "Caption" || name == "BooleanProperties") {
				continue
			}
			if err := setGenerationProperty(record, table, name, control.Properties[name]); err != nil {
				return fmt.Errorf("control %q: %w", control.Name, err)
			}
		}
		record.Raw, err = encodeEditedRecord(record, table)
		if err != nil {
			return err
		}
		if table == &formSpec {
			if err := g.level(fmt.Sprintf("%s/i%02d", path, id), "", control.Size, control.Properties, control.Controls, depth+1, control.Class, control.Tabs); err != nil {
				return err
			}
			record.Raw = nil
		}
		siteName, err := editedString(StoredString{}, control.Name, codePage)
		if err != nil {
			return err
		}
		flags := int64(0x33)
		if control.Class == 21 {
			flags = 0x32
		}
		if table == &formSpec {
			flags = 0x40023
		}
		if !control.Visible {
			flags &^= 2
		}
		site := &Site{Mask: 1<<0 | 1<<2 | 1<<4 | 1<<5 | 1<<6 | 1<<7 | 1<<8,
			Values:  map[string]int64{"NameData": packedStringLength(siteName), "ID": id, "BitFlags": flags, "ObjectStreamSize": int64(len(record.Raw)), "TabIndex": int64(control.TabIndex), "ClsidCacheIndex": int64(control.Class)},
			Strings: map[string]StoredString{"Name": siteName}, Position: &control.Position}
		if table == &formSpec {
			site.Mask &^= 1 << 5
			delete(site.Values, "ObjectStreamSize")
		}
		for _, property := range []struct {
			name, length string
			bit          uint32
		}{{"Tag", "TagData", 1 << 1}, {"ControlTipText", "ControlTipTextData", 1 << 11}} {
			if value, found := control.Properties[property.name]; found {
				text, ok := value.(string)
				if !ok {
					return ErrInvalidEdit
				}
				stored, err := editedString(StoredString{}, text, codePage)
				if err != nil {
					return err
				}
				site.Mask |= property.bit
				site.Values[property.length] = packedStringLength(stored)
				site.Strings[property.name] = stored
			}
		}
		site.Raw, err = encodeEditedSite(site)
		if err != nil {
			return err
		}
		if len(site.Raw)+4 > maxDesignerStreamSize-siteBytes {
			return fmt.Errorf("%w: site stream exceeds limit", ErrInvalidEdit)
		}
		siteBytes += len(site.Raw) + 4
		level.Sites = append(level.Sites, site)
		level.Controls = append(level.Controls, &Control{ID: int32(id), Site: site, Name: control.Name, CLSIDCacheIndex: control.Class, Record: record})
		if len(record.Raw) > maxDesignerStreamSize-len(level.ORaw) {
			return fmt.Errorf("%w: object stream exceeds limit", ErrInvalidEdit)
		}
		level.ORaw = append(level.ORaw, record.Raw...)
		level.DepthsRaw = append(level.DepthsRaw, 0, 1)
	}
	level.DepthsRaw = appendEditPadding(level.DepthsRaw, 4, nil)
	if g.nextID > maxSitesPerForm+1 {
		return fmt.Errorf("%w: total site limit", ErrInvalidEdit)
	}
	root.Values["NextAvailableID"] = g.nextID
	root.Values["ShapeCookie"] = g.nextID - firstID
	root.Raw, err = encodeEditedRecord(root, &formSpec)
	if err != nil {
		return err
	}
	f, err := encodeEditedLevel(level)
	if err != nil {
		return err
	}
	g.stored.Storages[path] = cfb.StorageMeta{}
	g.stored.Streams[path+"/f"] = f
	g.stored.Streams[path+"/o"] = level.ORaw
	g.stored.Streams[path+"/\x01CompObj"] = generationCompObj()
	if depth > 0 {
		identity, progID := frameCLSID, "Forms.Frame.1"
		if class == 7 {
			identity, progID = pageCLSID, "Forms.Form.1"
		}
		if class == 57 {
			identity, progID = multiPageCLSID, "Forms.MultiPage.1"
		}
		g.stored.Storages[path] = cfb.StorageMeta{CLSID: identity}
		g.stored.Streams[path+"/\x01CompObj"] = generationContainerCompObj(identity, progID)
	}
	if class == 57 {
		state := &MultiPage{Hidden: level.Controls[0], Pages: level.Controls[1:], PageProperties: map[int32][]byte{}}
		// Find the enclosing definition's enable state passed as a property.
		if enabled, ok := properties["MultiPageEnabled"].(bool); ok && !enabled {
			state.Properties = newGenerationRecord(&multiPagePropertiesSpec)
			state.Properties.Mask |= 1 << 3
		}
		x, err := encodeMultiPageX(state)
		if err != nil {
			return err
		}
		g.stored.Streams[path+"/x"] = x
	}
	return nil
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

// Excel-authored Frames bind both directory and CompObj identities to this
// CLSID; root UserForms use the existing zero-CLSID CompObj layout.
var frameCLSID = [16]byte{0x20, 0x20, 0x18, 0x6e, 0x60, 0xf4, 0xce, 0x11, 0x9b, 0xcd, 0x00, 0xaa, 0x00, 0x60, 0x8e, 0x01}
var multiPageCLSID = [16]byte{0x70, 0x13, 0xe3, 0x46, 0x7a, 0x3f, 0xce, 0x11, 0xbe, 0xd6, 0, 0xaa, 0, 0x61, 0x10, 0x80}
var pageCLSID = [16]byte{0xf0, 0x69, 0x2a, 0xc6, 0xdc, 0x16, 0xce, 0x11, 0x9e, 0x98, 0, 0xaa, 0, 0x57, 0x4a, 0x4f}

func generationContainerCompObj(identity [16]byte, progID string) []byte {
	b := generationCompObj()[:28]
	copy(b[12:28], identity[:])
	userType := "Microsoft Forms 2.0 Form\x00"
	if identity == frameCLSID {
		userType = "Microsoft Forms 2.0 Frame\x00"
	}
	for _, text := range []string{userType, "Embedded Object\x00", progID + "\x00"} {
		b = binary.LittleEndian.AppendUint32(b, uint32(len(text)))
		b = append(b, text...)
	}
	b = binary.LittleEndian.AppendUint32(b, 0x71b239f4)
	return append(b, make([]byte, 12)...)
}
