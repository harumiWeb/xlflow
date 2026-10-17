package oforms

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"math"
)

// Font is a partial persistence edit. Nil members leave stored state untouched.
// Wire definitions: MS-OFORMS 2.4.12 (StdFont), 2.5.26.1 (fmFontEffects),
// and 2.5.27 (FontHeight). StdFont is ASCII and uses 1/10000 point;
// TextProps is Unicode and uses twips. Neither representation uses machine DPI.
// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/e219ce61-420f-403b-8216-63890c935aa1
// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/394a188f-5071-4e7a-8dad-986d1989fd7b
// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/012f7562-a80c-4b5b-96c0-8b9c36319b34
type Font struct {
	Name                                   *string
	Size                                   *float64
	Bold, Italic, Underline, StrikeThrough *bool
	Charset                                *int
}

var textPropsCLSID = []byte{0x20, 0x09, 0xc2, 0xaf, 0x4e, 0xda, 0xce, 0x11, 0xb9, 0x43, 0x00, 0xaa, 0x00, 0x68, 0x87, 0xb4}

// ReadLevelFont projects either MS-OFORMS StdFont or TextProps without resolving
// machine fonts. Unknown and malformed representations are never replaced.
func ReadLevelFont(level *Level) (*Font, error) {
	if level == nil || len(level.FontRaw) == 0 {
		return nil, nil
	}
	raw := level.FontRaw
	if len(raw) < 16 {
		return nil, ErrUnsupportedEdit
	}
	switch binary.LittleEndian.Uint32(raw) {
	case guidStdFont:
		if len(raw) < 27 || raw[16] != 1 || len(raw) != 27+int(raw[26]) {
			return nil, ErrUnsupportedEdit
		}
		for _, b := range raw[27:] {
			if b > 127 {
				return nil, ErrUnsupportedEdit
			}
		}
		flags := raw[19]
		return &Font{Name: new(string(raw[27:])), Size: new(float64(binary.LittleEndian.Uint32(raw[22:])) / 10000), Charset: new(int(int16(binary.LittleEndian.Uint16(raw[17:])))), Bold: new(binary.LittleEndian.Uint16(raw[20:]) > 550), Italic: new(flags&2 != 0), Underline: new(flags&4 != 0), StrikeThrough: new(flags&8 != 0)}, nil
	case guidTextProps:
		record, n, err := parseRecord(raw[16:], &textPropsSpec, 1252, len(raw)-16)
		if err != nil || n != len(raw)-16 {
			return nil, ErrUnsupportedEdit
		}
		return ReadTextFont(record), nil
	default:
		return nil, ErrUnsupportedEdit
	}
}

// ReadTextFont retains omission rather than materializing rendering defaults.
func ReadTextFont(record *Record) *Font {
	if record == nil {
		return nil
	}
	f := &Font{}
	if v, ok := record.Strings["FontName"]; ok {
		f.Name = new(v.Text)
	}
	if v, ok := record.Values["FontHeight"]; ok {
		f.Size = new(float64(v) / 20)
	}
	if v, ok := record.Values["FontCharSet"]; ok {
		f.Charset = new(int(v))
	}
	if v, ok := record.Values["FontEffects"]; ok {
		// Omitted FontWeight has the MS-OFORMS default 400. Excel does not
		// turn that default into Bold merely because FontEffects bit0 is set.
		f.Bold, f.Italic, f.Underline, f.StrikeThrough = new(false), new(v&2 != 0), new(v&4 != 0), new(v&8 != 0)
	}
	// Excel uses the persisted weight for Bold when present, even if the
	// FontEffects bold bit disagrees. IFont::put_Weight defines Bold as >550.
	// https://learn.microsoft.com/en-us/windows/win32/api/ocidl/nf-ocidl-ifont-put_weight
	// Do not misreport a self-encoded bit as
	// evidence that Excel will load a bold font.
	if weight, ok := record.Values["FontWeight"]; ok {
		f.Bold = new(weight > 550)
	}
	if f.Name == nil && f.Size == nil && f.Charset == nil && f.Bold == nil {
		return nil
	}
	return f
}

// ReadControlFont handles the hidden TabStrip that owns MultiPage font state.
func ReadControlFont(control *Control) (*Font, error) {
	if control == nil {
		return nil, nil
	}
	if control.MultiPage != nil {
		return ReadControlFont(control.MultiPage.Hidden)
	}
	if control.Level != nil {
		return ReadLevelFont(control.Level)
	}
	if control.Record == nil {
		return nil, nil
	}
	return ReadTextFont(control.Record.TextProps), nil
}

func editTextFont(record *Record, font Font, cp uint16) error {
	if record == nil {
		return ErrUnsupportedEdit
	}
	if font.Name != nil {
		if *font.Name == "" {
			return ErrInvalidEdit
		}
		if err := setGenerationProperty(record, &textPropsSpec, "FontName", *font.Name); err != nil {
			return err
		}
	}
	if font.Size != nil {
		v := *font.Size * 20
		if math.IsNaN(v) || math.IsInf(v, 0) || v <= 0 || v > 4294967 || math.Abs(v-math.Round(v)) > 1e-7 {
			return fmt.Errorf("%w: font size requires positive whole twips", ErrInvalidEdit)
		}
		if err := setGenerationProperty(record, &textPropsSpec, "FontHeight", int64(math.Round(v))); err != nil {
			return err
		}
	}
	if font.Charset != nil {
		if *font.Charset < 0 || *font.Charset > 255 {
			return ErrInvalidEdit
		}
		if err := setGenerationProperty(record, &textPropsSpec, "FontCharSet", int64(*font.Charset)); err != nil {
			return err
		}
	}
	for bit, value := range []*bool{font.Bold, font.Italic, font.Underline, font.StrikeThrough} {
		if value == nil {
			continue
		}
		flags := record.Values["FontEffects"]
		if *value {
			flags |= 1 << bit
		} else {
			flags &^= 1 << bit
		}
		if err := setGenerationProperty(record, &textPropsSpec, "FontEffects", flags); err != nil {
			return err
		}
	}
	if font.Bold != nil {
		weight := int64(400)
		if *font.Bold {
			weight = 700
		}
		if err := setGenerationProperty(record, &textPropsSpec, "FontWeight", weight); err != nil {
			return err
		}
	}
	raw, err := encodeEditedRecord(record, &textPropsSpec)
	if err != nil {
		return err
	}
	record.Raw = raw
	return nil
}

func editLevelFont(level *Level, font Font, cp uint16) error {
	if level == nil {
		return ErrUnsupportedEdit
	}
	if len(level.FontRaw) == 0 {
		level.FontRaw = append(bytes.Clone(textPropsCLSID), []byte{0, 2, 4, 0, 0, 0, 0, 0}...)
	}
	if _, err := ReadLevelFont(level); err != nil {
		return err
	}
	raw := level.FontRaw
	if binary.LittleEndian.Uint32(raw) == guidStdFont {
		result := bytes.Clone(raw)
		if font.Name != nil {
			name := *font.Name
			if len(name) == 0 || len(name) > 31 {
				return ErrInvalidEdit
			}
			for _, b := range []byte(name) {
				if b > 127 || b == 0 {
					return fmt.Errorf("%w: StdFont requires ASCII face name", ErrUnsupportedEdit)
				}
			}
			result = append(result[:27], name...)
			result[26] = byte(len(name))
		}
		if font.Size != nil {
			n := *font.Size * 10000
			if math.IsNaN(n) || math.IsInf(n, 0) || n <= 0 || n > 655350000 || math.Abs(n-math.Round(n)) > 1e-6 {
				return ErrInvalidEdit
			}
			binary.LittleEndian.PutUint32(result[22:], uint32(math.Round(n)))
		}
		if font.Charset != nil {
			if *font.Charset < 0 || *font.Charset > 255 {
				return ErrInvalidEdit
			}
			binary.LittleEndian.PutUint16(result[17:], uint16(*font.Charset))
		}
		if font.Bold != nil {
			weight := uint16(400)
			if *font.Bold {
				weight = 700
			}
			binary.LittleEndian.PutUint16(result[20:], weight)
		}
		for i, v := range []*bool{font.Italic, font.Underline, font.StrikeThrough} {
			if v != nil {
				bit := byte(2 << i)
				if *v {
					result[19] |= bit
				} else {
					result[19] &^= bit
				}
			}
		}
		level.FontRaw = result
	} else {
		r, _, err := parseRecord(raw[16:], &textPropsSpec, cp, len(raw)-16)
		if err != nil {
			return err
		}
		if err = editTextFont(r, font, cp); err != nil {
			return err
		}
		level.FontRaw = append(bytes.Clone(raw[:16]), r.Raw...)
	}
	level.Record.Mask |= formMaskFont
	level.Record.Values["Font"] = math.MaxUint16
	return nil
}

func applyFontEdit(form *Form, control *Control, edit Edit, cp uint16, records map[*Record]Edit) error {
	font, ok := edit.Value.(Font)
	if !ok {
		return ErrInvalidEdit
	}
	level := form.Levels[0]
	if control != nil {
		if control.MultiPage != nil {
			control = control.MultiPage.Hidden
		}
		level = control.Level
	}
	if level != nil {
		if err := editLevelFont(level, font, cp); err != nil {
			return err
		}
		records[level.Record] = edit
		return nil
	}
	if control == nil || control.Record == nil || control.Record.TextProps == nil {
		return ErrUnsupportedEdit
	}
	if err := editTextFont(control.Record.TextProps, font, cp); err != nil {
		return err
	}
	records[control.Record] = edit
	return nil
}
