package compiler

import (
	"encoding/json"
	"fmt"
	"math"
	"reflect"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

// promoteLegacyAppearance affects the owned compiler snapshot only. It neither
// migrates authoring files nor changes unknown legacy properties.
func promoteLegacyAppearance(control *spec.FormSpecControl) {
	for key, raw := range control.Properties {
		switch strings.ToLower(strings.TrimSpace(key)) {
		case "font":
			if control.Font != nil {
				continue
			}
			body, err := json.Marshal(raw)
			if err != nil {
				continue
			}
			var members map[string]json.RawMessage
			if json.Unmarshal(body, &members) != nil || members == nil {
				continue
			}
			known := true
			for member := range members {
				switch member {
				case "name", "size", "bold", "italic", "underline", "strikeThrough", "charset":
				default:
					known = false
				}
			}
			if known {
				var font spec.FormSpecFont
				if json.Unmarshal(body, &font) == nil {
					control.Font = &font
				}
			}
		case "forecolor", "backcolor":
			n, ok := numberValue(raw)
			if !ok || math.Trunc(n) != n || n < 0 || n > math.MaxUint32 {
				continue
			}
			color := spec.ColorFromOLE(uint32(int64(n)))
			if strings.EqualFold(key, "foreColor") && control.ForeColor == nil {
				control.ForeColor = &color
			}
			if strings.EqualFold(key, "backColor") && control.BackColor == nil {
				control.BackColor = &color
			}
		case "textalign", "borderstyle", "picturealignment", "picturesizemode":
			n, ok := numberValue(raw)
			if !ok || math.Trunc(n) != n || n < 0 || n > 4 {
				continue
			}
			value := int(n)
			switch strings.ToLower(strings.TrimSpace(key)) {
			case "textalign":
				if control.TextAlign == nil {
					control.TextAlign = &value
				}
			case "borderstyle":
				if control.BorderStyle == nil {
					control.BorderStyle = &value
				}
			case "picturealignment":
				if control.PictureAlignment == nil {
					control.PictureAlignment = &value
				}
			case "picturesizemode":
				if control.PictureSizeMode == nil {
					control.PictureSizeMode = &value
				}
			}
		}
	}
}

func persistenceFont(font *spec.FormSpecFont) oforms.Font {
	if font == nil {
		return oforms.Font{}
	}
	return oforms.Font{Name: font.Name, Size: font.Size, Bold: font.Bold, Italic: font.Italic, Underline: font.Underline, StrikeThrough: font.StrikeThrough, Charset: font.Charset}
}

func fontPatch(before, after *spec.FormSpecFont) (oforms.Font, error) {
	old, next := persistenceFont(before), persistenceFont(after)
	r := reflect.ValueOf(&next).Elem()
	b := reflect.ValueOf(old)
	for i := range r.NumField() {
		if reflect.DeepEqual(r.Field(i).Interface(), b.Field(i).Interface()) {
			r.Field(i).SetZero()
		} else if r.Field(i).IsNil() {
			return oforms.Font{}, fmt.Errorf("font member reset is not supported")
		}
	}
	return next, nil
}

func fontMatches(patch, actual oforms.Font) bool {
	p, a := reflect.ValueOf(patch), reflect.ValueOf(actual)
	for i := range p.NumField() {
		if !p.Field(i).IsNil() && !reflect.DeepEqual(p.Field(i).Interface(), a.Field(i).Interface()) {
			return false
		}
	}
	return true
}

type typedField struct {
	name              string
	old, next, actual any
}

func controlTypedFields(old, next, actual spec.FormSpecControl) []typedField {
	return []typedField{
		{"font", old.Font, next.Font, actual.Font},
		{"foreColor", old.ForeColor, next.ForeColor, actual.ForeColor},
		{"backColor", old.BackColor, next.BackColor, actual.BackColor},
		{"textAlign", old.TextAlign, next.TextAlign, actual.TextAlign},
		{"borderStyle", old.BorderStyle, next.BorderStyle, actual.BorderStyle},
		{"pictureAlignment", old.PictureAlignment, next.PictureAlignment, actual.PictureAlignment},
		{"pictureSizeMode", old.PictureSizeMode, next.PictureSizeMode, actual.PictureSizeMode},
	}
}

func typedValue(value any) (any, error) {
	switch v := value.(type) {
	case *spec.FormSpecFont:
		if v == nil {
			return nil, nil
		}
		return persistenceFont(v), nil
	case *spec.FormSpecColor:
		if v == nil {
			return nil, nil
		}
		n, err := spec.ColorOLE(*v)
		return int64(n), err
	case *int:
		if v == nil {
			return nil, nil
		}
		return int64(*v), nil
	default:
		return nil, fmt.Errorf("unknown typed property")
	}
}

func (c *compilation) typedEdits(name, kind, path string, fields []typedField) error {
	for _, f := range fields {
		if reflect.DeepEqual(f.old, f.next) {
			continue
		}
		p := path + "." + f.name
		if kind != "Form" {
			if _, ok := spec.LookupControlProperty(kind, f.name); !ok {
				return c.fail(Unsupported, name, p, "property is not applicable to control")
			}
		}
		value, err := typedValue(f.next)
		if err != nil {
			return c.fail(Invalid, name, p, err.Error())
		}
		if value == nil {
			return c.fail(Unsupported, name, p, "property reset is not supported")
		}
		previous, err := typedValue(f.old)
		if err != nil {
			return c.fail(Invalid, name, p, err.Error())
		}
		actual, err := typedValue(f.actual)
		if err != nil {
			return c.fail(Invalid, name, p, err.Error())
		}
		if previous != nil {
			matches := reflect.DeepEqual(previous, actual)
			if f.name == "font" && actual != nil {
				matches = fontMatches(previous.(oforms.Font), actual.(oforms.Font))
			}
			if !matches {
				return c.fail(Stale, name, p, "before value does not match binary model")
			}
		}
		if f.name == "font" {
			value, err = fontPatch(f.old.(*spec.FormSpecFont), f.next.(*spec.FormSpecFont))
			if err != nil {
				return c.fail(Unsupported, name, p, err.Error())
			}
			if reflect.DeepEqual(value, oforms.Font{}) {
				continue
			}
		}
		persisted := strings.ToUpper(f.name[:1]) + f.name[1:]
		if err = c.add(name, persisted, value, p); err != nil {
			return err
		}
	}
	return nil
}

func hasTypedProperty(control spec.FormSpecControl, key string) bool {
	for _, field := range controlTypedFields(control, control, control) {
		if strings.EqualFold(field.name, key) {
			value, _ := typedValue(field.next)
			return value != nil
		}
	}
	return false
}

func overlayFont(old, desired *spec.FormSpecFont) *spec.FormSpecFont {
	if desired == nil {
		return old
	}
	result := spec.FormSpecFont{}
	if old != nil {
		result = *old
	}
	r, d := reflect.ValueOf(&result).Elem(), reflect.ValueOf(*desired)
	for i := range r.NumField() {
		if !d.Field(i).IsNil() {
			r.Field(i).Set(d.Field(i))
		}
	}
	return &result
}

func overlayTypedControl(target *spec.FormSpecControl, desired spec.FormSpecControl) {
	target.Font = overlayFont(target.Font, desired.Font)
	if desired.BackColor != nil {
		target.BackColor = desired.BackColor
	}
	if desired.ForeColor != nil {
		target.ForeColor = desired.ForeColor
	}
	if desired.BorderStyle != nil {
		target.BorderStyle = desired.BorderStyle
	}
	if desired.TextAlign != nil {
		target.TextAlign = desired.TextAlign
	}
	if desired.PictureAlignment != nil {
		target.PictureAlignment = desired.PictureAlignment
	}
	if desired.PictureSizeMode != nil {
		target.PictureSizeMode = desired.PictureSizeMode
	}
	for key := range target.Properties {
		if hasTypedProperty(desired, key) {
			delete(target.Properties, key)
		}
	}
}

func generationTypedEdits(form spec.FormSpec) ([]oforms.Edit, error) {
	var edits []oforms.Edit
	appendFields := func(name string, fields []typedField) error {
		for _, f := range fields {
			value, err := typedValue(f.next)
			if err != nil {
				return err
			}
			if value == nil {
				continue
			}
			if font, ok := value.(oforms.Font); ok && reflect.DeepEqual(font, oforms.Font{}) {
				continue
			}
			edits = append(edits, oforms.Edit{Control: name, Property: strings.ToUpper(f.name[:1]) + f.name[1:], Value: value})
		}
		return nil
	}
	if err := appendFields("", []typedField{{name: "font", next: form.Form.Font}, {name: "foreColor", next: form.Form.ForeColor}, {name: "backColor", next: form.Form.BackColor}, {name: "borderStyle", next: form.Form.BorderStyle}}); err != nil {
		return nil, err
	}
	for _, control := range form.Controls {
		if err := appendFields(control.Name, controlTypedFields(control, control, control)); err != nil {
			return nil, err
		}
	}
	return edits, nil
}
