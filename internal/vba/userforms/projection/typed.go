package projection

import (
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	forms "github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func projectFont(font *oforms.Font) *forms.FormSpecFont {
	if font == nil {
		return nil
	}
	return &forms.FormSpecFont{Name: font.Name, Size: font.Size, Bold: font.Bold, Italic: font.Italic, Underline: font.Underline, StrikeThrough: font.StrikeThrough, Charset: font.Charset}
}

func projectTypedForm(level *oforms.Level, target *forms.FormSpecForm) {
	font, err := oforms.ReadLevelFont(level)
	if err == nil {
		target.Font = projectFont(font)
	}
	if v, ok := level.Record.Values["BackColor"]; ok {
		target.BackColor = new(forms.ColorFromOLE(uint32(v)))
	}
	if v, ok := level.Record.Values["ForeColor"]; ok {
		target.ForeColor = new(forms.ColorFromOLE(uint32(v)))
	}
	if v, ok := level.Record.Values["BorderStyle"]; ok {
		target.BorderStyle = new(int(v))
	}
}

func projectTypedControl(source *oforms.Control, target *forms.FormSpecControl) {
	font, err := oforms.ReadControlFont(source)
	if _, applicable := forms.LookupControlProperty(target.Type, "font"); err == nil && applicable {
		target.Font = projectFont(font)
	}
	r := source.Record
	if r == nil {
		return
	}
	if source.MultiPage != nil && source.MultiPage.Hidden != nil {
		r = source.MultiPage.Hidden.Record
		if r == nil {
			return
		}
	}
	if _, ok := forms.LookupControlProperty(target.Type, "backColor"); ok {
		if v, found := r.Values["BackColor"]; found {
			target.BackColor = new(forms.ColorFromOLE(uint32(v)))
		}
	}
	if _, ok := forms.LookupControlProperty(target.Type, "foreColor"); ok {
		if v, found := r.Values["ForeColor"]; found {
			target.ForeColor = new(forms.ColorFromOLE(uint32(v)))
		}
	}
	if _, ok := forms.LookupControlProperty(target.Type, "borderStyle"); ok {
		if v, found := r.Values["BorderStyle"]; found {
			target.BorderStyle = new(int(v))
		}
	}
	if _, ok := forms.LookupControlProperty(target.Type, "textAlign"); ok && r.TextProps != nil {
		if v, found := r.TextProps.Values["ParagraphAlign"]; found {
			switch v {
			case 2:
				v = 3
			case 3:
				v = 2
			}
			target.TextAlign = new(int(v))
		}
	}
	if target.Type == "Image" {
		if v, ok := r.Values["PictureAlignment"]; ok {
			target.PictureAlignment = new(int(v))
		}
		if v, ok := r.Values["PictureSizeMode"]; ok {
			target.PictureSizeMode = new(int(v))
		}
	}
}
