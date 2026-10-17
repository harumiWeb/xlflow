package spec

import (
	"encoding/json"
	"fmt"
	"math"
	"reflect"
	"strconv"
	"strings"
)

// FormSpecFont is a partial edit. An omitted member preserves the persisted font.
type FormSpecFont struct {
	Name          *string  `json:"name,omitempty" yaml:"name,omitempty"`
	Size          *float64 `json:"size,omitempty" yaml:"size,omitempty"`
	Bold          *bool    `json:"bold,omitempty" yaml:"bold,omitempty"`
	Italic        *bool    `json:"italic,omitempty" yaml:"italic,omitempty"`
	Underline     *bool    `json:"underline,omitempty" yaml:"underline,omitempty"`
	StrikeThrough *bool    `json:"strikeThrough,omitempty" yaml:"strikeThrough,omitempty"`
	Charset       *int     `json:"charset,omitempty" yaml:"charset,omitempty"`
}

// FormSpecColor preserves the distinction between RGB and an OLE system color.
type FormSpecColor struct {
	Kind  string `json:"kind" yaml:"kind"`
	Value string `json:"value" yaml:"value"`
}

var systemColors = map[string]uint32{
	"scrollBar": 0, "desktop": 1, "activeCaption": 2, "inactiveCaption": 3,
	"menu": 4, "window": 5, "windowFrame": 6, "menuText": 7, "windowText": 8,
	"captionText": 9, "activeBorder": 10, "inactiveBorder": 11, "appWorkspace": 12,
	"highlight": 13, "highlightText": 14, "buttonFace": 15, "buttonShadow": 16,
	"grayText": 17, "buttonText": 18, "inactiveCaptionText": 19, "buttonHighlight": 20,
	"buttonDarkShadow": 21, "buttonLight": 22, "infoText": 23, "infoBackground": 24,
	"hotLight": 26, "gradientActiveCaption": 27, "gradientInactiveCaption": 28,
	"menuHighlight": 29, "menuBar": 30,
}

func ColorOLE(color FormSpecColor) (uint32, error) {
	switch color.Kind {
	case "rgb":
		if len(color.Value) == 7 && color.Value[0] == '#' {
			n, err := strconv.ParseUint(color.Value[1:], 16, 24)
			if err == nil {
				return uint32(n&255)<<16 | uint32(n&0xff00) | uint32(n>>16), nil
			}
		}
	case "system":
		if n, ok := systemColors[color.Value]; ok {
			return 0x80000000 | n, nil
		}
	case "ole":
		if len(color.Value) == 10 && strings.HasPrefix(color.Value, "0x") {
			n, err := strconv.ParseUint(color.Value[2:], 16, 32)
			if err == nil {
				return uint32(n), nil
			}
		}
	}
	return 0, fmt.Errorf("invalid %s color %q", color.Kind, color.Value)
}

func ColorFromOLE(value uint32) FormSpecColor {
	if value&0xffffff00 == 0x80000000 {
		for name, index := range systemColors {
			if value&255 == index {
				return FormSpecColor{Kind: "system", Value: name}
			}
		}
	}
	if value&0xff000000 == 0 {
		return FormSpecColor{Kind: "rgb", Value: fmt.Sprintf("#%02X%02X%02X", value&255, value>>8&255, value>>16&255)}
	}
	return FormSpecColor{Kind: "ole", Value: fmt.Sprintf("0x%08X", value)}
}

func withPropertyEditors(contract Contract) Contract {
	font := property(ValueTypeObject, false, SupportLevelSupported, "Partial MSForms font; size is in points and charset is numeric.", true)
	font.Editor = "font"
	font.Children = map[string]PropertyContract{}
	for _, name := range []string{"bold", "italic", "underline", "strikeThrough"} {
		font.Children[name] = property(ValueTypeBoolean, false, SupportLevelSupported, "Font style.", true)
	}
	font.Children["name"] = property(ValueTypeString, false, SupportLevelSupported, "Font family name.", true)
	font.Children["size"] = property(ValueTypeNumber, false, SupportLevelSupported, "Font size in points, greater than zero.", true)
	font.Children["charset"] = propertyWithAllowed(ValueTypeInteger, false, SupportLevelSupported, "Windows charset number; unlisted values remain valid.", true)
	sizeContract := font.Children["size"]
	sizeContract.Minimum = new(0.0)
	sizeContract.Maximum = new(3276.75)
	sizeContract.ExclusiveMinimum = true
	font.Children["size"] = sizeContract
	charsetContract := font.Children["charset"]
	charsetContract.Minimum = new(0.0)
	charsetContract.Maximum = new(255.0)
	font.Children["charset"] = charsetContract
	color := property(ValueTypeObject, false, SupportLevelSupported, "RGB, system or raw OLE color.", true)
	color.Editor = "color"
	color.Children = map[string]PropertyContract{
		"kind":  propertyWithAllowed(ValueTypeString, true, SupportLevelSupported, "Color representation.", true, "rgb", "system", "ole"),
		"value": property(ValueTypeString, true, SupportLevelSupported, "#RRGGBB, system color name, or 0xXXXXXXXX.", true),
	}
	contract.FormProperties["font"] = font
	contract.FormProperties["foreColor"] = color
	contract.FormProperties["backColor"] = color
	border := propertyWithAllowed(ValueTypeInteger, false, SupportLevelSupported, "MSForms border style.", true, "0", "1")
	border.Editor = "enum"
	contract.FormProperties["borderStyle"] = border
	textAlign := propertyWithAllowed(ValueTypeInteger, false, SupportLevelSupported, "Text alignment: 1 Left, 2 Center, 3 Right.", true, "1", "2", "3")
	textAlign.Editor = "enum"
	for key, control := range contract.Controls {
		if control.Type != "Page" && !strings.EqualFold(control.Type, "Image") && !strings.EqualFold(control.Type, "SpinButton") && !strings.EqualFold(control.Type, "ScrollBar") {
			control.Properties["font"] = font
		}
		if control.Type != "Page" {
			control.Properties["backColor"] = color
		}
		if control.Type != "Image" && control.Type != "Page" {
			control.Properties["foreColor"] = color
		}
		switch control.Type {
		case "Label", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton":
			control.Properties["textAlign"] = textAlign
		}
		switch control.Type {
		case "Frame", "TextBox", "ComboBox", "ListBox", "Label", "Image":
			control.Properties["borderStyle"] = border
		}
		if control.Type == "Image" {
			picture := control.Properties["picture"]
			picture.Editor = "picture"
			picture.Children = pictureActionProperties()
			control.Properties["picture"] = picture
			alignment := propertyWithAllowed(ValueTypeInteger, false, SupportLevelSupported, "Picture alignment: TopLeft, TopRight, Center, BottomLeft, BottomRight.", true, "0", "1", "2", "3", "4")
			alignment.Editor = "enum"
			control.Properties["pictureAlignment"] = alignment
			mode := propertyWithAllowed(ValueTypeInteger, false, SupportLevelSupported, "Picture size mode: Clip, Stretch, Zoom.", true, "0", "1", "3")
			mode.Editor = "enum"
			control.Properties["pictureSizeMode"] = mode
		}
		contract.Controls[key] = control
	}
	return contract
}

func typedAppearanceIssue(field, message string) ValidationIssue {
	return validationIssue("UFV021", SeverityError, message, field, "Use the typed property contract; source files remain UTF-8.", "")
}

func validateFont(font *FormSpecFont, path string) []ValidationIssue {
	var issues []ValidationIssue
	if font == nil {
		return issues
	}
	if font.Size != nil && (math.IsNaN(*font.Size) || math.IsInf(*font.Size, 0) || *font.Size <= 0 || *font.Size > 3276.75) {
		issues = append(issues, typedAppearanceIssue(path+".size", "Font size must be greater than zero and at most 3276.75 points."))
	}
	if font.Charset != nil && (*font.Charset < 0 || *font.Charset > 255) {
		issues = append(issues, typedAppearanceIssue(path+".charset", "Font charset must be an integer from 0 to 255."))
	}
	if font.Name != nil && (len(*font.Name) > 255 || strings.ContainsRune(*font.Name, 0)) {
		issues = append(issues, typedAppearanceIssue(path+".name", "Font name must contain at most 255 bytes and no NUL."))
	}
	return issues
}

func validateTypedForm(form FormSpecForm) []ValidationIssue {
	issues := validateFont(form.Font, "form.font")
	for name, color := range map[string]*FormSpecColor{"foreColor": form.ForeColor, "backColor": form.BackColor} {
		if color != nil {
			if _, err := ColorOLE(*color); err != nil {
				issues = append(issues, typedAppearanceIssue("form."+name, err.Error()))
			}
		}
	}
	if form.BorderStyle != nil && (*form.BorderStyle < 0 || *form.BorderStyle > 1) {
		issues = append(issues, typedAppearanceIssue("form.borderStyle", "BorderStyle must be 0 or 1."))
	}
	return issues
}

func validateTypedControl(control FormSpecControl, path string) []ValidationIssue {
	issues := validateFont(control.Font, path+".font")
	fields := map[string]any{"font": control.Font, "foreColor": control.ForeColor, "backColor": control.BackColor, "textAlign": control.TextAlign, "borderStyle": control.BorderStyle, "pictureAlignment": control.PictureAlignment, "pictureSizeMode": control.PictureSizeMode}
	for _, name := range sortedMapKeys(fields) {
		value := fields[name]
		if reflect.ValueOf(value).IsNil() {
			continue
		}
		contract, ok := LookupControlProperty(control.Type, name)
		if !ok || (control.ProgID != "" && !ProgIDMatchesControlType(control.Type, control.ProgID)) {
			issues = append(issues, invalidControlPropertyIssue(path+"."+name, control.Type))
			continue
		}
		body, _ := json.Marshal(value)
		var raw any
		_ = json.Unmarshal(body, &raw)
		if len(contract.AllowedValues) > 0 && !valueInAllowedValues(raw, contract.AllowedValues) {
			issues = append(issues, typedAppearanceIssue(path+"."+name, "Value is outside the supported property enum."))
		}
		if color, ok := value.(*FormSpecColor); ok {
			if _, err := ColorOLE(*color); err != nil {
				issues = append(issues, typedAppearanceIssue(path+"."+name, err.Error()))
			}
		}
		if legacy, present := lookupRawField(control.Properties, name); present && !legacyAppearanceEqual(name, raw, legacy) {
			issues = append(issues, typedAppearanceIssue(path+"."+name, "Typed and legacy property values conflict."))
		}
	}
	return issues
}

func legacyAppearanceEqual(name string, typed, legacy any) bool {
	if name == "foreColor" || name == "backColor" {
		body, _ := json.Marshal(typed)
		var color FormSpecColor
		if json.Unmarshal(body, &color) == nil {
			ole, err := ColorOLE(color)
			n, ok := rawFloat64(legacy)
			signed := int64(ole)
			if ole > math.MaxInt32 {
				signed -= 1 << 32
			}
			return err == nil && ok && (n == float64(ole) || n == float64(signed))
		}
	}
	a, _ := json.Marshal(typed)
	b, _ := json.Marshal(legacy)
	return string(a) == string(b)
}
