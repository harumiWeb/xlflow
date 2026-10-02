package compiler

import (
	"cmp"
	"errors"
	"fmt"
	"maps"
	"math"
	"reflect"
	"slices"
	"strings"
	"unicode"

	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

const (
	GenerationInvalid     = "userform_generation_invalid"
	GenerationUnsupported = "userform_generation_unsupported"
	GenerationConflict    = "userform_generation_conflict"
)

var generationDefaults = map[string]struct {
	class         uint16
	width, height float64
}{
	"label": {21, 72, 18}, "textbox": {23, 120, 18}, "commandbutton": {17, 72, 24},
	"checkbox": {26, 72, 18}, "optionbutton": {27, 72, 18}, "togglebutton": {28, 72, 18},
	"combobox": {25, 120, 18}, "listbox": {24, 120, 72}, "spinbutton": {16, 18, 36},
	"scrollbar": {47, 120, 18}, "image": {12, 72, 72},
}

// CompileNew builds a supported flat Designer from canonical authoring intent.
// It owns no code-behind, project references, Excel instance, or source files.
func CompileNew(input spec.FormSpec, codePage uint16) (*oforms.Form, error) {
	fail := func(code, control, path, reason string) error {
		return &Error{Code: code, Form: input.Form.Name, Control: control, Property: path, Reason: reason}
	}
	if input.CoordinateSystem != "" && input.CoordinateSystem != "points" && input.CoordinateSystem != "parent-relative" {
		return nil, fail(GenerationInvalid, "", "coordinateSystem", "requires points or parent-relative")
	}
	if err := vbaproject.ValidateWritableComponentIdentity(input.Form.Name, input.Form.Name, codePage); err != nil {
		return nil, fail(GenerationInvalid, "", "form.name", err.Error())
	}
	form, err := normalizedCopy(input)
	if err != nil {
		return nil, fail(GenerationInvalid, "", "spec", err.Error())
	}
	if form.Form.Width != nil || form.Form.Height != nil || (form.Form.Build != nil && (form.Form.Build.Width != nil || form.Form.Build.Height != nil)) {
		return nil, fail(GenerationUnsupported, "", "form", "outer dimensions require conversion; use form.build.clientWidth/clientHeight")
	}
	width, height := 240.0, 180.0
	caption := form.Form.Name
	if form.Form.Caption != nil {
		caption = *form.Form.Caption
	}
	if form.Form.Build != nil {
		if form.Form.Build.Caption != nil {
			caption = *form.Form.Build.Caption
		}
		if form.Form.Build.ClientWidth != nil {
			width = *form.Form.Build.ClientWidth
		}
		if form.Form.Build.ClientHeight != nil {
			height = *form.Form.Build.ClientHeight
		}
	}
	rootSize, err := generationSize(width, height)
	if err != nil {
		return nil, fail(GenerationInvalid, "", "form.build", err.Error())
	}
	definition := oforms.Definition{Name: form.Form.Name, Caption: caption, Size: rootSize}
	if !validGenerationName(form.Form.Name) {
		return nil, fail(GenerationInvalid, "", "form.name", "requires a VBA identifier")
	}
	originalIndices := make(map[string]int, len(form.Controls))
	for i, control := range form.Controls {
		originalIndices[control.ID] = i
	}
	slices.SortStableFunc(form.Controls, func(a, b spec.FormSpecControl) int { return cmp.Compare(*a.ZIndex, *b.ZIndex) })
	for i, control := range form.Controls {
		path := fmt.Sprintf("controls[%d]", originalIndices[control.ID])
		defaults, known := generationDefaults[strings.ToLower(control.Type)]
		if !known || control.ParentID != "" || len(control.Controls) != 0 {
			return nil, fail(GenerationUnsupported, control.Name, path, "only flat built-in common controls can be generated")
		}
		contract, _ := spec.LookupControlContract(control.Type)
		control.Type = contract.Type
		if control.ProgID != "" && !strings.EqualFold(control.ProgID, contract.ProgID) {
			return nil, fail(GenerationUnsupported, control.Name, path+".progId", "custom ProgID cannot be generated")
		}
		if !validGenerationName(control.Name) {
			return nil, fail(GenerationInvalid, control.Name, path+".name", "requires a VBA identifier")
		}
		if len(control.List) != 0 || control.SelectedIndex != nil {
			return nil, fail(GenerationUnsupported, control.Name, path, "list/selectedIndex persistence is not supported")
		}
		props, err := propertyMap(control.Properties)
		if err != nil {
			return nil, fail(GenerationConflict, control.Name, path+".properties", err.Error())
		}
		// Collapse every explicit authoring alias before applying defaults.
		for _, field := range []struct {
			key   string
			value any
		}{
			{"caption", pointerValue(control.Caption)}, {"text", pointerValue(control.Text)}, {"value", control.Value},
			{"left", pointerValue(control.Left)}, {"top", pointerValue(control.Top)}, {"width", pointerValue(control.Width)}, {"height", pointerValue(control.Height)},
			{"tabindex", pointerValue(control.TabIndex)}, {"enabled", pointerValue(control.Enabled)}, {"visible", pointerValue(control.Visible)},
		} {
			if field.value == nil {
				continue
			}
			if old, exists := props[field.key]; exists {
				a, e := generationValue(control.Type, field.key, old)
				if e != nil {
					return nil, fail(GenerationInvalid, control.Name, path+".properties."+field.key, e.Error())
				}
				b, e := generationValue(control.Type, field.key, field.value)
				if e != nil {
					return nil, fail(GenerationInvalid, control.Name, path+"."+field.key, e.Error())
				}
				if !reflect.DeepEqual(a, b) {
					return nil, fail(GenerationConflict, control.Name, path+"."+field.key, "conflicting property aliases")
				}
			}
			props[field.key] = field.value
		}
		values := map[string]any{}
		for _, key := range slices.Sorted(maps.Keys(props)) {
			name, known := propertyNames[key]
			if !known || !applicableProperty(control.Type, key) {
				return nil, fail(GenerationUnsupported, control.Name, path+".properties."+key, "property is not supported for this control")
			}
			value, e := generationValue(control.Type, key, props[key])
			if e != nil {
				return nil, fail(GenerationInvalid, control.Name, path+"."+key, e.Error())
			}
			if control.Type == "SpinButton" || control.Type == "ScrollBar" {
				if key == "value" {
					name = "Position"
				}
			}
			if old, exists := values[name]; exists && !reflect.DeepEqual(old, value) {
				return nil, fail(GenerationConflict, control.Name, path, "conflicting text/value aliases")
			}
			values[name] = value
		}
		item := oforms.ControlDefinition{Name: control.Name, Class: defaults.class, Visible: true, TabIndex: int16(i), Properties: map[string]any{}}
		if i > math.MaxInt16 {
			return nil, fail(GenerationInvalid, control.Name, path, "too many controls for default tab order")
		}
		item.Size, _ = generationSize(defaults.width, defaults.height)
		for name, value := range values {
			switch name {
			case "Left":
				item.Position.Left = value.(int32)
			case "Top":
				item.Position.Top = value.(int32)
			case "Width":
				item.Size.Width = value.(int32)
			case "Height":
				item.Size.Height = value.(int32)
			case "TabIndex":
				item.TabIndex = value.(int16)
			case "Visible":
				item.Visible = value.(bool)
			case "Enabled": // handled with the omission default below
			default:
				item.Properties[name] = value
			}
		}
		if _, ok := spec.LookupControlProperty(control.Type, "caption"); ok {
			if _, found := item.Properties["Caption"]; !found {
				item.Properties["Caption"] = control.Name
			}
		}
		if _, ok := spec.LookupControlProperty(control.Type, "value"); ok && control.Type != "ListBox" && control.Type != "SpinButton" && control.Type != "ScrollBar" {
			if _, found := item.Properties["Value"]; !found {
				v := ""
				if defaults.class >= 26 && defaults.class <= 28 {
					v = "0"
				}
				item.Properties["Value"] = v
			}
		}
		if defaults.class == 16 || defaults.class == 47 {
			if _, found := item.Properties["Position"]; !found {
				item.Properties["Position"] = int64(0)
			}
		}
		bits, ok := oforms.DefaultVariousPropertyBits(defaults.class)
		if !ok {
			bits = 0x1b
		}
		if defaults.class == 23 || defaults.class == 25 {
			bits = 0x2c80481b
		}
		if enabled, found := values["Enabled"]; found && !enabled.(bool) {
			bits &^= 2
		}
		item.Properties["VariousPropertyBits"] = bits
		definition.Controls = append(definition.Controls, item)
	}
	result, err := oforms.NewForm(definition, codePage)
	if err != nil {
		code := GenerationInvalid
		if errors.Is(err, oforms.ErrUnsupportedEdit) {
			code = GenerationUnsupported
		}
		return nil, &Error{Code: code, Form: form.Form.Name, Reason: err.Error(), Cause: err}
	}
	return result, nil
}

func validGenerationName(name string) bool {
	if name == "" {
		return false
	}
	for i, r := range name {
		valid := unicode.IsLetter(r) || i > 0 && (unicode.IsDigit(r) || r == '_')
		if !valid {
			return false
		}
	}
	return true
}

func generationSize(width, height float64) (oforms.Size, error) {
	w, e := convertValue("", "width", width)
	if e != nil {
		return oforms.Size{}, e
	}
	h, e := convertValue("", "height", height)
	if e != nil {
		return oforms.Size{}, e
	}
	return oforms.Size{Width: w.(int32), Height: h.(int32)}, nil
}

func generationValue(kind, key string, value any) (any, error) {
	if value == nil {
		return nil, fmt.Errorf("explicit null/reset is unsupported")
	}
	if key == "tabindex" {
		switch v := value.(type) {
		case int:
			return convertValue(kind, "tabIndex", v)
		case float64:
			return bagValue(kind, key, v)
		}
	}
	if key == "value" && (strings.EqualFold(kind, "SpinButton") || strings.EqualFold(kind, "ScrollBar")) {
		v, ok := value.(float64)
		if !ok {
			switch n := value.(type) {
			case int:
				v = float64(n)
				ok = true
			case int64:
				v = float64(n)
				ok = true
			}
		}
		if !ok || math.IsNaN(v) || math.IsInf(v, 0) || math.Trunc(v) != v || v < 0 || v > 100 {
			return nil, fmt.Errorf("value must be an integer in the supported default range 0..100")
		}
		return int64(v), nil
	}
	if (key == "value" || key == "text") && strings.EqualFold(kind, "ToggleButton") {
		return convertValue("CheckBox", key, value)
	}
	return bagValue(kind, key, value)
}
