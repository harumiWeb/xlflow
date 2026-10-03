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
	"scrollbar": {47, 120, 18}, "image": {12, 72, 72}, "frame": {14, 144, 108},
}

// CompileNew builds a supported Designer tree from canonical authoring intent.
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
	children := make(map[string][]spec.FormSpecControl, len(form.Controls))
	for i, control := range form.Controls {
		originalIndices[control.ID] = i
		children[control.ParentID] = append(children[control.ParentID], control)
	}
	for parent := range children {
		slices.SortStableFunc(children[parent], func(a, b spec.FormSpecControl) int {
			return cmp.Compare(*a.ZIndex, *b.ZIndex)
		})
	}
	var buildControls func(string, int) ([]oforms.ControlDefinition, error)
	buildControls = func(parentID string, depth int) ([]oforms.ControlDefinition, error) {
		result := make([]oforms.ControlDefinition, 0, len(children[parentID]))
		for siblingIndex, control := range children[parentID] {
			path := fmt.Sprintf("controls[%d]", originalIndices[control.ID])
			item, err := generationControlDefinition(control, siblingIndex)
			if err != nil {
				return nil, generationErrorForControl(form.Form.Name, control.Name, path, err)
			}
			if nested := children[control.ID]; len(nested) != 0 {
				if !strings.EqualFold(control.Type, "Frame") {
					return nil, fail(GenerationUnsupported, control.Name, path+".controls", "only Frame controls can contain generated controls")
				}
				if depth+1 >= oforms.MaxNestingDepth {
					return nil, fail(GenerationInvalid, control.Name, path+".controls", fmt.Sprintf("Designer nesting exceeds %d levels", oforms.MaxNestingDepth))
				}
				item.Controls, err = buildControls(control.ID, depth+1)
				if err != nil {
					return nil, err
				}
			}
			if strings.EqualFold(control.Type, "Frame") && depth+1 >= oforms.MaxNestingDepth {
				return nil, fail(GenerationInvalid, control.Name, path, fmt.Sprintf("Designer nesting exceeds %d levels", oforms.MaxNestingDepth))
			}
			result = append(result, item)
		}
		return result, nil
	}
	definition.Controls, err = buildControls("", 0)
	if err != nil {
		return nil, err
	}
	result, err := oforms.NewForm(definition, codePage)
	if err != nil {
		code := GenerationInvalid
		if errors.Is(err, oforms.ErrUnsupportedEdit) {
			code = GenerationUnsupported
		}
		if detail, ok := errors.AsType[*oforms.EditError](err); ok && detail.Control == "" && detail.Property == "Caption" {
			property := "form.caption"
			if form.Form.Build != nil && form.Form.Build.Caption != nil {
				property = "form.build.caption"
			}
			return nil, &Error{Code: code, Form: form.Form.Name, Property: property, Reason: detail.Error(), Cause: err}
		}
		return nil, &Error{Code: code, Form: form.Form.Name, Reason: err.Error(), Cause: err}
	}
	return result, nil
}

type generationControlError struct {
	code, property, reason string
}

func (e *generationControlError) Error() string { return e.reason }

func generationErrorForControl(form, control, path string, err error) error {
	if detail, ok := errors.AsType[*generationControlError](err); ok {
		property := path
		if detail.property != "" {
			property += "." + detail.property
		}
		return &Error{Code: detail.code, Form: form, Control: control, Property: property, Reason: detail.reason, Cause: err}
	}
	return &Error{Code: GenerationInvalid, Form: form, Control: control, Property: path, Reason: err.Error(), Cause: err}
}

func generationFailure(code, property, reason string) error {
	return &generationControlError{code: code, property: property, reason: reason}
}

func generationControlDefinition(control spec.FormSpecControl, tabIndex int) (oforms.ControlDefinition, error) {
	defaults, known := generationDefaults[strings.ToLower(control.Type)]
	if !known {
		return oforms.ControlDefinition{}, generationFailure(GenerationUnsupported, "type", "control type cannot be generated")
	}
	contract, _ := spec.LookupControlContract(control.Type)
	kind := contract.Type
	if control.ProgID != "" && !strings.EqualFold(control.ProgID, contract.ProgID) {
		return oforms.ControlDefinition{}, generationFailure(GenerationUnsupported, "progId", "custom ProgID cannot be generated")
	}
	if !validGenerationName(control.Name) {
		return oforms.ControlDefinition{}, generationFailure(GenerationInvalid, "name", "requires a VBA identifier")
	}
	if len(control.List) != 0 || control.SelectedIndex != nil {
		return oforms.ControlDefinition{}, generationFailure(GenerationUnsupported, "", "list/selectedIndex persistence is not supported")
	}
	if tabIndex < 0 || tabIndex > math.MaxInt16 {
		return oforms.ControlDefinition{}, generationFailure(GenerationInvalid, "", "too many controls for default tab order")
	}
	props, err := propertyMap(control.Properties)
	if err != nil {
		return oforms.ControlDefinition{}, generationFailure(GenerationConflict, "properties", err.Error())
	}
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
			a, convertErr := generationValue(kind, field.key, old)
			if convertErr != nil {
				return oforms.ControlDefinition{}, generationFailure(GenerationInvalid, "properties."+field.key, convertErr.Error())
			}
			b, convertErr := generationValue(kind, field.key, field.value)
			if convertErr != nil {
				return oforms.ControlDefinition{}, generationFailure(GenerationInvalid, field.key, convertErr.Error())
			}
			if !reflect.DeepEqual(a, b) {
				return oforms.ControlDefinition{}, generationFailure(GenerationConflict, field.key, "conflicting property aliases")
			}
		}
		props[field.key] = field.value
	}
	values := make(map[string]any, len(props))
	for _, key := range slices.Sorted(maps.Keys(props)) {
		name, known := propertyNames[key]
		if !known || !applicableProperty(kind, key) {
			return oforms.ControlDefinition{}, generationFailure(GenerationUnsupported, "properties."+key, "property is not supported for this control")
		}
		value, convertErr := generationValue(kind, key, props[key])
		if convertErr != nil {
			return oforms.ControlDefinition{}, generationFailure(GenerationInvalid, key, convertErr.Error())
		}
		if (kind == "SpinButton" || kind == "ScrollBar") && key == "value" {
			name = "Position"
		}
		if old, exists := values[name]; exists && !reflect.DeepEqual(old, value) {
			return oforms.ControlDefinition{}, generationFailure(GenerationConflict, "", "conflicting text/value aliases")
		}
		values[name] = value
	}
	item := oforms.ControlDefinition{Name: control.Name, Class: defaults.class, Visible: true, TabIndex: int16(tabIndex), Properties: map[string]any{}}
	item.Size, err = generationSize(defaults.width, defaults.height)
	if err != nil {
		return oforms.ControlDefinition{}, generationFailure(GenerationInvalid, "", err.Error())
	}
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
		case "Enabled": // handled using the control class's persisted flag below
		default:
			item.Properties[name] = value
		}
	}
	if _, ok := spec.LookupControlProperty(kind, "caption"); ok {
		if _, found := item.Properties["Caption"]; !found {
			item.Properties["Caption"] = control.Name
		}
	}
	if _, ok := spec.LookupControlProperty(kind, "value"); ok && kind != "ListBox" && kind != "SpinButton" && kind != "ScrollBar" {
		if _, found := item.Properties["Value"]; !found {
			value := ""
			if defaults.class >= 26 && defaults.class <= 28 {
				value = "0"
			}
			item.Properties["Value"] = value
		}
	}
	if defaults.class == 16 || defaults.class == 47 {
		if _, found := item.Properties["Position"]; !found {
			item.Properties["Position"] = int64(0)
		}
	}
	if defaults.class == 14 {
		bits := int64(0x8004)
		if enabled, found := values["Enabled"]; found && !enabled.(bool) {
			bits &^= 4
		}
		item.Properties["BooleanProperties"] = bits
	} else {
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
	}
	return item, nil
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
