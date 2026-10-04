package compiler

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"maps"
	"reflect"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

// CompileTemplate applies the explicitly authored part of desired to base.
//
// A template update has a complete control list, but the control IDs belong to
// the source document and are not persistence identities. Controls are matched
// by their exact names; the template's IDs are used only to resolve parentId.
// The projected base is the authoritative before snapshot, so omitted fields,
// observed values, warnings, and unsupported snapshot metadata do not become
// edits. The returned model is independent of both inputs.
func CompileTemplate(base *oforms.Form, desired spec.FormSpec, codePage uint16) (*oforms.Form, error) {
	if base == nil {
		return nil, &Error{Code: Invalid, Property: "base", Reason: "base form is nil"}
	}
	baseline, err := projection.Project(base)
	if err != nil {
		return nil, err
	}
	if err := validateTemplateCoordinateSystem(base.Name, desired.CoordinateSystem); err != nil {
		return nil, err
	}
	if desired.Form.Name != base.Name {
		return nil, templateError(base, Unsupported, "", "form.name", "form rename or mismatched form identity")
	}

	after, err := cloneFormSpec(baseline)
	if err != nil {
		return nil, templateError(base, Invalid, "", "template", err.Error())
	}
	if desired.CoordinateSystem != "" {
		after.CoordinateSystem = desired.CoordinateSystem
	}

	desired, err = normalizedCopy(desired)
	if err != nil {
		return nil, templateError(base, Invalid, "", "spec", err.Error())
	}
	if err := overlayTemplateForm(base, &after.Form, desired.Form); err != nil {
		return nil, err
	}
	after.Controls, err = overlayTemplateTopology(base, baseline.Controls, desired.Controls)
	if err != nil {
		return nil, err
	}
	after.Warnings = reconcileTemplatePictureWarnings(after.Warnings, desired.Controls)

	result, err := CompileEdits(base, baseline, after, codePage)
	if err != nil {
		return nil, err
	}
	result, err = applyExplicitTabSelection(result, desired, codePage)
	if err != nil {
		return nil, err
	}
	caption, explicit := templateExplicitRootCaption(desired.Form)
	if !explicit {
		return result, nil
	}
	// CompileEdits updates the Unicode f-stream. Re-apply an explicit root
	// caption through oforms so legacy generated templates also receive the
	// runtime-authoritative VBFrame Caption line. ApplyEdits skips the stream
	// write when the existing line already has the same semantic value.
	result, err = oforms.ApplyEdits(result, []oforms.Edit{{Control: "", Property: "Caption", Value: *caption}}, codePage)
	if err != nil {
		property := "form.caption"
		if desired.Form.Build != nil && desired.Form.Build.Caption != nil {
			property = "form.build.caption"
		}
		return nil, templateError(base, Invalid, "", property, err.Error())
	}
	return result, nil
}

func templateExplicitRootCaption(form spec.FormSpecForm) (*string, bool) {
	if form.Build != nil && form.Build.Caption != nil {
		return form.Build.Caption, true
	}
	if form.Caption != nil {
		return form.Caption, true
	}
	return nil, false
}

func validateTemplateCoordinateSystem(form, coordinateSystem string) error {
	switch coordinateSystem {
	case "", "points", "parent-relative":
		return nil
	default:
		return &Error{
			Code:     Invalid,
			Form:     form,
			Property: "coordinateSystem",
			Reason:   "geometry requires points or parent-relative coordinates",
		}
	}
}

func overlayTemplateForm(base *oforms.Form, target *spec.FormSpecForm, desired spec.FormSpecForm) error {
	if desired.Caption != nil && desired.Build != nil && desired.Build.Caption != nil && *desired.Caption != *desired.Build.Caption {
		return templateError(base, Conflict, "", "form.caption", "aliases request different values for the same persisted property")
	}

	caption := desired.Caption
	if desired.Build != nil && desired.Build.Caption != nil {
		caption = desired.Build.Caption
	}
	if caption != nil {
		target.Caption = new(*caption)
		if target.Build == nil {
			target.Build = &spec.FormSpecBuildForm{}
		}
		target.Build.Caption = new(*caption)
	}

	if desired.Width != nil {
		target.Width = new(*desired.Width)
	}
	if desired.Height != nil {
		target.Height = new(*desired.Height)
	}
	if desired.Build == nil {
		return nil
	}
	if target.Build == nil {
		target.Build = &spec.FormSpecBuildForm{}
	}
	for _, field := range []struct {
		target **float64
		source *float64
	}{
		{&target.Build.Width, desired.Build.Width},
		{&target.Build.Height, desired.Build.Height},
		{&target.Build.ClientWidth, desired.Build.ClientWidth},
		{&target.Build.ClientHeight, desired.Build.ClientHeight},
	} {
		if field.source != nil {
			*field.target = new(*field.source)
		}
	}
	return nil
}

type templateAliasRequest struct {
	canonical  string
	key        string
	raw        any
	normalized any
}

func overlayTemplateControl(base *oforms.Form, target *spec.FormSpecControl, desired spec.FormSpecControl, index int) error {
	requests, err := templateAliasRequests(desired)
	if err != nil {
		if conversionErr, ok := errors.AsType[*templateConversionError](err); ok {
			return templateError(base, Invalid, desired.Name, fmt.Sprintf("controls[%d].%s", index, conversionErr.key), conversionErr.err.Error())
		}
		return templateError(base, Invalid, desired.Name, fmt.Sprintf("controls[%d]", index), err.Error())
	}
	groups := groupTemplateAliases(requests)
	for _, canonical := range slices.Sorted(maps.Keys(groups)) {
		group := groups[canonical]
		if len(group) == 0 {
			continue
		}
		first := group[0]
		for _, request := range group[1:] {
			if !reflect.DeepEqual(first.normalized, request.normalized) {
				return templateError(base, Conflict, desired.Name, fmt.Sprintf("controls[%d].%s", index, strings.ToLower(canonical)), "aliases request different values for the same persisted property")
			}
		}
		applyTemplateAlias(target, first)
	}

	properties := maps.Clone(target.Properties)
	if properties == nil {
		properties = make(map[string]any)
	}
	bagKeys := make(map[string][]string)
	for key := range desired.Properties {
		canonical, known := templateCanonicalProperty(key)
		if known {
			bagKeys[canonical] = append(bagKeys[canonical], key)
		}
	}
	for _, canonical := range slices.Sorted(maps.Keys(bagKeys)) {
		keys := bagKeys[canonical]
		slices.Sort(keys)
		for key := range properties {
			if previous, known := templateCanonicalProperty(key); known && previous == canonical {
				delete(properties, key)
			}
		}
		key := keys[0]
		properties[key] = desired.Properties[key]
	}
	for _, canonical := range slices.Sorted(maps.Keys(groups)) {
		group := groups[canonical]
		if _, explicitBag := bagKeys[canonical]; explicitBag {
			continue
		}
		for key := range properties {
			previous, known := templateCanonicalProperty(key)
			if !known || previous != canonical {
				continue
			}
			properties[key] = group[0].raw
		}
	}
	for key, value := range desired.Properties {
		if _, known := templateCanonicalProperty(key); !known {
			properties[key] = value
		}
	}
	if len(properties) == 0 {
		target.Properties = nil
	} else {
		target.Properties = properties
	}

	if desired.Text != nil && target.Text == nil {
		target.Text = new(*desired.Text)
	}
	if desired.List != nil {
		target.List = slices.Clone(desired.List)
	}
	if desired.SelectedIndex != nil {
		target.SelectedIndex = new(*desired.SelectedIndex)
	}
	if desired.Tabs != nil {
		target.Tabs = overlayTabs(target.Tabs, desired.Tabs)
	}
	if desired.Picture != nil {
		pictureAction := *desired.Picture
		pictureAction.Data = bytes.Clone(desired.Picture.Data)
		target.Picture = &pictureAction
		target.Unsupported = slices.DeleteFunc(target.Unsupported, func(value string) bool {
			return strings.EqualFold(strings.TrimSpace(value), "picture")
		})
		if len(target.Unsupported) == 0 {
			target.Unsupported = nil
		}
	}
	return nil
}

type templateConversionError struct {
	key string
	err error
}

func (e *templateConversionError) Error() string {
	return fmt.Sprintf("%s: %v", e.key, e.err)
}

func (e *templateConversionError) Unwrap() error {
	return e.err
}

func templateAliasRequests(control spec.FormSpecControl) ([]templateAliasRequest, error) {
	requests := make([]templateAliasRequest, 0, len(control.Properties))
	add := func(canonical, key string, raw any, normalized any) {
		requests = append(requests, templateAliasRequest{canonical: canonical, key: key, raw: raw, normalized: normalized})
	}
	if control.Caption != nil {
		value, err := convertValue(control.Type, "caption", *control.Caption)
		if err != nil {
			return nil, &templateConversionError{key: "caption", err: err}
		}
		add("Caption", "caption", *control.Caption, value)
	}
	for _, field := range []struct {
		name, key string
		value     *string
	}{{"Tag", "tag", control.Tag}, {"ControlTipText", "controlTipText", control.ControlTipText}, {"Accelerator", "accelerator", control.Accelerator}} {
		if field.value != nil {
			add(field.name, field.key, *field.value, *field.value)
		}
	}
	if control.Text != nil {
		value, err := convertValue(control.Type, "text", *control.Text)
		if err != nil {
			return nil, &templateConversionError{key: "text", err: err}
		}
		add("Value", "text", *control.Text, value)
	}
	if control.Value != nil {
		value, err := convertValue(control.Type, "value", control.Value)
		if err != nil {
			return nil, &templateConversionError{key: "value", err: err}
		}
		add("Value", "value", control.Value, value)
	}
	for _, field := range []struct {
		canonical string
		key       string
		raw       any
	}{
		{"Left", "left", pointerValue(control.Left)},
		{"Top", "top", pointerValue(control.Top)},
		{"Width", "width", pointerValue(control.Width)},
		{"Height", "height", pointerValue(control.Height)},
		{"TabIndex", "tabindex", pointerValue(control.TabIndex)},
		{"Enabled", "enabled", pointerValue(control.Enabled)},
		{"Visible", "visible", pointerValue(control.Visible)},
	} {
		if field.raw == nil {
			continue
		}
		conversionKey := field.key
		if conversionKey == "tabindex" {
			conversionKey = "tabIndex"
		}
		value, err := convertValue(control.Type, conversionKey, field.raw)
		if err != nil {
			return nil, &templateConversionError{key: field.key, err: err}
		}
		add(field.canonical, field.key, field.raw, value)
	}
	for _, key := range slices.Sorted(maps.Keys(control.Properties)) {
		canonical, known := templateCanonicalProperty(key)
		if !known {
			continue
		}
		value, err := templateNormalizeBagValue(control.Type, key, control.Properties[key])
		if err != nil {
			return nil, &templateConversionError{key: "properties." + key, err: err}
		}
		add(canonical, key, control.Properties[key], value)
	}
	return requests, nil
}
func groupTemplateAliases(requests []templateAliasRequest) map[string][]templateAliasRequest {
	groups := make(map[string][]templateAliasRequest)
	for _, request := range requests {
		groups[request.canonical] = append(groups[request.canonical], request)
	}
	return groups
}

func applyTemplateAlias(control *spec.FormSpecControl, request templateAliasRequest) {
	switch request.canonical {
	case "Tag":
		if value, ok := request.raw.(string); ok {
			control.Tag = new(value)
		}
	case "ControlTipText":
		if value, ok := request.raw.(string); ok {
			control.ControlTipText = new(value)
		}
	case "Accelerator":
		if value, ok := request.raw.(string); ok {
			control.Accelerator = new(value)
		}
	case "Caption":
		if value, ok := request.raw.(string); ok {
			control.Caption = new(value)
		}
	case "Value":
		control.Value = request.raw
		if value, ok := request.raw.(string); ok && (request.key == "text" || strings.EqualFold(control.Type, "TextBox")) {
			control.Text = new(value)
		}
	case "Left":
		if value, ok := request.raw.(float64); ok {
			control.Left = new(value)
		}
	case "Top":
		if value, ok := request.raw.(float64); ok {
			control.Top = new(value)
		}
	case "Width":
		if value, ok := request.raw.(float64); ok {
			control.Width = new(value)
		}
	case "Height":
		if value, ok := request.raw.(float64); ok {
			control.Height = new(value)
		}
	case "TabIndex":
		if value, ok := request.normalized.(int16); ok {
			control.TabIndex = new(int(value))
		} else if value, ok := request.raw.(int); ok {
			control.TabIndex = new(value)
		}
	case "Enabled":
		if value, ok := request.raw.(bool); ok {
			control.Enabled = new(value)
		}
	case "Visible":
		if value, ok := request.raw.(bool); ok {
			control.Visible = new(value)
		}
	}
}

func templateCanonicalProperty(key string) (string, bool) {
	canonical, ok := propertyNames[strings.ToLower(strings.TrimSpace(key))]
	return canonical, ok
}

func templateNormalizeBagValue(kind, key string, value any) (any, error) {
	return bagValue(kind, strings.ToLower(strings.TrimSpace(key)), value)
}

func cloneFormSpec(input spec.FormSpec) (spec.FormSpec, error) {
	body, err := json.Marshal(input)
	if err != nil {
		return spec.FormSpec{}, err
	}
	var clone spec.FormSpec
	if err := json.Unmarshal(body, &clone); err != nil {
		return spec.FormSpec{}, err
	}
	clonePictureData(input.Controls, clone.Controls)
	return clone, nil
}

func clonePictureData(input, output []spec.FormSpecControl) {
	for index := range input {
		if index >= len(output) {
			return
		}
		if input[index].Picture != nil {
			if output[index].Picture == nil {
				pictureAction := *input[index].Picture
				output[index].Picture = &pictureAction
			}
			output[index].Picture.Data = bytes.Clone(input[index].Picture.Data)
		}
		clonePictureData(input[index].Controls, output[index].Controls)
	}
}

func reconcileTemplatePictureWarnings(warnings []spec.FormSpecWarning, controls []spec.FormSpecControl) []spec.FormSpecWarning {
	explicit := make(map[string]bool)
	for _, control := range controls {
		if control.Picture != nil {
			explicit[strings.ToLower(control.Name)] = true
		}
	}
	if len(explicit) == 0 {
		return warnings
	}
	const prefix = "Unsupported Designer properties were preserved only in the binary model: "
	result := make([]spec.FormSpecWarning, 0, len(warnings))
	for _, warning := range warnings {
		if warning.Code != "unsupported_properties" || !explicit[strings.ToLower(warning.Control)] || !strings.HasPrefix(warning.Message, prefix) || !strings.HasSuffix(warning.Message, ".") {
			result = append(result, warning)
			continue
		}
		properties := strings.Split(strings.TrimSuffix(strings.TrimPrefix(warning.Message, prefix), "."), ", ")
		properties = slices.DeleteFunc(properties, func(value string) bool { return strings.EqualFold(strings.TrimSpace(value), "picture") })
		if len(properties) == 0 {
			continue
		}
		warning.Message = prefix + strings.Join(properties, ", ") + "."
		result = append(result, warning)
	}
	return result
}

func templateError(base *oforms.Form, code, control, property, reason string) error {
	form := ""
	if base != nil {
		form = base.Name
	}
	return &Error{Code: code, Form: form, Control: control, Property: property, Reason: reason}
}
