package compiler

import (
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

	matched, err := validateTemplateTopology(base, baseline.Controls, desired.Controls)
	if err != nil {
		return nil, err
	}
	if err := overlayTemplateForm(base, &after.Form, desired.Form); err != nil {
		return nil, err
	}
	for index := range after.Controls {
		if err := overlayTemplateControl(base, &after.Controls[index], matched[index], index); err != nil {
			return nil, err
		}
	}

	result, err := CompileEdits(base, baseline, after, codePage)
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

func validateTemplateTopology(base *oforms.Form, baseline, desired []spec.FormSpecControl) ([]spec.FormSpecControl, error) {
	if len(desired) != len(baseline) {
		return nil, templateError(base, Unsupported, "", "controls", "control addition/removal is not supported")
	}

	baselineByName := make(map[string]int, len(baseline))
	baselineParentByName := make(map[string]string, len(baseline))
	baselineNameByID := make(map[string]string, len(baseline))
	for index, control := range baseline {
		key := strings.ToLower(control.Name)
		if _, exists := baselineByName[key]; exists {
			return nil, templateError(base, Unsupported, control.Name, fmt.Sprintf("controls[%d].name", index), "ambiguous control names")
		}
		baselineByName[key] = index
		baselineNameByID[control.ID] = control.Name
	}
	for _, control := range baseline {
		baselineParentByName[control.Name] = baselineNameByID[control.ParentID]
	}

	desiredByID := make(map[string]spec.FormSpecControl, len(desired))
	desiredByName := make(map[string]spec.FormSpecControl, len(desired))
	for index, control := range desired {
		path := fmt.Sprintf("controls[%d]", index)
		if strings.TrimSpace(control.Name) == "" {
			return nil, templateError(base, Unsupported, "", path+".name", "control name is required")
		}
		nameKey := strings.ToLower(control.Name)
		if previous, exists := desiredByName[nameKey]; exists {
			return nil, templateError(base, Unsupported, control.Name, path+".name", fmt.Sprintf("duplicate control name also appears for %q", previous.Name))
		}
		desiredByName[nameKey] = control
		if strings.TrimSpace(control.ID) == "" {
			return nil, templateError(base, Unsupported, control.Name, path+".id", "control IDs are required to resolve parentId")
		}
		if _, exists := desiredByID[control.ID]; exists {
			return nil, templateError(base, Unsupported, control.Name, path+".id", "duplicate control ID")
		}
		desiredByID[control.ID] = control
	}

	desiredParentByName := make(map[string]string, len(desired))
	for index, control := range desired {
		path := fmt.Sprintf("controls[%d]", index)
		baselineIndex, exists := baselineByName[strings.ToLower(control.Name)]
		if !exists {
			return nil, templateError(base, Unsupported, control.Name, path, "control addition/removal is not supported")
		}
		actual := baseline[baselineIndex]
		if control.Name != actual.Name || control.Type != actual.Type {
			return nil, templateError(base, Unsupported, control.Name, path, "control identity, type, parent, or ordering change")
		}
		if control.ProgID != "" && !strings.EqualFold(control.ProgID, actual.ProgID) {
			return nil, templateError(base, Unsupported, control.Name, path+".progId", "control identity, type, parent, or ordering change")
		}
		if control.ZIndex != nil && !reflect.DeepEqual(control.ZIndex, actual.ZIndex) {
			return nil, templateError(base, Unsupported, control.Name, path+".zIndex", "control identity, type, parent, or ordering change")
		}

		parentName := ""
		if control.ParentID != "" {
			parent, ok := desiredByID[control.ParentID]
			if !ok {
				return nil, templateError(base, Unsupported, control.Name, path+".parentId", "parent control was not found")
			}
			parentName = parent.Name
		}
		if parentName != baselineParentByName[actual.Name] {
			return nil, templateError(base, Unsupported, control.Name, path+".parentId", "control identity, type, parent, or ordering change")
		}
		desiredParentByName[control.Name] = parentName
	}
	parents := map[string]struct{}{"": {}}
	for _, parent := range baselineParentByName {
		parents[parent] = struct{}{}
	}
	for _, parent := range desiredParentByName {
		parents[parent] = struct{}{}
	}
	for _, parent := range slices.Sorted(maps.Keys(parents)) {
		if !slices.Equal(siblingNames(baseline, baselineParentByName, parent), siblingNames(desired, desiredParentByName, parent)) {
			return nil, templateError(base, Unsupported, "", "controls", "control identity, type, parent, or ordering change")
		}
	}
	matched := make([]spec.FormSpecControl, len(baseline))
	for index, control := range baseline {
		matched[index] = desiredByName[strings.ToLower(control.Name)]
	}
	return matched, nil
}

func siblingNames(controls []spec.FormSpecControl, parentByName map[string]string, parentName string) []string {
	result := make([]string, 0)
	for _, control := range controls {
		if parentByName[control.Name] == parentName {
			result = append(result, control.Name)
		}
	}
	return result
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
	return clone, nil
}

func templateError(base *oforms.Form, code, control, property, reason string) error {
	form := ""
	if base != nil {
		form = base.Name
	}
	return &Error{Code: code, Form: form, Control: control, Property: property, Reason: reason}
}
