// Package compiler applies explicit FormSpec differences to an existing,
// lossless MS-OFORMS model. It never opens Excel or consumes FRX artifacts.
package compiler

import (
	"encoding/json"
	"errors"
	"fmt"
	"maps"
	"math"
	"reflect"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

// Error describes a rejected edit. Code is stable; Reason is explanatory text.
type Error struct {
	Code     string
	Form     string
	Control  string
	Property string
	Reason   string
	Cause    error
}

func (e *Error) Error() string {
	return fmt.Sprintf("%s: %s/%s %s: %s", e.Code, e.Form, e.Control, e.Property, e.Reason)
}

func (e *Error) Unwrap() error { return e.Cause }

const (
	Unsupported = "userform_edit_unsupported"
	Invalid     = "userform_edit_invalid"
	Stale       = "userform_edit_stale_input"
	Conflict    = "userform_edit_conflict"
)

type pendingEdit struct {
	oforms.Edit
	path string
}

type compilation struct {
	form     *oforms.Form
	before   spec.FormSpec
	after    spec.FormSpec
	actual   spec.FormSpec
	controls map[string]*oforms.Control
	edits    []pendingEdit
}

// CompileEdits compiles changes between two complete snapshots of the same
// form. Snapshot-only metadata is not authoring input. It returns a newly read,
// signed model; neither snapshots nor base are changed on success or failure.
func CompileEdits(base *oforms.Form, before, after spec.FormSpec, codePage uint16) (*oforms.Form, error) {
	if _, err := oforms.SerializeForm(base, codePage); err != nil {
		return nil, err
	}
	c := compilation{form: base, controls: make(map[string]*oforms.Control)}
	var err error
	c.before, err = normalizedCopy(before)
	if err != nil {
		return nil, c.fail(Invalid, "", "before", err.Error())
	}
	c.after, err = normalizedCopy(after)
	if err != nil {
		return nil, c.fail(Invalid, "", "after", err.Error())
	}
	c.actual, err = projection.Project(base)
	if err != nil {
		return nil, err
	}
	if err := c.topology(); err != nil {
		return nil, err
	}
	if err := c.formEdits(); err != nil {
		return nil, err
	}
	afterByID := make(map[string]spec.FormSpecControl, len(c.after.Controls))
	actualByName := make(map[string]spec.FormSpecControl, len(c.actual.Controls))
	for _, control := range c.after.Controls {
		afterByID[control.ID] = control
	}
	for _, control := range c.actual.Controls {
		actualByName[strings.ToLower(control.Name)] = control
	}
	for index, old := range c.before.Controls {
		if err := c.controlEdits(old, afterByID[old.ID], actualByName[strings.ToLower(old.Name)], fmt.Sprintf("controls[%d]", index)); err != nil {
			return nil, err
		}
	}
	edits := make([]oforms.Edit, len(c.edits))
	for i, edit := range c.edits {
		edits[i] = edit.Edit
	}
	result, err := oforms.ApplyEdits(base, edits, codePage)
	if err != nil {
		code := Invalid
		if errors.Is(err, oforms.ErrUnsupportedEdit) {
			code = Unsupported
		}
		control, path := "", ""
		if detail, ok := errors.AsType[*oforms.EditError](err); ok {
			control, path = detail.Control, detail.Property
			for _, edit := range c.edits {
				if edit.Control == control && edit.Property == path {
					path = edit.path
					break
				}
			}
		}
		return nil, &Error{Code: code, Form: base.Name, Control: control, Property: path, Reason: err.Error(), Cause: err}
	}
	for _, edit := range c.edits {
		var got any
		if edit.Control == "" {
			got = result.Levels[0].Record.Strings[edit.Property].Text
		} else {
			control := findModelControl(result.Controls, edit.Control)
			var found bool
			got, found = editedProperty(control, edit.Property)
			if !found {
				return nil, c.fail(Invalid, edit.Control, edit.path, "edited property is missing after read-back")
			}
		}
		if !reflect.DeepEqual(got, edit.Value) {
			return nil, c.fail(Invalid, edit.Control, edit.path, "edited value differs after read-back")
		}
	}
	return result, nil
}

func findModelControl(controls []*oforms.Control, name string) *oforms.Control {
	for _, control := range controls {
		if control.Name == name {
			return control
		}
		if child := findModelControl(control.Children, name); child != nil {
			return child
		}
	}
	return nil
}

func editedProperty(control *oforms.Control, name string) (any, bool) {
	if control == nil {
		return nil, false
	}
	switch name {
	case "Left", "Top":
		if control.Site.Position == nil {
			return nil, false
		}
		if name == "Left" {
			return control.Site.Position.Left, true
		}
		return control.Site.Position.Top, true
	case "Width", "Height":
		if control.Record == nil {
			return nil, false
		}
		size, found := control.Record.Sizes["Size"]
		if !found {
			size, found = control.Record.Sizes["DisplayedSize"]
		}
		if name == "Width" {
			return size.Width, found
		}
		return size.Height, found
	case "TabIndex":
		if control.TabIndex == nil {
			return nil, false
		}
		return *control.TabIndex, true
	case "Visible":
		v, found := control.Site.Values["BitFlags"]
		return v&(1<<1) != 0, found
	case "Enabled":
		if control.Record == nil {
			return nil, false
		}
		if control.Record.Type == "Form" {
			v, found := control.Record.Values["BooleanProperties"]
			return v&4 != 0, found
		}
		v, found := control.Record.Values["VariousPropertyBits"]
		return v&(1<<1) != 0, found
	default:
		return binaryProperty(control, name)
	}
}

func normalizedCopy(input spec.FormSpec) (spec.FormSpec, error) {
	body, err := json.Marshal(input)
	if err != nil {
		return spec.FormSpec{}, err
	}
	var copied spec.FormSpec
	if err := json.Unmarshal(body, &copied); err != nil {
		return spec.FormSpec{}, err
	}
	// Normalize topology on the owned copy, but do not synthesize form build
	// intent from observed values before detecting differences.
	rawForm := copied.Form
	if copied.Form.Build != nil {
		copied.Form.Build = new(*copied.Form.Build)
	}
	if copied.Form.Observed != nil {
		copied.Form.Observed = new(*copied.Form.Observed)
	}
	copied = spec.NormalizeFormSpec(copied)
	for _, issue := range spec.ValidateFormSpecStrict(copied) {
		if issue.Severity == spec.SeverityError {
			return spec.FormSpec{}, fmt.Errorf("%s: %s", issue.Field, issue.Message)
		}
	}
	copied.Form = rawForm
	return copied, nil
}

func (c *compilation) fail(code, control, path, reason string) error {
	return &Error{Code: code, Form: c.form.Name, Control: control, Property: path, Reason: reason}
}

func (c *compilation) topology() error {
	if c.before.Form.Name != c.form.Name || c.after.Form.Name != c.form.Name {
		return c.fail(Unsupported, "", "form.name", "form rename or mismatched form identity")
	}
	if len(c.before.Controls) != len(c.actual.Controls) || len(c.after.Controls) != len(c.before.Controls) {
		return c.fail(Unsupported, "", "controls", "control addition/removal is not supported")
	}
	var visit func([]*oforms.Control)
	visit = func(controls []*oforms.Control) {
		for _, control := range controls {
			c.controls[strings.ToLower(control.Name)] = control
			visit(control.Children)
		}
	}
	visit(c.form.Controls)
	if len(c.controls) != len(c.actual.Controls) {
		return c.fail(Unsupported, "", "controls", "ambiguous control names")
	}
	oldNames := make(map[string]string)
	actualNames := make(map[string]string)
	for _, control := range c.before.Controls {
		oldNames[control.ID] = control.Name
	}
	for _, control := range c.actual.Controls {
		actualNames[control.ID] = control.Name
	}
	for i, old := range c.before.Controls {
		next, actual := c.after.Controls[i], c.actual.Controls[i]
		path := fmt.Sprintf("controls[%d]", i)
		if old.ID != next.ID || old.Name != next.Name || old.Type != next.Type || old.ProgID != next.ProgID || old.ParentID != next.ParentID || !reflect.DeepEqual(old.ZIndex, next.ZIndex) {
			return c.fail(Unsupported, old.Name, path, "control identity, type, parent, or ordering change")
		}
		if old.Name != actual.Name || !strings.EqualFold(old.Type, actual.Type) || (old.ProgID != "" && !strings.EqualFold(old.ProgID, actual.ProgID)) || oldNames[old.ParentID] != actualNames[actual.ParentID] || !reflect.DeepEqual(old.ZIndex, actual.ZIndex) {
			return c.fail(Stale, old.Name, path, "before topology does not match the binary model")
		}
	}
	return nil
}

func (c *compilation) formEdits() error {
	old, next := c.before.Form, c.after.Form
	oldBuild, nextBuild := spec.FormSpecBuildForm{}, spec.FormSpecBuildForm{}
	if old.Build != nil {
		oldBuild = *old.Build
	}
	if next.Build != nil {
		nextBuild = *next.Build
	}
	if !reflect.DeepEqual(old.Width, next.Width) || !reflect.DeepEqual(old.Height, next.Height) || !reflect.DeepEqual(oldBuild.Width, nextBuild.Width) || !reflect.DeepEqual(oldBuild.Height, nextBuild.Height) {
		return c.fail(Unsupported, "", "form", "root dimensions are client dimensions, not Excel outer dimensions")
	}
	oldCaption, nextCaption := old.Caption, next.Caption
	if !reflect.DeepEqual(oldBuild.Caption, nextBuild.Caption) {
		oldCaption, nextCaption = oldBuild.Caption, nextBuild.Caption
	}
	if reflect.DeepEqual(oldCaption, nextCaption) {
		return nil
	}
	if nextCaption == nil {
		return c.fail(Unsupported, "", "form.caption", "property reset is not supported; use an explicit empty string")
	}
	if oldCaption != nil && (c.actual.Form.Caption == nil && *oldCaption != "" || c.actual.Form.Caption != nil && *oldCaption != *c.actual.Form.Caption) {
		return c.fail(Stale, "", "form.caption", "before caption does not match the binary model")
	}
	return c.add("", "Caption", *nextCaption, "form.caption")
}

func (c *compilation) controlEdits(old, next, actual spec.FormSpecControl, path string) error {
	fields := []struct {
		name, persisted   string
		old, next, actual any
	}{
		{"caption", "Caption", pointerValue(old.Caption), pointerValue(next.Caption), pointerValue(actual.Caption)},
		{"text", "Value", pointerValue(old.Text), pointerValue(next.Text), pointerValue(actual.Text)},
		{"value", "Value", old.Value, next.Value, actual.Value},
		{"left", "Left", pointerValue(old.Left), pointerValue(next.Left), pointerValue(actual.Left)},
		{"top", "Top", pointerValue(old.Top), pointerValue(next.Top), pointerValue(actual.Top)},
		{"width", "Width", pointerValue(old.Width), pointerValue(next.Width), pointerValue(actual.Width)},
		{"height", "Height", pointerValue(old.Height), pointerValue(next.Height), pointerValue(actual.Height)},
		{"tabIndex", "TabIndex", pointerValue(old.TabIndex), pointerValue(next.TabIndex), pointerValue(actual.TabIndex)},
		{"enabled", "Enabled", pointerValue(old.Enabled), pointerValue(next.Enabled), pointerValue(actual.Enabled)},
		{"visible", "Visible", pointerValue(old.Visible), pointerValue(next.Visible), pointerValue(actual.Visible)},
		{"list", "List", old.List, next.List, actual.List},
		{"selectedIndex", "ListIndex", pointerValue(old.SelectedIndex), pointerValue(next.SelectedIndex), pointerValue(actual.SelectedIndex)},
	}
	for _, field := range fields {
		if reflect.DeepEqual(field.old, field.next) {
			continue
		}
		fieldPath := path + "." + field.name
		if !supportedControl(old.Type) {
			return c.fail(Unsupported, old.Name, fieldPath, "control type cannot be edited")
		}
		if !applicableProperty(old.Type, strings.ToLower(field.name)) {
			return c.fail(Unsupported, old.Name, fieldPath, "property is not applicable to this control type")
		}
		if field.next == nil {
			return c.fail(Unsupported, old.Name, fieldPath, "property reset is not supported")
		}
		if field.name == "list" || field.name == "selectedIndex" {
			return c.fail(Unsupported, old.Name, fieldPath, "list/selection persistence is not established")
		}
		if strings.EqualFold(old.Type, "ListBox") && (field.name == "text" || field.name == "value") {
			return c.fail(Unsupported, old.Name, fieldPath, "ListBox value depends on list state that is not persisted for the supported layout")
		}
		if strings.EqualFold(old.Type, "ComboBox") && field.name == "text" {
			field.actual = actual.Value
		}
		previous, err := convertValue(old.Type, field.name, field.old)
		if err != nil {
			return c.fail(Invalid, old.Name, fieldPath, err.Error())
		}
		persisted, err := convertValue(old.Type, field.name, field.actual)
		if err != nil {
			return c.fail(Invalid, old.Name, fieldPath, err.Error())
		}
		if field.old != nil && !sameBaseline(field.name, field.old, field.actual, previous, persisted) {
			return c.fail(Stale, old.Name, fieldPath, "before value does not match the binary model")
		}
		value, err := convertValue(old.Type, field.name, field.next)
		if err != nil {
			return c.fail(Invalid, old.Name, fieldPath, err.Error())
		}
		if err := c.add(old.Name, field.persisted, value, fieldPath); err != nil {
			return err
		}
	}
	return c.propertyEdits(old, next, actual, path)
}

func pointerValue[T any](value *T) any {
	if value == nil {
		return nil
	}
	return *value
}

func sameBaseline(property string, before, actual, convertedBefore, convertedActual any) bool {
	if property == "left" || property == "top" || property == "width" || property == "height" {
		b, beforeOK := before.(float64)
		a, actualOK := actual.(float64)
		// Excel exposes twips while the persistence model retains HIMETRIC.
		// Match the Designer parity tolerance instead of reporting its rounding
		// as a concurrent source edit.
		return beforeOK && actualOK && math.Abs(b-a) <= 0.05
	}
	return reflect.DeepEqual(convertedBefore, convertedActual)
}

func supportedControl(kind string) bool {
	return slices.Contains([]string{"label", "textbox", "combobox", "listbox", "commandbutton", "checkbox", "optionbutton", "frame"}, strings.ToLower(kind))
}

func convertValue(kind, property string, value any) (any, error) {
	if value == nil {
		return nil, nil
	}
	switch property {
	case "left", "top", "width", "height":
		v, ok := value.(float64)
		if !ok || math.IsNaN(v) || math.IsInf(v, 0) || ((property == "width" || property == "height") && v < 0) {
			return nil, fmt.Errorf("invalid geometry")
		}
		h := math.Round(v * 2540 / 72)
		if h < math.MinInt32 || h > math.MaxInt32 {
			return nil, fmt.Errorf("geometry exceeds HIMETRIC range")
		}
		return int32(h), nil
	case "tabIndex":
		v, ok := value.(int)
		if !ok || v < 0 || v > math.MaxInt16 {
			return nil, fmt.Errorf("tabIndex must be in 0..32767")
		}
		return int16(v), nil
	case "value", "text":
		if strings.EqualFold(kind, "CheckBox") || strings.EqualFold(kind, "OptionButton") {
			switch v := value.(type) {
			case bool:
				if v {
					return "1", nil
				}
				return "0", nil
			case float64:
				if v == 1 || v == -1 {
					return "1", nil
				}
				if v == 0 {
					return "0", nil
				}
			case string:
				switch strings.ToLower(strings.TrimSpace(v)) {
				case "true", "1", "-1":
					return "1", nil
				case "false", "0":
					return "0", nil
				}
			}
			return nil, fmt.Errorf("checked state must be Boolean or 0/1/-1; tri-state is not supported")
		}
		if v, ok := value.(string); ok {
			return v, nil
		}
		return nil, fmt.Errorf("text/value must be a string for this control")
	default:
		return value, nil
	}
}

var propertyNames = map[string]string{
	"tag": "Tag", "controltiptext": "ControlTipText", "groupname": "GroupName",
	"backcolor": "BackColor", "forecolor": "ForeColor", "bordercolor": "BorderColor",
	"borderstyle": "BorderStyle", "maxlength": "MaxLength",
	"caption": "Caption", "text": "Value", "value": "Value",
	"left": "Left", "top": "Top", "width": "Width", "height": "Height",
	"tabindex": "TabIndex", "enabled": "Enabled", "visible": "Visible",
}

func propertyMap(properties map[string]any) (map[string]any, error) {
	result := make(map[string]any, len(properties))
	for _, key := range slices.Sorted(maps.Keys(properties)) {
		value := properties[key]
		name := strings.ToLower(strings.TrimSpace(key))
		if _, duplicate := result[name]; duplicate {
			return nil, fmt.Errorf("duplicate property alias %q", key)
		}
		result[name] = value
	}
	return result, nil
}

func (c *compilation) propertyEdits(old, next, actual spec.FormSpecControl, path string) error {
	previous, err := propertyMap(old.Properties)
	if err != nil {
		return c.fail(Conflict, old.Name, path+".properties", err.Error())
	}
	desired, err := propertyMap(next.Properties)
	if err != nil {
		return c.fail(Conflict, old.Name, path+".properties", err.Error())
	}
	keys := make([]string, 0, len(previous)+len(desired))
	for key := range previous {
		keys = append(keys, key)
	}
	for key := range desired {
		keys = append(keys, key)
	}
	slices.Sort(keys)
	keys = slices.Compact(keys)
	for _, key := range keys {
		oldValue, newValue := previous[key], desired[key]
		if reflect.DeepEqual(oldValue, newValue) {
			continue
		}
		fieldPath := path + ".properties." + key
		property, known := propertyNames[key]
		if !known || !supportedControl(old.Type) {
			return c.fail(Unsupported, old.Name, fieldPath, "property is not supported for pure-Go editing")
		}
		if !applicableProperty(old.Type, key) {
			return c.fail(Unsupported, old.Name, fieldPath, "property is not applicable to this control type")
		}
		if newValue == nil {
			return c.fail(Unsupported, old.Name, fieldPath, "property reset is not supported")
		}
		value, err := bagValue(old.Type, key, newValue)
		if err != nil {
			return c.fail(Invalid, old.Name, fieldPath, err.Error())
		}
		if oldValue != nil {
			prior, err := bagValue(old.Type, key, oldValue)
			if err != nil {
				return c.fail(Invalid, old.Name, fieldPath, err.Error())
			}
			actualValue, ok := snapshotProperty(actual, key)
			persisted := actualValue
			if ok {
				persisted, err = convertValue(old.Type, canonicalField(key), persisted)
				if err != nil {
					return c.fail(Invalid, old.Name, fieldPath, err.Error())
				}
			} else {
				persisted, ok = binaryProperty(c.controls[strings.ToLower(old.Name)], property)
			}
			if !ok || !sameBaseline(key, oldValue, actualValue, prior, persisted) {
				return c.fail(Stale, old.Name, fieldPath, "before property does not match the binary model")
			}
		}
		if err := c.add(old.Name, property, value, fieldPath); err != nil {
			return err
		}
	}
	return nil
}

func canonicalField(key string) string {
	if key == "tabindex" {
		return "tabIndex"
	}
	return key
}

func applicableProperty(kind, key string) bool {
	switch key {
	case "groupname":
		return strings.EqualFold(kind, "OptionButton")
	case "maxlength":
		return strings.EqualFold(kind, "TextBox") || strings.EqualFold(kind, "ComboBox")
	case "caption", "text", "value":
		if strings.EqualFold(kind, "ListBox") && (key == "text" || key == "value") {
			return false
		}
		_, ok := spec.LookupControlProperty(kind, key)
		return ok
	default:
		return true
	}
}

func snapshotProperty(control spec.FormSpecControl, key string) (any, bool) {
	switch key {
	case "caption":
		return pointerValue(control.Caption), true
	case "text":
		if strings.EqualFold(control.Type, "ComboBox") {
			return control.Value, true
		}
		return pointerValue(control.Text), true
	case "value":
		return control.Value, true
	case "left":
		return pointerValue(control.Left), true
	case "top":
		return pointerValue(control.Top), true
	case "width":
		return pointerValue(control.Width), true
	case "height":
		return pointerValue(control.Height), true
	case "tabindex":
		return pointerValue(control.TabIndex), true
	case "enabled":
		return pointerValue(control.Enabled), true
	case "visible":
		return pointerValue(control.Visible), true
	default:
		return nil, false
	}
}

func bagValue(kind, key string, value any) (any, error) {
	switch key {
	case "tag", "controltiptext", "groupname", "caption":
		v, ok := value.(string)
		if !ok {
			return nil, fmt.Errorf("property must be a string")
		}
		return v, nil
	case "backcolor", "forecolor", "bordercolor", "borderstyle", "maxlength":
		v, ok := value.(float64)
		if !ok || math.IsNaN(v) || math.IsInf(v, 0) || v != math.Trunc(v) || v < 0 || v > math.MaxUint32 {
			return nil, fmt.Errorf("property must be an unsigned 32-bit integer")
		}
		if key == "borderstyle" && v > 1 {
			return nil, fmt.Errorf("borderStyle must be 0 or 1")
		}
		if key == "maxlength" && v > math.MaxInt32 {
			return nil, fmt.Errorf("maxLength exceeds Excel range")
		}
		return int64(v), nil
	case "enabled", "visible":
		v, ok := value.(bool)
		if !ok {
			return nil, fmt.Errorf("property must be Boolean")
		}
		return v, nil
	case "tabindex":
		v, ok := value.(float64)
		if !ok || v != math.Trunc(v) || v < 0 || v > math.MaxInt16 {
			return nil, fmt.Errorf("tabIndex must be in 0..32767")
		}
		return int16(v), nil
	default:
		return convertValue(kind, key, value)
	}
}

func binaryProperty(control *oforms.Control, name string) (any, bool) {
	if control == nil {
		return nil, false
	}
	if name == "Tag" || name == "ControlTipText" {
		v := control.Site.Strings[name]
		return v.Text, true
	}
	if control.Record == nil {
		return nil, false
	}
	if v, ok := control.Record.Strings[name]; ok {
		return v.Text, true
	}
	if name == "GroupName" {
		return "", true
	}
	v, ok := control.Record.Values[name]
	return v, ok
}

func (c *compilation) add(control, property string, value any, path string) error {
	for _, previous := range c.edits {
		if previous.Control != control || previous.Property != property {
			continue
		}
		if reflect.DeepEqual(previous.Value, value) {
			return nil
		}
		return c.fail(Conflict, control, path, "aliases request different values for the same persisted property")
	}
	c.edits = append(c.edits, pendingEdit{Edit: oforms.Edit{Control: control, Property: property, Value: value}, path: path})
	return nil
}
