// Package projection converts lossless MS-OFORMS persistence models into the
// canonical, user-facing UserForm specification.
package projection

import (
	"fmt"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	forms "github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

const himetricPerInch = 2540.0

var controlIdentities = map[string]struct {
	typeName string
	progID   string
}{
	"MSForms.CheckBox":      {"CheckBox", "Forms.CheckBox.1"},
	"MSForms.ComboBox":      {"ComboBox", "Forms.ComboBox.1"},
	"MSForms.CommandButton": {"CommandButton", "Forms.CommandButton.1"},
	"MSForms.Form":          {"Page", "Forms.Page.1"},
	"MSForms.Frame":         {"Frame", "Forms.Frame.1"},
	"MSForms.Image":         {"Image", "Forms.Image.1"},
	"MSForms.Label":         {"Label", "Forms.Label.1"},
	"MSForms.ListBox":       {"ListBox", "Forms.ListBox.1"},
	"MSForms.MultiPage":     {"MultiPage", "Forms.MultiPage.1"},
	"MSForms.OptionButton":  {"OptionButton", "Forms.OptionButton.1"},
	"MSForms.ScrollBar":     {"ScrollBar", "Forms.ScrollBar.1"},
	"MSForms.SpinButton":    {"SpinButton", "Forms.SpinButton.1"},
	"MSForms.TabStrip":      {"TabStrip", "Forms.TabStrip.1"},
	"MSForms.TextBox":       {"TextBox", "Forms.TextBox.1"},
	"MSForms.ToggleButton":  {"ToggleButton", "Forms.ToggleButton.1"},
}

var unsupportedScalarProperties = map[string]string{
	"Accelerator":        "accelerator",
	"BackColor":          "backColor",
	"BorderColor":        "borderColor",
	"BorderStyle":        "borderStyle",
	"BoundColumn":        "boundColumn",
	"ColumnCount":        "columnCount",
	"Cycle":              "cycle",
	"Delay":              "delay",
	"DropButtonStyle":    "dropButtonStyle",
	"ForeColor":          "foreColor",
	"HelpContextID":      "helpContextId",
	"LargeChange":        "largeChange",
	"ListRows":           "listRows",
	"ListStyle":          "listStyle",
	"ListWidth":          "listWidth",
	"MatchEntry":         "matchEntry",
	"Max":                "max",
	"MaxLength":          "maxLength",
	"Min":                "min",
	"MousePointer":       "mousePointer",
	"MultiSelect":        "multiSelect",
	"Orientation":        "orientation",
	"PasswordChar":       "passwordChar",
	"PictureAlignment":   "pictureAlignment",
	"PicturePosition":    "picturePosition",
	"PictureSizeMode":    "pictureSizeMode",
	"Position":           "position",
	"ProportionalThumb":  "proportionalThumb",
	"ScrollBars":         "scrollBars",
	"ShowDropButtonWhen": "showDropButtonWhen",
	"SmallChange":        "smallChange",
	"SpecialEffect":      "specialEffect",
	"TabFixedHeight":     "tabFixedHeight",
	"TabFixedWidth":      "tabFixedWidth",
	"TabOrientation":     "tabOrientation",
	"TabStyle":           "tabStyle",
	"TakeFocusOnClick":   "takeFocusOnClick",
	"TextColumn":         "textColumn",
	"Zoom":               "zoom",
}

// Project converts a parsed form without mutating or exposing its retained
// binary persistence state.
func Project(form *oforms.Form) (forms.FormSpec, error) {
	if form == nil {
		return forms.FormSpec{}, fmt.Errorf("project MS-OFORMS form: nil form")
	}
	if strings.TrimSpace(form.Name) == "" {
		return forms.FormSpec{}, fmt.Errorf("project MS-OFORMS form: form name is empty")
	}
	if len(form.Levels) == 0 || form.Levels[0] == nil || form.Levels[0].Record == nil {
		return forms.FormSpec{}, fmt.Errorf("project MS-OFORMS form %q: root form record is missing", form.Name)
	}

	result := forms.FormSpec{
		SchemaVersion:    1,
		Kind:             "xlflow.userform",
		Basis:            "designer",
		CoordinateSystem: "parent-relative",
		Form:             forms.FormSpecForm{Name: form.Name},
		Controls:         []forms.FormSpecControl{},
		Warnings:         []forms.FormSpecWarning{},
	}
	projectForm(form.Levels[0], &result)

	state := projectionState{spec: &result}
	state.projectControls(form.Controls, "")
	return forms.NormalizeFormSpec(result), nil
}

type projectionState struct {
	spec      *forms.FormSpec
	nextID    int
	unnamedID int
}

func (s *projectionState) projectControls(controls []*oforms.Control, parentID string) {
	for index, source := range controls {
		if source == nil {
			continue
		}
		s.nextID++
		id := fmt.Sprintf("control_%03d", s.nextID)
		control := forms.FormSpecControl{
			ID:       id,
			ParentID: parentID,
			ZIndex:   new(index),
			Name:     source.Name,
		}
		if strings.TrimSpace(control.Name) == "" {
			s.unnamedID++
			control.Name = fmt.Sprintf("<unnamed_%d>", s.unnamedID)
			s.spec.Warnings = append(s.spec.Warnings, forms.FormSpecWarning{
				Code:    "unnamed_control_placeholder",
				Message: "A control without a stable name was persisted with a generated placeholder name.",
				Control: control.Name,
			})
		}

		identity, known := controlIdentities[source.Kind]
		if known {
			control.Type = identity.typeName
			control.ProgID = identity.progID
		} else {
			control.Type = forms.UnsupportedControlPlaceholderType
			control.Unsupported = []string{forms.UnsupportedControlTypeProperty}
			s.spec.Warnings = append(s.spec.Warnings, forms.FormSpecWarning{
				Code:    "unsupported_control_type",
				Message: fmt.Sprintf("Control type %q is structurally valid but cannot be represented as an authorable FormSpec control.", source.Kind),
				Control: control.Name,
			})
		}

		projectControl(source, &control)
		unsupported := unsupportedControlProperties(source, control.Type)
		if !known {
			unsupported = append(unsupported, forms.UnsupportedControlTypeProperty)
		}
		slices.Sort(unsupported)
		unsupported = slices.Compact(unsupported)
		control.Unsupported = unsupported
		if len(unsupported) > 0 {
			s.spec.Warnings = append(s.spec.Warnings, forms.FormSpecWarning{
				Code:    "unsupported_properties",
				Message: fmt.Sprintf("Unsupported Designer properties were preserved only in the binary model: %s.", strings.Join(unsupported, ", ")),
				Control: control.Name,
			})
		}

		s.spec.Controls = append(s.spec.Controls, control)
		s.projectControls(source.Children, id)
	}
}

func projectForm(level *oforms.Level, result *forms.FormSpec) {
	record := level.Record
	if caption, ok := record.Strings["Caption"]; ok {
		result.Form.Caption = new(caption.Text)
	}
	if size, ok := record.Sizes["DisplayedSize"]; ok {
		result.Form.Width = new(points(size.Width))
		result.Form.Height = new(points(size.Height))
	}
	unsupported := unsupportedRecordProperties(record, map[string]bool{
		"Caption": true, "DisplayedSize": true,
	})
	if len(level.MouseIconRaw) > 0 {
		unsupported = append(unsupported, "mouseIcon")
	}
	if len(level.FontRaw) > 0 {
		unsupported = append(unsupported, "font")
	}
	if len(level.PictureRaw) > 0 {
		unsupported = append(unsupported, "picture")
	}
	for name := range level.ExtraStreams {
		unsupported = append(unsupported, "extraStream:"+name)
	}
	slices.Sort(unsupported)
	unsupported = slices.Compact(unsupported)
	if len(unsupported) > 0 {
		result.Warnings = append(result.Warnings, forms.FormSpecWarning{
			Code:    "unsupported_properties",
			Message: fmt.Sprintf("Unsupported form Designer properties were preserved only in the binary model: %s.", strings.Join(unsupported, ", ")),
		})
	}
}

func projectControl(source *oforms.Control, target *forms.FormSpecControl) {
	if source.Site != nil {
		if source.Site.Position != nil {
			target.Left = new(points(source.Site.Position.Left))
			target.Top = new(points(source.Site.Position.Top))
		}
		flags := int64(0x33)
		if value, ok := source.Site.Values["BitFlags"]; ok {
			flags = value
		}
		target.Visible = new(flags&(1<<1) != 0)
	}
	if source.TabIndex != nil {
		target.TabIndex = new(int(*source.TabIndex))
	}
	if source.Record == nil {
		return
	}
	record := source.Record
	if size, ok := record.Sizes["Size"]; ok {
		target.Width = new(points(size.Width))
		target.Height = new(points(size.Height))
	} else if size, ok := record.Sizes["DisplayedSize"]; ok {
		target.Width = new(points(size.Width))
		target.Height = new(points(size.Height))
	}
	if caption, ok := record.Strings["Caption"]; ok {
		target.Caption = new(caption.Text)
	}
	if value, ok := record.Strings["Value"]; ok && supportsControlValue(target.Type) {
		target.Value = projectedControlValue(target.Type, value.Text)
		if strings.EqualFold(target.Type, "TextBox") {
			target.Text = new(value.Text)
		}
	}
	if selected, ok := record.Values["ListIndex"]; ok && supportsSelectedIndex(target.Type) {
		target.SelectedIndex = new(int(selected))
	}
	if bits, ok := record.Values["VariousPropertyBits"]; ok {
		target.Enabled = new(bits&(1<<1) != 0)
	} else {
		target.Enabled = new(true)
	}
}

func unsupportedControlProperties(control *oforms.Control, controlType string) []string {
	unsupported := make([]string, 0)
	if control.Site != nil {
		if _, ok := control.Site.Values["HelpContextID"]; ok {
			unsupported = append(unsupported, "helpContextId")
		}
		for _, name := range []string{"Tag", "ControlTipText", "RuntimeLicKey", "ControlSource", "RowSource"} {
			if value, ok := control.Site.Strings[name]; ok && value.Text != "" {
				unsupported = append(unsupported, lowerFirst(name))
			}
		}
	}
	if control.Record != nil {
		projected := map[string]bool{
			"Caption": true, "DisplayedSize": true,
			"Size": true, "VariousPropertyBits": true,
		}
		projected["ListIndex"] = supportsSelectedIndex(controlType)
		projected["Value"] = supportsControlValue(controlType)
		unsupported = append(unsupported, unsupportedRecordProperties(control.Record, projected)...)
		if _, ok := control.Record.Values["ListIndex"]; ok && !supportsSelectedIndex(controlType) {
			unsupported = append(unsupported, "selectedIndex")
		}
	}
	if len(control.OpaqueRaw) > 0 {
		unsupported = append(unsupported, "opaqueControlData")
	}
	return unsupported
}

func supportsSelectedIndex(controlType string) bool {
	return controlType == "ComboBox" || controlType == "ListBox"
}

func supportsControlValue(controlType string) bool {
	switch controlType {
	case "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton":
		return true
	default:
		return false
	}
}

func projectedControlValue(controlType, value string) any {
	switch controlType {
	case "CheckBox", "OptionButton", "ToggleButton":
		switch strings.ToLower(strings.TrimSpace(value)) {
		case "1", "-1", "true":
			return "True"
		case "0", "false":
			return "False"
		}
	}
	return value
}

func unsupportedRecordProperties(record *oforms.Record, projected map[string]bool) []string {
	if record == nil {
		return nil
	}
	unsupported := make([]string, 0)
	for name := range record.Values {
		if projected[name] {
			continue
		}
		if public, ok := unsupportedScalarProperties[name]; ok {
			unsupported = append(unsupported, public)
		}
	}
	for name := range record.Strings {
		if projected[name] {
			continue
		}
		unsupported = append(unsupported, lowerFirst(name))
	}
	for name := range record.Sizes {
		if projected[name] {
			continue
		}
		unsupported = append(unsupported, lowerFirst(name))
	}
	for name := range record.Arrays {
		unsupported = append(unsupported, lowerFirst(name))
	}
	for name := range record.Pictures {
		unsupported = append(unsupported, lowerFirst(name))
	}
	if hasPersistedTextProperties(record.TextProps) {
		unsupported = append(unsupported, "textProps")
	}
	if len(record.TailRaw) > 0 {
		unsupported = append(unsupported, "opaqueRecordTail")
	}
	return unsupported
}

func hasPersistedTextProperties(record *oforms.Record) bool {
	return record != nil && (record.Mask != 0 || len(record.Values) != 0 || len(record.Strings) != 0 ||
		len(record.Sizes) != 0 || len(record.Arrays) != 0 || len(record.Pictures) != 0 || len(record.TailRaw) != 0)
}

func points(value int32) float64 {
	return float64(value) * 72.0 / himetricPerInch
}

func lowerFirst(value string) string {
	if value == "" {
		return ""
	}
	return strings.ToLower(value[:1]) + value[1:]
}
