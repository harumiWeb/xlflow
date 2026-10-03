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
	"BooleanProperties":  "booleanProperties",
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
	if err := state.projectControls(form.Controls, "", nil); err != nil {
		return forms.FormSpec{}, err
	}
	return forms.NormalizeFormSpec(result), nil
}

type projectionState struct {
	spec      *forms.FormSpec
	nextID    int
	unnamedID int
}

func (s *projectionState) projectControls(controls []*oforms.Control, parentID string, pageTabs []oforms.Tab) error {
	for index, source := range controls {
		if source == nil {
			continue
		}

		var tabStrip *oforms.TabStrip
		if source.MultiPage != nil {
			if source.MultiPage.Hidden == nil {
				return fmt.Errorf("project MultiPage %q: hidden TabStrip is missing", source.Name)
			}
			var err error
			tabStrip, err = controlTabStrip(source.MultiPage.Hidden)
			if err != nil {
				return fmt.Errorf("project MultiPage %q hidden TabStrip: %w", source.Name, err)
			}
			if len(source.MultiPage.Pages) > 0 && len(tabStrip.Tabs) != len(source.MultiPage.Pages) {
				return fmt.Errorf("project MultiPage %q: %d tabs do not match %d pages", source.Name, len(tabStrip.Tabs), len(source.MultiPage.Pages))
			}
		} else if source.Kind == "MSForms.TabStrip" {
			var err error
			tabStrip, err = controlTabStrip(source)
			if err != nil {
				return fmt.Errorf("project TabStrip %q: %w", source.Name, err)
			}
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
		if pageTabs != nil {
			if index >= len(pageTabs) {
				return fmt.Errorf("project Page %q: hidden TabStrip metadata is missing", source.Name)
			}
			tab := pageTabs[index]
			control.Caption = new(tab.Caption)
			control.ControlTipText = new(tab.ControlTipText)
			control.Accelerator = new(tab.Accelerator)
			control.Enabled = new(tab.Enabled)
			control.Visible = new(tab.Visible)
			if source.Site != nil {
				if tag, ok := source.Site.Strings["Tag"]; ok {
					control.Tag = new(tag.Text)
				}
			}
		}
		if tabStrip != nil {
			if source.MultiPage != nil {
				selectedIndex := int32(0)
				if source.MultiPage.Hidden.Record != nil {
					selectedIndex = int32(source.MultiPage.Hidden.Record.Values["ListIndex"])
				}
				if len(source.MultiPage.Pages) == 0 {
					selectedIndex = -1
				}
				control.SelectedIndex = new(int(selectedIndex))
				if source.MultiPage.Properties != nil {
					control.Enabled = new(source.MultiPage.Properties.Mask&(1<<3) == 0)
				}
			} else {
				control.Tabs = projectTabs(tabStrip.Tabs)
				control.Observed = &forms.FormSpecObservedControl{Tabs: projectTabs(tabStrip.Tabs)}
				selectedIndex := tabStrip.SelectedIndex
				if len(tabStrip.Tabs) == 0 {
					selectedIndex = -1
				}
				control.SelectedIndex = new(int(selectedIndex))
			}
		}

		unsupported := unsupportedControlProperties(source, control.Type)
		unsupported = append(unsupported, unsupportedLevelProperties(source.Level)...)
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

		if source.MultiPage != nil {
			if err := s.projectControls(source.MultiPage.Pages, id, tabStrip.Tabs); err != nil {
				return err
			}
			continue
		}
		if err := s.projectControls(source.Children, id, nil); err != nil {
			return err
		}
	}
	return nil
}
func controlTabStrip(control *oforms.Control) (*oforms.TabStrip, error) {
	if control == nil {
		return nil, fmt.Errorf("control is missing")
	}
	if control.TabStrip != nil {
		return control.TabStrip, nil
	}
	if control.Record == nil {
		return nil, fmt.Errorf("TabStrip record is missing")
	}
	return oforms.ParseTabStrip(control.Record)
}

func projectTabs(tabs []oforms.Tab) []forms.FormSpecTab {
	projected := make([]forms.FormSpecTab, 0, len(tabs))
	for _, tab := range tabs {
		projected = append(projected, forms.FormSpecTab{
			Name:           tab.Name,
			Caption:        new(tab.Caption),
			ControlTipText: new(tab.ControlTipText),
			Tag:            new(tab.Tag),
			Accelerator:    new(tab.Accelerator),
			Enabled:        new(tab.Enabled),
			Visible:        new(tab.Visible),
		})
	}
	return projected
}

func projectForm(level *oforms.Level, result *forms.FormSpec) {
	record := level.Record
	if caption, ok := record.Strings["Caption"]; ok {
		result.Form.Caption = new(caption.Text)
	}
	if size, ok := record.Sizes["DisplayedSize"]; ok {
		clientWidth := points(size.Width)
		clientHeight := points(size.Height)
		result.Form.Observed = &forms.FormSpecObservedForm{
			ClientWidth:  &clientWidth,
			ClientHeight: &clientHeight,
		}
		buildClientWidth := clientWidth
		buildClientHeight := clientHeight
		result.Form.Build = &forms.FormSpecBuildForm{
			ClientWidth:  &buildClientWidth,
			ClientHeight: &buildClientHeight,
		}
	}
	unsupported := unsupportedRecordProperties(record, map[string]bool{
		"BooleanProperties": isDefaultBooleanProperties("Form", record.Values["BooleanProperties"]),
		"Caption":           true, "DisplayedSize": true,
	})
	unsupported = append(unsupported, unsupportedLevelProperties(level)...)
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
		if source.Site.Position != nil && target.Type != "Page" {
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
	if size, ok := record.Sizes["Size"]; ok && target.Type != "Page" {
		target.Width = new(points(size.Width))
		target.Height = new(points(size.Height))
	} else if size, ok := record.Sizes["DisplayedSize"]; ok && target.Type != "Page" {
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
	if position, ok := record.Values["Position"]; ok && supportsNumericControlValue(target.Type) {
		target.Value = int(position)
	}
	if selected, ok := record.Values["ListIndex"]; ok && supportsSelectedIndex(target.Type) {
		target.SelectedIndex = new(int(selected))
	}
	if bits, ok := record.Values["VariousPropertyBits"]; ok {
		target.Enabled = new(bits&(1<<1) != 0)
	} else if record.Type == "Form" {
		// Frame/Page store Enabled in BooleanProperties bit2, not the
		// embedded-control VariousPropertyBits field.
		bits := int64(4)
		if stored, found := record.Values["BooleanProperties"]; found {
			bits = stored
		}
		target.Enabled = new(bits&4 != 0)
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
		if flags, ok := control.Site.Values["BitFlags"]; ok {
			defaultFlags := defaultSiteFlags(controlType)
			if flags&^int64(1<<1) != defaultFlags&^int64(1<<1) {
				unsupported = append(unsupported, "siteFlags")
			}
		}
		for _, name := range []string{"Tag", "ControlTipText", "RuntimeLicKey", "ControlSource", "RowSource"} {
			if controlType == "Page" && (name == "Tag" || name == "ControlTipText") {
				continue
			}
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
		projected["Position"] = supportsNumericControlValue(controlType)
		projected["BooleanProperties"] = isDefaultBooleanProperties(controlType, control.Record.Values["BooleanProperties"])
		if controlType == "TabStrip" {
			for _, name := range []string{"Items", "TipStrings", "TabNames", "Tags", "Accelerators"} {
				projected[name] = true
			}
		}
		unsupported = append(unsupported, unsupportedRecordProperties(control.Record, projected)...)
		if bits, ok := control.Record.Values["VariousPropertyBits"]; ok {
			if defaultBits, known := defaultVariousPropertyBits(controlType); !known || bits&^int64(1<<1) != defaultBits&^int64(1<<1) {
				unsupported = append(unsupported, "variousPropertyBits")
			}
		}
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
	return controlType == "ComboBox" || controlType == "ListBox" || controlType == "TabStrip"
}

func defaultSiteFlags(controlType string) int64 {
	switch controlType {
	case "Frame", "MultiPage", "Page":
		return 0x40023
	case "Label":
		return 0x32
	default:
		return 0x33
	}
}

func defaultVariousPropertyBits(controlType string) (int64, bool) {
	switch controlType {
	case "TextBox":
		// Retain the projection's established Excel-authored TextBox default
		// (WordWrap set). The writer's omission default is independently
		// specified by MS-OFORMS and must not be inferred from this snapshot.
		return 0x2c80481b, true
	case "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton":
		return oforms.DefaultVariousPropertyBits(23)
	case "Image", "SpinButton", "ScrollBar":
		return 0x1b, true
	case "Label":
		return oforms.DefaultVariousPropertyBits(21)
	case "CommandButton":
		return oforms.DefaultVariousPropertyBits(17)
	}
	return 0, false
}

func isDefaultBooleanProperties(controlType string, value int64) bool {
	switch controlType {
	case "Form":
		return value == 0x4004
	case "Frame", "Page":
		return value&^int64(4) == 0x8000
	case "MultiPage":
		return value&^int64(4) == 0xc000
	default:
		return false
	}
}

func unsupportedLevelProperties(level *oforms.Level) []string {
	if level == nil {
		return nil
	}
	unsupported := make([]string, 0)
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
	return unsupported
}

func supportsControlValue(controlType string) bool {
	switch controlType {
	case "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton", "SpinButton", "ScrollBar":
		return true
	default:
		return false
	}
}

func supportsNumericControlValue(controlType string) bool {
	return controlType == "SpinButton" || controlType == "ScrollBar"
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
		if projected[name] {
			continue
		}
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
