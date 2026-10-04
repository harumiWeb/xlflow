package spec

import (
	"slices"
	"strings"
)

type ValueType string

const (
	ValueTypeString      ValueType = "string"
	ValueTypeNumber      ValueType = "number"
	ValueTypeInteger     ValueType = "integer"
	ValueTypeBoolean     ValueType = "boolean"
	ValueTypeAny         ValueType = "any"
	ValueTypeStringArray ValueType = "stringArray"
	ValueTypeObject      ValueType = "object"
	ValueTypeObjectArray ValueType = "objectArray"
)

type SupportLevel string

const (
	SupportLevelSupported       SupportLevel = "supported"
	SupportLevelBestEffort      SupportLevel = "best-effort"
	SupportLevelObservedOnly    SupportLevel = "observed-only"
	SupportLevelSnapshotOnly    SupportLevel = "snapshot-only"
	SupportLevelCustomUnchecked SupportLevel = "custom/unchecked"
)

type PropertyContract struct {
	ValueType          ValueType
	Required           bool
	SupportLevel       SupportLevel
	Description        string
	IncludeInAuthoring bool
	AllowedValues      []string
	ApplicableControls []string
}

type ControlContract struct {
	Type                     string
	ProgID                   string
	CanContainChildren       bool
	AllowedChildTypes        []string
	AllowedParentTypes       []string
	ExcludedCommonProperties []string
	Properties               map[string]PropertyContract
}

type Contract struct {
	SchemaVersion           int
	DocumentProperties      map[string]PropertyContract
	FormProperties          map[string]PropertyContract
	CommonControlProperties map[string]PropertyContract
	TabProperties           map[string]PropertyContract
	Controls                map[string]ControlContract
}

var userFormContract = newUserFormContract()

func UserFormContract() Contract {
	return cloneContract(userFormContract)
}

func LookupControlContract(typeName string) (ControlContract, bool) {
	control, ok := userFormContract.Controls[contractKey(typeName)]
	if !ok {
		return ControlContract{}, false
	}
	return cloneControlContract(control), true
}

func LookupControlProperty(typeName, propertyName string) (PropertyContract, bool) {
	property, ok := lookupProperty(ControlProperties(typeName), propertyName)
	if !ok {
		return PropertyContract{}, false
	}
	return clonePropertyContract(property), true
}

// ControlProperties returns the properties available on the control type,
// combining common and type-specific properties while applying exclusions.
func ControlProperties(typeName string) map[string]PropertyContract {
	properties := clonePropertyMap(userFormContract.CommonControlProperties)
	control, ok := userFormContract.Controls[contractKey(typeName)]
	if !ok {
		return properties
	}
	for _, name := range control.ExcludedCommonProperties {
		deleteProperty(properties, name)
	}
	for name, property := range control.Properties {
		properties[name] = clonePropertyContract(property)
	}
	return properties
}

func deleteProperty(properties map[string]PropertyContract, name string) {
	for key := range properties {
		if strings.EqualFold(key, name) {
			delete(properties, key)
		}
	}
}

// LookupTabProperty returns metadata for one standalone TabStrip tab field.
func LookupTabProperty(propertyName string) (PropertyContract, bool) {
	property, ok := lookupProperty(userFormContract.TabProperties, propertyName)
	return clonePropertyContract(property), ok
}

// LookupFormBuildProperty returns metadata for the explicit form.build
// override object. It is shared by source validation and editor intelligence.
func LookupFormBuildProperty(propertyName string) (PropertyContract, bool) {
	property, ok := lookupProperty(formBuildProperties(), propertyName)
	return clonePropertyContract(property), ok
}

// LookupFormObservedProperty returns metadata for captured form.observed
// fields. These fields are snapshots rather than normal authoring inputs.
func LookupFormObservedProperty(propertyName string) (PropertyContract, bool) {
	property, ok := lookupProperty(formObservedProperties(), propertyName)
	return clonePropertyContract(property), ok
}

// LookupPictureActionProperty returns metadata for fields nested under an
// Image control's picture action object.
func LookupPictureActionProperty(propertyName string) (PropertyContract, bool) {
	property, ok := lookupProperty(pictureActionProperties(), propertyName)
	if !ok {
		return PropertyContract{}, false
	}
	return clonePropertyContract(property), true
}

func PictureActionProperties() map[string]PropertyContract {
	return clonePropertyMap(pictureActionProperties())
}

func pictureActionProperties() map[string]PropertyContract {
	return map[string]PropertyContract{
		"path":   property(ValueTypeString, false, SupportLevelSupported, "Project-root-relative BMP or JPEG asset path.", true),
		"remove": propertyWithAllowed(ValueTypeBoolean, false, SupportLevelSupported, "Clear the existing Image picture.", true, "true"),
	}
}

// LookupObservedControlProperty returns snapshot metadata for fields nested
// under controls[*].observed. The same field names may be build-supported at a
// control's top level, but their observed copies describe captured state.
func LookupObservedControlProperty(typeName, propertyName string) (PropertyContract, bool) {
	contract := UserFormContract()
	for _, name := range []string{"left", "top", "width", "height", "tabIndex", "enabled", "visible", "unsupported", "properties"} {
		if !strings.EqualFold(name, propertyName) {
			continue
		}
		property, ok := lookupProperty(contract.CommonControlProperties, name)
		return observedProperty(property), ok
	}
	if control, ok := LookupControlContract(typeName); ok {
		if property, found := lookupProperty(control.Properties, propertyName); found {
			return observedProperty(property), true
		}
	}
	return PropertyContract{}, false
}

func observedProperty(property PropertyContract) PropertyContract {
	property.SupportLevel = SupportLevelSnapshotOnly
	property.IncludeInAuthoring = false
	property.ApplicableControls = append([]string(nil), property.ApplicableControls...)
	return property
}

func BuiltInControlProgID(typeName string) (string, bool) {
	control, ok := userFormContract.Controls[contractKey(typeName)]
	if !ok {
		return "", false
	}
	return control.ProgID, true
}

func LookupControlContractByProgID(progID string) (ControlContract, bool) {
	normalized := strings.TrimSpace(progID)
	if normalized == "" {
		return ControlContract{}, false
	}
	for _, control := range userFormContract.Controls {
		if strings.EqualFold(normalized, control.ProgID) {
			return cloneControlContract(control), true
		}
	}
	return ControlContract{}, false
}

func ControlCanContainChildren(typeName string) bool {
	control, ok := userFormContract.Controls[contractKey(typeName)]
	return ok && control.CanContainChildren
}

func FormSpecControlCanContainChildren(control FormSpecControl) (bool, bool) {
	if progID := strings.TrimSpace(control.ProgID); progID != "" {
		if contract, ok := LookupControlContractByProgID(progID); ok {
			return contract.CanContainChildren, true
		}
		return false, false
	}
	contract, ok := LookupControlContract(control.Type)
	if !ok {
		return false, false
	}
	return contract.CanContainChildren, true
}

// FormSpecControlParentAllowsChild applies the built-in parent/child
// constraints in the canonical UserForm contract. The bool result indicates
// whether both controls are known well enough to make a decision.
func FormSpecControlParentAllowsChild(parent, child FormSpecControl) (bool, bool) {
	parentContract, parentKnown := formSpecControlContract(parent)
	childContract, childKnown := formSpecControlContract(child)
	if !parentKnown {
		if childKnown && childContract.AllowedParentTypes != nil {
			return false, true
		}
		return false, false
	}
	if !parentContract.CanContainChildren {
		return false, true
	}
	if parentContract.AllowedChildTypes != nil && (!childKnown || !containsControlType(parentContract.AllowedChildTypes, childContract.Type)) {
		return false, true
	}
	if childKnown && childContract.AllowedParentTypes != nil && !containsControlType(childContract.AllowedParentTypes, parentContract.Type) {
		return false, true
	}
	return true, true
}

func formSpecControlContract(control FormSpecControl) (ControlContract, bool) {
	if progID := strings.TrimSpace(control.ProgID); progID != "" {
		if contract, ok := LookupControlContractByProgID(progID); ok {
			return contract, true
		}
	}
	return LookupControlContract(control.Type)
}

func containsControlType(values []string, typeName string) bool {
	for _, value := range values {
		if strings.EqualFold(value, typeName) {
			return true
		}
	}
	return false
}

func ProgIDMatchesControlType(typeName, progID string) bool {
	builtInProgID, ok := BuiltInControlProgID(typeName)
	return ok && strings.EqualFold(strings.TrimSpace(progID), builtInProgID)
}

func newUserFormContract() Contract {
	common := map[string]PropertyContract{
		"id":          property(ValueTypeString, true, SupportLevelSupported, "Stable control identifier used by parentId references.", true),
		"name":        property(ValueTypeString, true, SupportLevelSupported, "VBA control name.", true),
		"type":        property(ValueTypeString, true, SupportLevelSupported, "Canonical xlflow control type.", true),
		"progId":      property(ValueTypeString, false, SupportLevelSupported, "COM ProgID used when creating the control. Explicit values override the built-in type mapping.", false),
		"parentId":    property(ValueTypeString, false, SupportLevelSupported, "ID of the containing control.", true),
		"zIndex":      property(ValueTypeInteger, false, SupportLevelSupported, "Sibling ordering hint for Designer rebuild.", true),
		"left":        property(ValueTypeNumber, false, SupportLevelSupported, "Left coordinate relative to the parent container.", true),
		"top":         property(ValueTypeNumber, false, SupportLevelSupported, "Top coordinate relative to the parent container.", true),
		"width":       property(ValueTypeNumber, false, SupportLevelSupported, "Control width in Designer points.", true),
		"height":      property(ValueTypeNumber, false, SupportLevelSupported, "Control height in Designer points.", true),
		"tabIndex":    property(ValueTypeInteger, false, SupportLevelSupported, "Control tab order index.", true),
		"enabled":     property(ValueTypeBoolean, false, SupportLevelSupported, "Whether the control is enabled.", true),
		"visible":     property(ValueTypeBoolean, false, SupportLevelSupported, "Whether the control is visible.", true),
		"observed":    property(ValueTypeObject, false, SupportLevelSnapshotOnly, "Snapshot state captured from Excel for review and fallback build intent.", false),
		"controls":    property(ValueTypeObjectArray, false, SupportLevelSnapshotOnly, "Legacy nested input accepted and normalized to flat controls.", false),
		"properties":  property(ValueTypeObject, false, SupportLevelCustomUnchecked, "Unchecked custom property bag for future or non-standard controls.", false),
		"unsupported": property(ValueTypeStringArray, false, SupportLevelSnapshotOnly, "Properties observed during inspection that xlflow did not model.", false),
	}
	tabs := map[string]PropertyContract{
		"name":           property(ValueTypeString, true, SupportLevelSupported, "TabStrip tab name; names must be unique within the owning TabStrip, ignoring case.", true),
		"caption":        property(ValueTypeString, false, SupportLevelSupported, "Text displayed on the tab.", true),
		"controlTipText": property(ValueTypeString, false, SupportLevelSupported, "Tooltip shown for the tab.", true),
		"tag":            property(ValueTypeString, false, SupportLevelSupported, "Opaque metadata associated with the tab.", true),
		"accelerator":    property(ValueTypeString, false, SupportLevelSupported, "Keyboard accelerator for the tab.", true),
		"enabled":        property(ValueTypeBoolean, false, SupportLevelSupported, "Whether the tab is enabled.", true),
		"visible":        property(ValueTypeBoolean, false, SupportLevelSupported, "Whether the tab is visible.", true),
	}
	multiPage := control("MultiPage", "Forms.MultiPage.1", true, map[string]PropertyContract{
		"selectedIndex": property(ValueTypeInteger, false, SupportLevelSupported, "Zero-based selected Page index, or -1 when no Page is selected.", true),
	})
	multiPage.AllowedChildTypes = []string{"Page"}
	page := control("Page", "Forms.Page.1", true, typeProperties(
		"caption", "Text displayed on the page tab.",
		"controlTipText", "Tooltip shown for the page.",
		"tag", "Opaque metadata associated with the page.",
		"accelerator", "Keyboard accelerator for the page.",
	))
	page.AllowedParentTypes = []string{"MultiPage"}
	page.ExcludedCommonProperties = []string{"left", "top", "width", "height"}
	tabStrip := control("TabStrip", "Forms.TabStrip.1", false, map[string]PropertyContract{
		"selectedIndex": property(ValueTypeInteger, false, SupportLevelSupported, "Zero-based selected Tab index, or -1 when no Tab is selected.", true),
		"tabs":          property(ValueTypeObjectArray, false, SupportLevelSupported, "Standalone TabStrip tabs. An explicit empty array removes all tabs.", true),
	})

	return Contract{
		SchemaVersion: 1,
		DocumentProperties: map[string]PropertyContract{
			"schemaVersion":    propertyWithAllowed(ValueTypeInteger, true, SupportLevelSupported, "UserForm spec schema version.", false, "1"),
			"kind":             propertyWithAllowed(ValueTypeString, true, SupportLevelSupported, "Document kind discriminator.", false, "xlflow.userform"),
			"basis":            propertyWithAllowed(ValueTypeString, true, SupportLevelSupported, "Inspection/build basis for the spec.", false, "designer"),
			"coordinateSystem": propertyWithAllowed(ValueTypeString, false, SupportLevelSupported, "Coordinate basis used by geometry fields.", false, "points", "parent-relative"),
			"form":             property(ValueTypeObject, true, SupportLevelSupported, "Top-level UserForm metadata and build intent.", false),
			"controls":         property(ValueTypeObjectArray, true, SupportLevelSupported, "Flat control list.", false),
			"warnings":         property(ValueTypeObjectArray, false, SupportLevelSnapshotOnly, "Snapshot-oriented warnings persisted with the spec.", false),
		},
		FormProperties: map[string]PropertyContract{
			"name":     property(ValueTypeString, true, SupportLevelSupported, "VBA UserForm component name.", true),
			"caption":  property(ValueTypeString, false, SupportLevelSupported, "UserForm caption.", true),
			"width":    property(ValueTypeNumber, false, SupportLevelBestEffort, "Form width; applied through VBComponent properties when available.", true),
			"height":   property(ValueTypeNumber, false, SupportLevelBestEffort, "Form height; applied through VBComponent properties when available.", true),
			"build":    property(ValueTypeObject, false, SupportLevelSupported, "Authoritative build intent for form-level properties.", false),
			"observed": property(ValueTypeObject, false, SupportLevelSnapshotOnly, "Observed form state captured from Excel.", false),
		},
		CommonControlProperties: common,
		TabProperties:           tabs,
		Controls: map[string]ControlContract{
			"label":         control("Label", "Forms.Label.1", false, typeProperties("caption", "Text displayed by the label.")),
			"textbox":       control("TextBox", "Forms.TextBox.1", false, typeProperties("text", "TextBox text.", "value", "TextBox value.")),
			"combobox":      control("ComboBox", "Forms.ComboBox.1", false, listControlProperties()),
			"listbox":       control("ListBox", "Forms.ListBox.1", false, listControlProperties()),
			"commandbutton": control("CommandButton", "Forms.CommandButton.1", false, typeProperties("caption", "Button caption.")),
			"checkbox":      control("CheckBox", "Forms.CheckBox.1", false, typeProperties("caption", "CheckBox caption.", "value", "CheckBox checked state or tri-state value.")),
			"optionbutton":  control("OptionButton", "Forms.OptionButton.1", false, typeProperties("caption", "OptionButton caption.", "value", "OptionButton selected state.")),
			"togglebutton":  control("ToggleButton", "Forms.ToggleButton.1", false, typeProperties("caption", "ToggleButton caption.", "value", "ToggleButton selected state.")),
			"spinbutton":    control("SpinButton", "Forms.SpinButton.1", false, integerValueProperties("SpinButton value.")),
			"scrollbar":     control("ScrollBar", "Forms.ScrollBar.1", false, integerValueProperties("ScrollBar value.")),
			"image": control("Image", "Forms.Image.1", false, map[string]PropertyContract{
				"picture":          property(ValueTypeObject, false, SupportLevelSupported, "Replace or remove an Image picture using a project-root-relative BMP or JPEG asset.", true),
				"pictureAlignment": propertyWithAllowed(ValueTypeInteger, false, SupportLevelSupported, "Image picture alignment (0 through 4).", true, "0", "1", "2", "3", "4"),
				"pictureSizeMode":  propertyWithAllowed(ValueTypeInteger, false, SupportLevelSupported, "Image picture size mode: clip (0), stretch (1), or zoom (3).", true, "0", "1", "3"),
			}),
			"frame":     control("Frame", "Forms.Frame.1", true, typeProperties("caption", "Frame caption.")),
			"multipage": multiPage,
			"page":      page,
			"tabstrip":  tabStrip,
		},
	}
}

func property(valueType ValueType, required bool, supportLevel SupportLevel, description string, includeInAuthoring bool) PropertyContract {
	return PropertyContract{
		ValueType:          valueType,
		Required:           required,
		SupportLevel:       supportLevel,
		Description:        description,
		IncludeInAuthoring: includeInAuthoring,
	}
}

func propertyWithAllowed(valueType ValueType, required bool, supportLevel SupportLevel, description string, includeInAuthoring bool, allowedValues ...string) PropertyContract {
	property := property(valueType, required, supportLevel, description, includeInAuthoring)
	property.AllowedValues = append([]string(nil), allowedValues...)
	return property
}

func control(typeName, progID string, canContainChildren bool, properties map[string]PropertyContract) ControlContract {
	for name, property := range properties {
		property.ApplicableControls = append(property.ApplicableControls, typeName)
		properties[name] = property
	}
	return ControlContract{
		Type:               typeName,
		ProgID:             progID,
		CanContainChildren: canContainChildren,
		Properties:         properties,
	}
}

func typeProperties(entries ...string) map[string]PropertyContract {
	properties := make(map[string]PropertyContract, len(entries)/2)
	for i := 0; i+1 < len(entries); i += 2 {
		name := entries[i]
		properties[name] = property(controlPropertyValueType(name), false, SupportLevelSupported, entries[i+1], true)
	}
	return properties
}

func integerValueProperties(description string) map[string]PropertyContract {
	return map[string]PropertyContract{
		"value": property(ValueTypeInteger, false, SupportLevelSupported, description, true),
	}
}

func listControlProperties() map[string]PropertyContract {
	properties := typeProperties("text", "Displayed text.", "value", "Current control value.")
	properties["list"] = property(ValueTypeStringArray, false, SupportLevelObservedOnly, "Design-time list items captured from Excel; build attempts to apply them best-effort.", true)
	properties["selectedIndex"] = property(ValueTypeInteger, false, SupportLevelObservedOnly, "Selected list index captured from Excel; build attempts to apply it best-effort.", true)
	return properties
}

func controlPropertyValueType(name string) ValueType {
	switch contractKey(name) {
	case "value":
		return ValueTypeAny
	default:
		return ValueTypeString
	}
}

func contractKey(value string) string {
	return strings.ToLower(strings.TrimSpace(value))
}

func lookupProperty(properties map[string]PropertyContract, name string) (PropertyContract, bool) {
	if property, ok := properties[name]; ok {
		return property, true
	}
	normalized := contractKey(name)
	for key, property := range properties {
		if contractKey(key) == normalized {
			return property, true
		}
	}
	return PropertyContract{}, false
}

func cloneContract(contract Contract) Contract {
	return Contract{
		SchemaVersion:           contract.SchemaVersion,
		DocumentProperties:      clonePropertyMap(contract.DocumentProperties),
		FormProperties:          clonePropertyMap(contract.FormProperties),
		CommonControlProperties: clonePropertyMap(contract.CommonControlProperties),
		TabProperties:           clonePropertyMap(contract.TabProperties),
		Controls:                cloneControlMap(contract.Controls),
	}
}

func cloneControlMap(values map[string]ControlContract) map[string]ControlContract {
	if len(values) == 0 {
		return nil
	}
	cloned := make(map[string]ControlContract, len(values))
	for key, value := range values {
		cloned[key] = cloneControlContract(value)
	}
	return cloned
}

func cloneControlContract(value ControlContract) ControlContract {
	return ControlContract{
		Type:                     value.Type,
		ProgID:                   value.ProgID,
		CanContainChildren:       value.CanContainChildren,
		AllowedChildTypes:        slices.Clone(value.AllowedChildTypes),
		AllowedParentTypes:       slices.Clone(value.AllowedParentTypes),
		ExcludedCommonProperties: slices.Clone(value.ExcludedCommonProperties),
		Properties:               clonePropertyMap(value.Properties),
	}
}

func clonePropertyMap(values map[string]PropertyContract) map[string]PropertyContract {
	if len(values) == 0 {
		return nil
	}
	cloned := make(map[string]PropertyContract, len(values))
	for key, value := range values {
		cloned[key] = clonePropertyContract(value)
	}
	return cloned
}

func clonePropertyContract(value PropertyContract) PropertyContract {
	value.AllowedValues = append([]string(nil), value.AllowedValues...)
	value.ApplicableControls = append([]string(nil), value.ApplicableControls...)
	return value
}
