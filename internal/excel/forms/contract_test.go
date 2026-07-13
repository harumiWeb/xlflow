package forms

import "testing"

func TestUserFormContractBuiltInControlsAndProgIDs(t *testing.T) {
	tests := []struct {
		typeName           string
		progID             string
		canContainChildren bool
	}{
		{"Label", "Forms.Label.1", false},
		{"TextBox", "Forms.TextBox.1", false},
		{"ComboBox", "Forms.ComboBox.1", false},
		{"ListBox", "Forms.ListBox.1", false},
		{"CommandButton", "Forms.CommandButton.1", false},
		{"CheckBox", "Forms.CheckBox.1", false},
		{"OptionButton", "Forms.OptionButton.1", false},
		{"Frame", "Forms.Frame.1", true},
	}

	contract := UserFormContract()
	if contract.SchemaVersion != 1 {
		t.Fatalf("SchemaVersion = %d, want 1", contract.SchemaVersion)
	}
	for _, tc := range tests {
		t.Run(tc.typeName, func(t *testing.T) {
			control, ok := LookupControlContract(tc.typeName)
			if !ok {
				t.Fatalf("LookupControlContract(%q) missing", tc.typeName)
			}
			if control.Type != tc.typeName {
				t.Fatalf("control type = %q, want %q", control.Type, tc.typeName)
			}
			if control.ProgID != tc.progID {
				t.Fatalf("ProgID = %q, want %q", control.ProgID, tc.progID)
			}
			if control.CanContainChildren != tc.canContainChildren {
				t.Fatalf("CanContainChildren = %v, want %v", control.CanContainChildren, tc.canContainChildren)
			}
			progID, ok := BuiltInControlProgID(tc.typeName)
			if !ok || progID != tc.progID {
				t.Fatalf("BuiltInControlProgID(%q) = %q, %v; want %q, true", tc.typeName, progID, ok, tc.progID)
			}
			byProgID, ok := LookupControlContractByProgID(tc.progID)
			if !ok || byProgID.Type != tc.typeName {
				t.Fatalf("LookupControlContractByProgID(%q) = %#v, %v; want %q, true", tc.progID, byProgID, ok, tc.typeName)
			}
			if !ProgIDMatchesControlType(tc.typeName, tc.progID) {
				t.Fatalf("ProgIDMatchesControlType(%q, %q) = false", tc.typeName, tc.progID)
			}
		})
	}
}

func TestUserFormContractDocumentAndCommonProperties(t *testing.T) {
	contract := UserFormContract()
	tests := []struct {
		name          string
		valueType     ValueType
		required      bool
		supportLevel  SupportLevel
		allowedValues []string
	}{
		{"schemaVersion", ValueTypeInteger, true, SupportLevelSupported, []string{"1"}},
		{"kind", ValueTypeString, true, SupportLevelSupported, []string{"xlflow.userform"}},
		{"basis", ValueTypeString, true, SupportLevelSupported, []string{"designer"}},
		{"coordinateSystem", ValueTypeString, false, SupportLevelSupported, []string{"points", "parent-relative"}},
		{"warnings", ValueTypeObjectArray, false, SupportLevelSnapshotOnly, nil},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			property, ok := contract.DocumentProperties[tc.name]
			if !ok {
				t.Fatalf("document property %q missing", tc.name)
			}
			assertProperty(t, property, tc.valueType, tc.required, tc.supportLevel)
			for _, want := range tc.allowedValues {
				if !containsString(property.AllowedValues, want) {
					t.Fatalf("property %q allowed values = %#v, want %q", tc.name, property.AllowedValues, want)
				}
			}
		})
	}

	commonTests := []struct {
		name         string
		valueType    ValueType
		required     bool
		supportLevel SupportLevel
	}{
		{"id", ValueTypeString, true, SupportLevelSupported},
		{"name", ValueTypeString, true, SupportLevelSupported},
		{"type", ValueTypeString, true, SupportLevelSupported},
		{"progId", ValueTypeString, false, SupportLevelSupported},
		{"parentId", ValueTypeString, false, SupportLevelSupported},
		{"zIndex", ValueTypeInteger, false, SupportLevelSupported},
		{"left", ValueTypeNumber, false, SupportLevelSupported},
		{"top", ValueTypeNumber, false, SupportLevelSupported},
		{"width", ValueTypeNumber, false, SupportLevelSupported},
		{"height", ValueTypeNumber, false, SupportLevelSupported},
		{"tabIndex", ValueTypeInteger, false, SupportLevelSupported},
		{"enabled", ValueTypeBoolean, false, SupportLevelSupported},
		{"visible", ValueTypeBoolean, false, SupportLevelSupported},
	}
	for _, tc := range commonTests {
		t.Run(tc.name, func(t *testing.T) {
			property, ok := LookupControlProperty("TextBox", tc.name)
			if !ok {
				t.Fatalf("common control property %q missing", tc.name)
			}
			assertProperty(t, property, tc.valueType, tc.required, tc.supportLevel)
		})
	}
}

func TestUserFormContractTypeSpecificProperties(t *testing.T) {
	tests := []struct {
		typeName     string
		propertyName string
		valueType    ValueType
		supportLevel SupportLevel
	}{
		{"Label", "caption", ValueTypeString, SupportLevelSupported},
		{"TextBox", "text", ValueTypeString, SupportLevelSupported},
		{"TextBox", "value", ValueTypeAny, SupportLevelSupported},
		{"ComboBox", "list", ValueTypeStringArray, SupportLevelObservedOnly},
		{"ComboBox", "selectedIndex", ValueTypeInteger, SupportLevelObservedOnly},
		{"ListBox", "list", ValueTypeStringArray, SupportLevelObservedOnly},
		{"CommandButton", "caption", ValueTypeString, SupportLevelSupported},
		{"CheckBox", "value", ValueTypeAny, SupportLevelSupported},
		{"OptionButton", "value", ValueTypeAny, SupportLevelSupported},
		{"Frame", "caption", ValueTypeString, SupportLevelSupported},
	}
	for _, tc := range tests {
		t.Run(tc.typeName+"_"+tc.propertyName, func(t *testing.T) {
			property, ok := LookupControlProperty(tc.typeName, tc.propertyName)
			if !ok {
				t.Fatalf("LookupControlProperty(%q, %q) missing", tc.typeName, tc.propertyName)
			}
			assertProperty(t, property, tc.valueType, false, tc.supportLevel)
			if tc.propertyName != "id" && len(property.ApplicableControls) > 0 && !containsString(property.ApplicableControls, tc.typeName) {
				t.Fatalf("ApplicableControls = %#v, want %q", property.ApplicableControls, tc.typeName)
			}
		})
	}

	if property, ok := LookupControlProperty("Label", "list"); ok {
		t.Fatalf("Label.list should not be applicable, got %#v", property)
	}
	if property, ok := LookupControlProperty("UnknownControl", "caption"); ok {
		t.Fatalf("unknown type-specific property should not be claimed, got %#v", property)
	}
	if property, ok := LookupControlProperty("UnknownControl", "id"); !ok || property.SupportLevel != SupportLevelSupported {
		t.Fatalf("custom controls should receive common structural property support, got %#v, %v", property, ok)
	}
}

func TestUserFormContractSupportLevelsAreRepresented(t *testing.T) {
	contract := UserFormContract()
	seen := map[SupportLevel]bool{}
	for _, property := range contract.DocumentProperties {
		seen[property.SupportLevel] = true
	}
	for _, property := range contract.FormProperties {
		seen[property.SupportLevel] = true
	}
	for _, property := range contract.CommonControlProperties {
		seen[property.SupportLevel] = true
	}
	for _, control := range contract.Controls {
		for _, property := range control.Properties {
			seen[property.SupportLevel] = true
		}
	}

	for _, supportLevel := range []SupportLevel{
		SupportLevelSupported,
		SupportLevelBestEffort,
		SupportLevelObservedOnly,
		SupportLevelSnapshotOnly,
		SupportLevelCustomUnchecked,
	} {
		if !seen[supportLevel] {
			t.Fatalf("support level %q is not represented", supportLevel)
		}
	}
}

func TestControlProgIDCompatibilityAndMismatchDetection(t *testing.T) {
	progID, err := ControlProgID(FormSpecControl{Type: "TextBox"})
	if err != nil {
		t.Fatal(err)
	}
	if progID != "Forms.TextBox.1" {
		t.Fatalf("ControlProgID(TextBox) = %q", progID)
	}

	customProgID, err := ControlProgID(FormSpecControl{Type: "VendorWidget", ProgID: "Vendor.Widget.1"})
	if err != nil {
		t.Fatal(err)
	}
	if customProgID != "Vendor.Widget.1" {
		t.Fatalf("custom progID = %q", customProgID)
	}

	mismatchedProgID, err := ControlProgID(FormSpecControl{Type: "TextBox", ProgID: "Forms.Label.1"})
	if err != nil {
		t.Fatal(err)
	}
	if mismatchedProgID != "Forms.Label.1" {
		t.Fatalf("explicit mismatched progID = %q", mismatchedProgID)
	}
	if ProgIDMatchesControlType("TextBox", mismatchedProgID) {
		t.Fatal("ProgIDMatchesControlType should detect a known type/custom progID mismatch")
	}

	if _, err := ControlProgID(FormSpecControl{Type: "VendorWidget"}); err == nil {
		t.Fatal("unknown type without explicit progID should still fail")
	}
}

func TestValidateFormSpecRejectsKnownNonContainerParent(t *testing.T) {
	spec := FormSpec{
		SchemaVersion: 1,
		Kind:          "xlflow.userform",
		Basis:         "designer",
		Form:          FormSpecForm{Name: "UserForm1"},
		Controls: []FormSpecControl{
			{ID: "txt", Name: "TextBox1", Type: "TextBox"},
			{ID: "lbl", ParentID: "txt", Name: "Label1", Type: "Label"},
		},
	}
	err := ValidateFormSpec(spec)
	if err == nil {
		t.Fatal("expected non-container parent validation error")
	}
	specErr, ok := err.(*SpecError)
	if !ok {
		t.Fatalf("expected SpecError, got %T", err)
	}
	if specErr.Field != "controls[1].parentId" {
		t.Fatalf("field = %q", specErr.Field)
	}
}

func TestValidateFormSpecAllowsFrameAndCustomParent(t *testing.T) {
	frameSpec := FormSpec{
		SchemaVersion: 1,
		Kind:          "xlflow.userform",
		Basis:         "designer",
		Form:          FormSpecForm{Name: "UserForm1"},
		Controls: []FormSpecControl{
			{ID: "frame", Name: "Frame1", Type: "Frame"},
			{ID: "lbl", ParentID: "frame", Name: "Label1", Type: "Label"},
		},
	}
	if err := ValidateFormSpec(frameSpec); err != nil {
		t.Fatalf("Frame parent should be valid: %v", err)
	}

	customSpec := FormSpec{
		SchemaVersion: 1,
		Kind:          "xlflow.userform",
		Basis:         "designer",
		Form:          FormSpecForm{Name: "UserForm1"},
		Controls: []FormSpecControl{
			{ID: "custom", Name: "Custom1", Type: "VendorWidget", ProgID: "Vendor.Widget.1"},
			{ID: "lbl", ParentID: "custom", Name: "Label1", Type: "Label"},
		},
	}
	if err := ValidateFormSpec(customSpec); err != nil {
		t.Fatalf("custom parent should remain unchecked for compatibility: %v", err)
	}
}

func TestValidateFormSpecHonorsExplicitContainerProgIDOverride(t *testing.T) {
	spec := FormSpec{
		SchemaVersion: 1,
		Kind:          "xlflow.userform",
		Basis:         "designer",
		Form:          FormSpecForm{Name: "UserForm1"},
		Controls: []FormSpecControl{
			{ID: "frame", Name: "Frame1", Type: "TextBox", ProgID: "Forms.Frame.1"},
			{ID: "lbl", ParentID: "frame", Name: "Label1", Type: "Label"},
		},
	}
	if err := ValidateFormSpec(spec); err != nil {
		t.Fatalf("explicit container progID should override non-container type for validation: %v", err)
	}
	canContainChildren, known := FormSpecControlCanContainChildren(spec.Controls[0])
	if !known || !canContainChildren {
		t.Fatalf("FormSpecControlCanContainChildren = %v, %v; want true, true", canContainChildren, known)
	}
}

func TestValidateFormSpecRejectsExplicitKnownNonContainerProgIDOverride(t *testing.T) {
	spec := FormSpec{
		SchemaVersion: 1,
		Kind:          "xlflow.userform",
		Basis:         "designer",
		Form:          FormSpecForm{Name: "UserForm1"},
		Controls: []FormSpecControl{
			{ID: "parent", Name: "Parent1", Type: "Frame", ProgID: "Forms.TextBox.1"},
			{ID: "lbl", ParentID: "parent", Name: "Label1", Type: "Label"},
		},
	}
	err := ValidateFormSpec(spec)
	if err == nil {
		t.Fatal("expected explicit known non-container progID to be rejected")
	}
	specErr, ok := err.(*SpecError)
	if !ok {
		t.Fatalf("expected SpecError, got %T", err)
	}
	if specErr.Field != "controls[1].parentId" {
		t.Fatalf("field = %q", specErr.Field)
	}
}

func assertProperty(t *testing.T, property PropertyContract, valueType ValueType, required bool, supportLevel SupportLevel) {
	t.Helper()
	if property.ValueType != valueType {
		t.Fatalf("ValueType = %q, want %q", property.ValueType, valueType)
	}
	if property.Required != required {
		t.Fatalf("Required = %v, want %v", property.Required, required)
	}
	if property.SupportLevel != supportLevel {
		t.Fatalf("SupportLevel = %q, want %q", property.SupportLevel, supportLevel)
	}
	if property.Description == "" {
		t.Fatal("Description is empty")
	}
}

func containsString(values []string, want string) bool {
	for _, value := range values {
		if value == want {
			return true
		}
	}
	return false
}
