package projection

import (
	"encoding/json"
	"fmt"
	"math"
	"os"
	"path/filepath"
	"slices"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	forms "github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func TestProjectSimpleFixture(t *testing.T) {
	form := readFixture(t, "p4_form.bin")
	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	if got.SchemaVersion != 1 || got.Kind != "xlflow.userform" || got.Basis != "designer" {
		t.Fatalf("document identity = %#v", got)
	}
	if got.CoordinateSystem != "parent-relative" || got.Form.Name != "UserForm1" {
		t.Fatalf("form identity = %#v", got.Form)
	}
	if got.Form.Width != nil || got.Form.Height != nil {
		t.Fatalf("legacy form dimensions should remain absent, got %#v/%#v", got.Form.Width, got.Form.Height)
	}
	if got.Form.Observed == nil || got.Form.Observed.Width != nil || got.Form.Observed.Height != nil || got.Form.Observed.InsideWidth != nil || got.Form.Observed.InsideHeight != nil || got.Form.Observed.ClientWidth == nil || got.Form.Observed.ClientHeight == nil || *got.Form.Observed.ClientWidth <= 0 || *got.Form.Observed.ClientHeight <= 0 {
		t.Fatalf("observed form dimensions = %#v", got.Form.Observed)
	}
	if got.Form.Build == nil || got.Form.Build.Width != nil || got.Form.Build.Height != nil || got.Form.Build.ClientWidth == nil || got.Form.Build.ClientHeight == nil || *got.Form.Build.ClientWidth <= 0 || *got.Form.Build.ClientHeight <= 0 {
		t.Fatalf("build form dimensions = %#v", got.Form.Build)
	}
	if len(got.Controls) == 0 {
		t.Fatal("projected form has no controls")
	}
	button := got.Controls[0]
	if button.ID != fmt.Sprintf("control_%03d", 1) || button.Name != "CommandButton1" || button.Type != "CommandButton" || button.ProgID != "Forms.CommandButton.1" {
		t.Fatalf("first control = %#v", button)
	}
	if button.Caption == nil || *button.Caption != "CommandButton1" {
		t.Fatalf("button caption = %#v", button.Caption)
	}
	if button.Width == nil || math.Abs(*button.Width-126) > 0.001 {
		t.Fatalf("button width = %#v, want 126 points", button.Width)
	}
	wantHeight := 2328 * 72.0 / 2540.0
	if button.Height == nil || math.Abs(*button.Height-wantHeight) > 0.001 {
		t.Fatalf("button height = %#v, want %v points", button.Height, wantHeight)
	}
	if button.Visible == nil || !*button.Visible || button.Enabled == nil || !*button.Enabled {
		t.Fatalf("button visible/enabled = %#v/%#v", button.Visible, button.Enabled)
	}
	if issues := forms.ValidateFormSpecStrict(got); hasErrors(issues) {
		t.Fatalf("projected fixture did not validate: %#v", issues)
	}
}

func TestProjectNestedFixturePreservesParentsAndSiblingOrder(t *testing.T) {
	got, err := Project(readFixture(t, "p6_nested_form.bin"))
	if err != nil {
		t.Fatal(err)
	}
	byID := make(map[string]forms.FormSpecControl, len(got.Controls))
	children := 0
	for _, control := range got.Controls {
		byID[control.ID] = control
		if control.ParentID != "" {
			children++
			if _, ok := byID[control.ParentID]; !ok {
				t.Fatalf("control %q appeared before parent %q", control.ID, control.ParentID)
			}
		}
	}
	if children == 0 {
		t.Fatal("nested fixture projected no parent relationships")
	}
	containerTypes := []string{"Frame", "MultiPage", "Page"}
	for _, control := range got.Controls {
		if control.Enabled == nil {
			t.Fatalf("control %q omitted the persisted default enabled state", control.Name)
		}
		if control.ParentID == "" {
			continue
		}
		parent := byID[control.ParentID]
		if !slices.Contains(containerTypes, parent.Type) && !forms.IsUnsupportedControlPlaceholder(parent) {
			t.Fatalf("control %q has unexpected parent type %q", control.Name, parent.Type)
		}
	}
	if issues := forms.ValidateFormSpecStrict(got); hasErrors(issues) {
		t.Fatalf("projected nested fixture did not validate: %#v", issues)
	}
	path := filepath.Join(t.TempDir(), "NestedForm.json")
	if err := forms.WriteSnapshot(forms.SnapshotOutput{Path: path, DisplayPath: path, Format: "json"}, got); err != nil {
		t.Fatalf("WriteSnapshot: %v", err)
	}
	if _, err := forms.LoadFormSpec(forms.SpecInput{Path: path, DisplayPath: path, Format: "json"}); err != nil {
		t.Fatalf("LoadFormSpec: %v", err)
	}
}

func TestProjectOpaqueControlUsesSnapshotOnlyPlaceholder(t *testing.T) {
	form := &oforms.Form{
		Name: "OpaqueForm",
		Levels: []*oforms.Level{{
			Record: &oforms.Record{
				Type: "Form", Values: map[string]int64{}, Strings: map[string]oforms.StoredString{},
				Sizes: map[string]oforms.Size{}, Arrays: map[string][]byte{}, Pictures: map[string][]byte{},
			},
		}},
		Controls: []*oforms.Control{{
			Name: "VendorControl1", Kind: "ActiveX.Control", OpaqueRaw: []byte{1, 2, 3},
			Site: &oforms.Site{
				Values: map[string]int64{"BitFlags": 0x33}, Strings: map[string]oforms.StoredString{},
				Position: &oforms.Position{Left: 2540, Top: -1270},
			},
		}},
	}
	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	control := got.Controls[0]
	if !forms.IsUnsupportedControlPlaceholder(control) {
		t.Fatalf("control = %#v, want snapshot placeholder", control)
	}
	if control.Left == nil || *control.Left != 72 || control.Top == nil || *control.Top != -36 {
		t.Fatalf("placeholder geometry = left %#v top %#v", control.Left, control.Top)
	}
	if !slices.Equal(control.Unsupported, []string{"controlType", "opaqueControlData"}) {
		t.Fatalf("unsupported = %q", control.Unsupported)
	}
	if issues := forms.ValidateFormSpecStrict(got); hasErrors(issues) {
		t.Fatalf("snapshot placeholder should validate with warnings: %#v", issues)
	}
	input := forms.SpecInput{DisplayPath: "OpaqueForm.yaml", Format: "yaml"}
	if err := forms.ValidateFormSpecForAuthoring(input, got); err == nil {
		t.Fatal("snapshot placeholder was accepted for authoring")
	}
}

func TestProjectDoesNotExposeUnsupportedCommandButtonValue(t *testing.T) {
	form := &oforms.Form{
		Name:   "ButtonForm",
		Levels: []*oforms.Level{{Record: emptyRecord("Form")}},
		Controls: []*oforms.Control{{
			Name: "CommandButton1", Kind: "MSForms.CommandButton",
			Record: &oforms.Record{
				Type: "CommandButton", Values: map[string]int64{},
				Strings: map[string]oforms.StoredString{"Value": {Text: "False"}},
				Sizes:   map[string]oforms.Size{}, Arrays: map[string][]byte{}, Pictures: map[string][]byte{},
			},
		}},
	}
	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	control := got.Controls[0]
	if control.Value != nil {
		t.Fatalf("CommandButton value = %#v, want omitted", control.Value)
	}
	if !slices.Contains(control.Unsupported, "value") {
		t.Fatalf("unsupported = %q, want value", control.Unsupported)
	}
	if issues := forms.ValidateFormSpecStrict(got); hasErrors(issues) {
		t.Fatalf("projected CommandButton did not validate: %#v", issues)
	}
}

func TestProjectNewControlValuesAndDefaultFlags(t *testing.T) {
	spin := emptyRecord("SpinButton")
	spin.Values["Position"] = 7
	spin.Values["VariousPropertyBits"] = 0x1b
	scroll := emptyRecord("ScrollBar")
	scroll.Values["Position"] = 42
	scroll.Values["VariousPropertyBits"] = 0x1b
	toggle := emptyRecord("ToggleButton")
	toggle.Strings["Value"] = oforms.StoredString{Text: "1"}
	toggle.Values["VariousPropertyBits"] = 0x2c80081b
	image := emptyRecord("Image")
	image.Values["VariousPropertyBits"] = 0x1b
	form := &oforms.Form{
		Name:   "NewControlsForm",
		Levels: []*oforms.Level{{Record: emptyRecord("Form")}},
		Controls: []*oforms.Control{
			{Name: "ToggleButton1", Kind: "MSForms.ToggleButton", Record: toggle},
			{Name: "SpinButton1", Kind: "MSForms.SpinButton", Record: spin},
			{Name: "ScrollBar1", Kind: "MSForms.ScrollBar", Record: scroll},
			{Name: "Image1", Kind: "MSForms.Image", Record: image},
		},
	}

	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	byName := make(map[string]forms.FormSpecControl, len(got.Controls))
	for _, control := range got.Controls {
		byName[control.Name] = control
	}
	if value := byName["ToggleButton1"].Value; value != "True" {
		t.Fatalf("ToggleButton value = %#v, want True", value)
	}
	for _, test := range []struct {
		name string
		want int
	}{
		{name: "SpinButton1", want: 7},
		{name: "ScrollBar1", want: 42},
	} {
		control := byName[test.name]
		if control.Value != test.want {
			t.Fatalf("%s value = %#v, want %d", test.name, control.Value, test.want)
		}
		if control.Enabled == nil || !*control.Enabled {
			t.Fatalf("%s enabled = %#v, want true", test.name, control.Enabled)
		}
		if slices.Contains(control.Unsupported, "position") || slices.Contains(control.Unsupported, "variousPropertyBits") {
			t.Fatalf("%s unsupported = %q, want Position and default bits projected", test.name, control.Unsupported)
		}
	}
	if imageControl := byName["Image1"]; imageControl.Enabled == nil || !*imageControl.Enabled || slices.Contains(imageControl.Unsupported, "variousPropertyBits") {
		t.Fatalf("Image projection = %#v, want enabled with default bits", imageControl)
	}
}

func TestProjectTabStripSnapshotCanBeWrittenAndLoaded(t *testing.T) {
	form := &oforms.Form{
		Name:   "TabForm",
		Levels: []*oforms.Level{{Record: emptyRecord("Form")}},
		Controls: []*oforms.Control{{
			Name: "TabStrip1", Kind: "MSForms.TabStrip",
			Record: &oforms.Record{Type: "TabStrip", Values: map[string]int64{"ListIndex": 1}},
		}},
	}
	projected, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	if projected.Controls[0].SelectedIndex != nil {
		t.Fatalf("TabStrip selectedIndex = %#v, want omitted unsupported state", projected.Controls[0].SelectedIndex)
	}
	if !slices.Contains(projected.Controls[0].Unsupported, "selectedIndex") {
		t.Fatalf("TabStrip unsupported = %q, want selectedIndex", projected.Controls[0].Unsupported)
	}

	path := filepath.Join(t.TempDir(), "TabForm.json")
	if err := forms.WriteSnapshot(forms.SnapshotOutput{Path: path, DisplayPath: path, Format: "json"}, projected); err != nil {
		t.Fatalf("WriteSnapshot: %v", err)
	}
	loaded, err := forms.LoadFormSpec(forms.SpecInput{Path: path, DisplayPath: path, Format: "json"})
	if err != nil {
		t.Fatalf("LoadFormSpec: %v", err)
	}
	if loaded.Controls[0].SelectedIndex != nil || !slices.Contains(loaded.Controls[0].Unsupported, "selectedIndex") {
		t.Fatalf("loaded TabStrip = %#v, want selectedIndex recorded as unsupported", loaded.Controls[0])
	}
}

func TestProjectReportsUnmodeledSiteStrings(t *testing.T) {
	form := &oforms.Form{
		Name:   "SiteForm",
		Levels: []*oforms.Level{{Record: emptyRecord("Form")}},
		Controls: []*oforms.Control{{
			Name: "TextBox1", Kind: "MSForms.TextBox",
			Site: &oforms.Site{Strings: map[string]oforms.StoredString{
				"Name":           {Text: "TextBox1"},
				"Tag":            {Text: "tag-value"},
				"ControlTipText": {Text: "tip"},
				"RuntimeLicKey":  {Text: "license"},
				"ControlSource":  {Text: "Sheet1!A1"},
				"RowSource":      {Text: "Sheet1!A1:A5"},
			}},
		}},
	}
	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	want := []string{"controlSource", "controlTipText", "rowSource", "runtimeLicKey", "tag"}
	if !slices.Equal(got.Controls[0].Unsupported, want) {
		t.Fatalf("unsupported = %q, want %q", got.Controls[0].Unsupported, want)
	}
	warnings := 0
	for _, warning := range got.Warnings {
		if warning.Code == "unsupported_properties" && warning.Control == "TextBox1" {
			warnings++
			for _, name := range want {
				if !strings.Contains(warning.Message, name) {
					t.Errorf("warning %q omits %q", warning.Message, name)
				}
			}
		}
	}
	if warnings != 1 {
		t.Fatalf("unsupported_properties warnings for TextBox1 = %d, want 1", warnings)
	}
}

func TestProjectReportsUnmodeledControlFlags(t *testing.T) {
	form := &oforms.Form{
		Name:   "FlagsForm",
		Levels: []*oforms.Level{{Record: emptyRecord("Form")}},
		Controls: []*oforms.Control{{
			Name: "TextBox1", Kind: "MSForms.TextBox",
			Site:   &oforms.Site{Values: map[string]int64{"BitFlags": 0x35}},
			Record: &oforms.Record{Type: "TextBox", Values: map[string]int64{"VariousPropertyBits": 0x6}},
		}},
	}
	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	want := []string{"siteFlags", "variousPropertyBits"}
	if !slices.Equal(got.Controls[0].Unsupported, want) {
		t.Fatalf("unsupported = %q, want %q", got.Controls[0].Unsupported, want)
	}
}

func TestProjectReportsNestedLevelAndFormBooleanProperties(t *testing.T) {
	nestedLevel := &oforms.Level{
		Record:       &oforms.Record{Type: "Form", Values: map[string]int64{"BooleanProperties": 2}},
		MouseIconRaw: []byte{1}, FontRaw: []byte{2}, PictureRaw: []byte{3},
		ExtraStreams: map[string][]byte{"custom": {4}},
	}
	form := &oforms.Form{
		Name: "NestedForm",
		Levels: []*oforms.Level{{Record: &oforms.Record{
			Type: "Form", Values: map[string]int64{"BooleanProperties": 1},
		}}},
		Controls: []*oforms.Control{{
			Name: "Frame1", Kind: "MSForms.Frame",
			Record: nestedLevel.Record, Level: nestedLevel,
		}},
	}
	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	if len(got.Warnings) != 2 {
		t.Fatalf("warnings = %#v, want form and Frame warnings", got.Warnings)
	}
	if !strings.Contains(got.Warnings[0].Message, "booleanProperties") {
		t.Fatalf("form warning = %#v, want booleanProperties", got.Warnings[0])
	}
	frame := got.Controls[0]
	want := []string{"booleanProperties", "extraStream:custom", "font", "mouseIcon", "picture"}
	if !slices.Equal(frame.Unsupported, want) {
		t.Fatalf("Frame unsupported = %q, want %q", frame.Unsupported, want)
	}
	if got.Warnings[1].Control != "Frame1" || !strings.Contains(got.Warnings[1].Message, "picture") {
		t.Fatalf("Frame warning = %#v", got.Warnings[1])
	}
}

func TestProjectDoesNotReportPersistedBitfieldDefaults(t *testing.T) {
	form := &oforms.Form{
		Name: "DefaultsForm",
		Levels: []*oforms.Level{{Record: &oforms.Record{
			Type: "Form", Values: map[string]int64{"BooleanProperties": 0x4004},
		}}},
		Controls: []*oforms.Control{
			{
				Name: "Frame1", Kind: "MSForms.Frame",
				Site:   &oforms.Site{Values: map[string]int64{"BitFlags": 0x40023}},
				Record: &oforms.Record{Type: "Frame", Values: map[string]int64{"BooleanProperties": 0x8004}},
			},
			{
				Name: "TextBox1", Kind: "MSForms.TextBox",
				Site:   &oforms.Site{Values: map[string]int64{"BitFlags": 0x33}},
				Record: &oforms.Record{Type: "TextBox", Values: map[string]int64{"VariousPropertyBits": 0x2c80481b}},
			},
		},
	}

	got, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	if len(got.Warnings) != 0 {
		t.Fatalf("warnings = %#v, want no warnings for persisted defaults", got.Warnings)
	}
	for _, control := range got.Controls {
		if len(control.Unsupported) != 0 {
			t.Errorf("%s unsupported = %q, want none for persisted defaults", control.Name, control.Unsupported)
		}
	}
}

func TestProjectedControlValueMatchesExcelBooleanSpelling(t *testing.T) {
	for _, test := range []struct {
		persisted string
		want      string
	}{
		{persisted: "1", want: "True"},
		{persisted: "0", want: "False"},
		{persisted: "-1", want: "True"},
	} {
		if got := projectedControlValue("CheckBox", test.persisted); got != test.want {
			t.Fatalf("projectedControlValue(CheckBox, %q) = %#v, want %q", test.persisted, got, test.want)
		}
	}
	if got := projectedControlValue("TextBox", "1"); got != "1" {
		t.Fatalf("projectedControlValue(TextBox, 1) = %#v, want raw text", got)
	}
}

func TestProjectIsDeterministic(t *testing.T) {
	form := readFixture(t, "p6_nested_form.bin")
	first, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	second, err := Project(form)
	if err != nil {
		t.Fatal(err)
	}
	firstJSON, err := json.Marshal(first)
	if err != nil {
		t.Fatal(err)
	}
	secondJSON, err := json.Marshal(second)
	if err != nil {
		t.Fatal(err)
	}
	if string(firstJSON) != string(secondJSON) {
		t.Fatalf("projection is unstable:\n%s\n%s", firstJSON, secondJSON)
	}
}

func readFixture(t *testing.T, name string) *oforms.Form {
	t.Helper()
	body, err := os.ReadFile(filepath.Join("..", "..", "..", "pack", "vbaproject", "testdata", "corpus", name))
	if err != nil {
		t.Fatal(err)
	}
	container, err := cfb.Open(body)
	if err != nil {
		t.Fatal(err)
	}
	form, err := oforms.ReadForm(container, "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	return form
}

func hasErrors(issues []forms.ValidationIssue) bool {
	for _, issue := range issues {
		if issue.Severity == forms.SeverityError {
			return true
		}
	}
	return false
}

func emptyRecord(recordType string) *oforms.Record {
	return &oforms.Record{
		Type: recordType, Values: map[string]int64{}, Strings: map[string]oforms.StoredString{},
		Sizes: map[string]oforms.Size{}, Arrays: map[string][]byte{}, Pictures: map[string][]byte{},
	}
}
