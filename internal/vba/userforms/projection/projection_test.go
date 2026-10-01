package projection

import (
	"encoding/json"
	"fmt"
	"math"
	"os"
	"path/filepath"
	"slices"
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
	if got.Form.Width == nil || got.Form.Height == nil || *got.Form.Width <= 0 || *got.Form.Height <= 0 {
		t.Fatalf("form dimensions = %#v", got.Form)
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
