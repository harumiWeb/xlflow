package forms

import (
	"encoding/json"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"gopkg.in/yaml.v3"
)

type contractSnapshot struct {
	BuiltInControls []struct {
		Type   string `json:"type"`
		ProgID string `json:"progId"`
	} `json:"builtInControls"`
	Bridge struct {
		FormFields    []string `json:"formFields"`
		ControlFields []string `json:"controlFields"`
	} `json:"bridge"`
}

func TestRepresentativeUserFormFixtureMatchesCanonicalContract(t *testing.T) {
	fixture := readFormsFixture(t, "representative-userform.yaml")
	issues, err := ValidateFormSpecSource(SpecInput{Format: "yaml", DisplayPath: "representative-userform.yaml"}, fixture)
	if err != nil {
		t.Fatal(err)
	}
	if hasValidationErrors(issues) {
		t.Fatalf("representative fixture issues = %#v", issues)
	}
	var spec FormSpec
	if err := yaml.Unmarshal(fixture, &spec); err != nil {
		t.Fatal(err)
	}
	if len(spec.Controls) != 8 {
		t.Fatalf("fixture controls = %d, want 8", len(spec.Controls))
	}
	seen := map[string]FormSpecControl{}
	for _, control := range spec.Controls {
		seen[control.Type] = control
		progID, ok := BuiltInControlProgID(control.Type)
		if !ok || control.ProgID != progID {
			t.Fatalf("%s progId = %q, want %q", control.Type, control.ProgID, progID)
		}
	}
	for _, typeName := range []string{"Frame", "Label", "TextBox", "ComboBox", "ListBox", "CommandButton", "CheckBox", "OptionButton"} {
		if _, ok := seen[typeName]; !ok {
			t.Fatalf("fixture missing %s", typeName)
		}
	}
	if seen["Label"].ParentID != "frame_main" {
		t.Fatalf("nested Label parentId = %q", seen["Label"].ParentID)
	}
}

func TestContractSnapshotMatchesGoContractAndSpecTags(t *testing.T) {
	var snapshot contractSnapshot
	if err := json.Unmarshal(readFormsFixture(t, "contract-snapshot.json"), &snapshot); err != nil {
		t.Fatal(err)
	}
	contract := UserFormContract()
	if len(snapshot.BuiltInControls) != len(contract.Controls) {
		t.Fatalf("snapshot built-in controls = %d, contract controls = %d", len(snapshot.BuiltInControls), len(contract.Controls))
	}
	for _, expected := range snapshot.BuiltInControls {
		control, ok := LookupControlContract(expected.Type)
		if !ok || control.ProgID != expected.ProgID {
			t.Fatalf("contract %s = %#v, want progId %q", expected.Type, control, expected.ProgID)
		}
	}
	assertJSONFields(t, reflect.TypeOf(FormSpecForm{}), snapshot.Bridge.FormFields)
	assertJSONFields(t, reflect.TypeOf(FormSpecControl{}), snapshot.Bridge.ControlFields)

	doc, err := os.ReadFile(filepath.Join("..", "..", "..", "vitepress", "reference", "userform-spec.md"))
	if err != nil {
		t.Fatal(err)
	}
	for _, expected := range snapshot.BuiltInControls {
		if !documentationHasControl(string(doc), expected.Type, expected.ProgID) {
			t.Fatalf("UserForm documentation missing %s / %s", expected.Type, expected.ProgID)
		}
	}
}

func documentationHasControl(doc, typeName, progID string) bool {
	for _, line := range strings.Split(doc, "\n") {
		if strings.Contains(line, "| `"+typeName+"`") && strings.Contains(line, "`"+progID+"`") {
			return true
		}
	}
	return false
}

func readFormsFixture(t *testing.T, name string) []byte {
	t.Helper()
	body, err := os.ReadFile(filepath.Join("intel", "testdata", name))
	if err != nil {
		t.Fatal(err)
	}
	return body
}

func assertJSONFields(t *testing.T, typ reflect.Type, wants []string) {
	t.Helper()
	seen := map[string]bool{}
	for index := 0; index < typ.NumField(); index++ {
		tag := strings.Split(typ.Field(index).Tag.Get("json"), ",")[0]
		if tag != "" && tag != "-" {
			seen[tag] = true
		}
	}
	for _, want := range wants {
		if !seen[want] {
			t.Fatalf("%s is missing JSON field %q", typ.Name(), want)
		}
	}
}
