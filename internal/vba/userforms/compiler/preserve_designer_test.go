package compiler

import (
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func TestCompileTemplateCodeOnlyRejectsDesignerEditsAtomically(t *testing.T) {
	for _, scenario := range []string{"caption", "control", "remove"} {
		t.Run(scenario, func(t *testing.T) {
			base, snapshot := fixture(t, "p6_nested_form.bin")
			desired := templateInput(t, base, snapshot)
			before, err := oforms.SerializeForm(base, 932)
			if err != nil {
				t.Fatal(err)
			}
			switch scenario {
			case "caption":
				desired.Form.Caption = new("Changed root")
			case "control":
				desired.Controls[0].Caption = new("Changed control")
			case "remove":
				desired.Controls = []spec.FormSpecControl{}
			}
			_, err = CompileTemplateCodeOnly(base, desired, 932)
			assertTemplateErrorCode(t, err, Unsupported)
			after, err := oforms.SerializeForm(base, 932)
			if err != nil {
				t.Fatal(err)
			}
			assertDesignerBytes(t, before, after)
		})
	}
}

func TestCompileTemplateCodeOnlyPreservesNestedDesigner(t *testing.T) {
	base, snapshot := fixture(t, "p6_nested_form.bin")
	before, err := oforms.SerializeForm(base, 932)
	if err != nil {
		t.Fatal(err)
	}
	result, err := CompileTemplateCodeOnly(base, templateInput(t, base, snapshot), 932)
	if err != nil {
		t.Fatal(err)
	}
	after, err := oforms.SerializeForm(result, 932)
	if err != nil {
		t.Fatal(err)
	}
	assertDesignerBytes(t, before, after)
}
