package compiler

import (
	"encoding/json"
	"errors"
	"os"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/oforms"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func TestTypedAllStandardControlMatrix(t *testing.T) {
	body, err := os.ReadFile("testdata/typed-properties.json")
	if err != nil {
		t.Fatal(err)
	}
	var input spec.FormSpec
	if err = json.Unmarshal(body, &input); err != nil {
		t.Fatal(err)
	}
	form, err := CompileNew(input, 932)
	if err != nil {
		t.Fatal(err)
	}
	got, err := projection.Project(form)
	if err != nil {
		t.Fatal(err)
	}
	if len(got.Controls) != 15 {
		t.Fatalf("controls: %d", len(got.Controls))
	}
	for _, want := range input.Controls {
		var actual spec.FormSpecControl
		for _, control := range got.Controls {
			if control.Name == want.Name {
				actual = control
				break
			}
		}
		if actual.Name == "" {
			t.Fatalf("lost %s", want.Name)
		}
		if want.Type == "MultiPage" {
			model := findControl(form.Controls, want.Name)
			if model == nil || model.MultiPage == nil || model.MultiPage.Hidden.Record.Values["BackColor"] != 0x00ffeedd || model.MultiPage.Hidden.Record.Values["ForeColor"] != 0x00563412 {
				t.Fatal("MultiPage colors were not routed to hidden TabStrip")
			}
			if _, stored := model.Record.Values["BackColor"]; stored {
				t.Fatal("MultiPage edit incorrectly changed container color")
			}
		}
		if want.Font != nil && !fontMatches(persistenceFont(want.Font), persistenceFont(actual.Font)) {
			t.Fatalf("%s lost font", want.Type)
		}
		if !reflect.DeepEqual(want.ForeColor, actual.ForeColor) || !reflect.DeepEqual(want.BackColor, actual.BackColor) || !reflect.DeepEqual(want.TextAlign, actual.TextAlign) || !reflect.DeepEqual(want.BorderStyle, actual.BorderStyle) {
			t.Fatalf("%s color/enums differ: %+v", want.Type, actual)
		}
	}
}

func typedFixture() spec.FormSpec {
	return spec.FormSpec{SchemaVersion: 1, Kind: "xlflow.userform", Basis: "designer", CoordinateSystem: "points", Form: spec.FormSpecForm{Name: "Typed", Font: &spec.FormSpecFont{Name: new("Tahoma"), Size: new(9.5), Charset: new(128)}, BackColor: &spec.FormSpecColor{Kind: "system", Value: "buttonFace"}}, Controls: []spec.FormSpecControl{
		{ID: "label", Type: "Label", Name: "LabelOne", Font: &spec.FormSpecFont{Name: new("ＭＳ ゴシック"), Size: new(12.5), Bold: new(true), Italic: new(false), Charset: new(128)}, ForeColor: &spec.FormSpecColor{Kind: "rgb", Value: "#123456"}, TextAlign: new(2), BorderStyle: new(1)},
		{ID: "image", Type: "Image", Name: "ImageOne", BackColor: &spec.FormSpecColor{Kind: "ole", Value: "0x80000123"}, PictureAlignment: new(4), PictureSizeMode: new(3)},
	}}
}

func TestTypedGenerationAndTemplateOverlay(t *testing.T) {
	input := typedFixture()
	form, err := CompileNew(input, 932)
	if err != nil {
		t.Fatal(err)
	}
	before, err := projection.Project(form)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(before.Form.Font, input.Form.Font) || !reflect.DeepEqual(before.Form.BackColor, input.Form.BackColor) {
		t.Fatalf("root projection changed typed state: %+v", before.Form)
	}
	label, image := before.Controls[0], before.Controls[1]
	if *label.Font.Name != "ＭＳ ゴシック" || !*label.Font.Bold || *label.TextAlign != 2 || *label.BorderStyle != 1 || !reflect.DeepEqual(label.ForeColor, input.Controls[0].ForeColor) {
		t.Fatalf("label typed state differs: %+v", label)
	}
	if *image.PictureAlignment != 4 || *image.PictureSizeMode != 3 || !reflect.DeepEqual(image.BackColor, input.Controls[1].BackColor) {
		t.Fatalf("image typed state differs: %+v", image)
	}
	// The template contains only one font member. Unspecified members and the
	// unknown OLE token must survive canonical file-backed compilation.
	desired := spec.FormSpec{SchemaVersion: 1, Kind: "xlflow.userform", Basis: "designer", Form: spec.FormSpecForm{Name: "Typed", Font: &spec.FormSpecFont{Bold: new(true)}}, Controls: []spec.FormSpecControl{{ID: "label", Type: "Label", Name: "LabelOne", Font: &spec.FormSpecFont{Italic: new(true)}}, {ID: "image", Type: "Image", Name: "ImageOne"}}}
	result, err := CompileTemplate(form, desired, 932)
	if err != nil {
		t.Fatal(err)
	}
	after, err := projection.Project(result)
	if err != nil {
		t.Fatal(err)
	}
	if after.Form.Font.Name == nil || *after.Form.Font.Name != "Tahoma" || !*after.Form.Font.Bold || *after.Controls[0].Font.Name != "ＭＳ ゴシック" || !*after.Controls[0].Font.Italic || !*after.Controls[0].Font.Bold || !reflect.DeepEqual(after.Controls[1].BackColor, image.BackColor) {
		t.Fatal("template partial edit replaced unauthored state")
	}
	serialized, _ := oforms.SerializeForm(result, 932)
	noop, err := CompileTemplate(result, desired, 932)
	if err != nil {
		t.Fatal(err)
	}
	again, _ := oforms.SerializeForm(noop, 932)
	if !reflect.DeepEqual(serialized, again) {
		t.Fatal("repeated typed template changed bytes")
	}
}

func TestTypedEditsMigrationStaleAndMemberRemoval(t *testing.T) {
	form, err := CompileNew(typedFixture(), 932)
	if err != nil {
		t.Fatal(err)
	}
	before, err := projection.Project(form)
	if err != nil {
		t.Fatal(err)
	}
	legacy := snapshotCopy(t, before)
	legacy.Controls[0].ForeColor = nil
	legacy.Controls[0].Properties = map[string]any{"ForeColor": float64(0x563412)}
	after := snapshotCopy(t, legacy)
	after.Controls[0].ForeColor = &spec.FormSpecColor{Kind: "system", Value: "windowText"}
	delete(after.Controls[0].Properties, "ForeColor")
	result, err := CompileEdits(form, legacy, after, 932)
	if err != nil {
		t.Fatal(err)
	}
	if got := result.Controls[0].Record.Values["ForeColor"]; got != 0x80000008 {
		t.Fatalf("migration wrote %#x", got)
	}
	for _, test := range []struct {
		name string
		edit func(*spec.FormSpec, *spec.FormSpec)
		code string
	}{
		{"stale", func(b, a *spec.FormSpec) {
			b.Controls[0].Font.Name = new("Wrong")
			a.Controls[0].Font.Name = new("Arial")
		}, Stale},
		{"member-removal", func(b, a *spec.FormSpec) { a.Controls[0].Font.Size = nil }, Unsupported},
	} {
		t.Run(test.name, func(t *testing.T) {
			b, a := snapshotCopy(t, before), snapshotCopy(t, before)
			test.edit(&b, &a)
			_, err := CompileEdits(form, b, a, 932)
			detail, ok := errors.AsType[*Error](err)
			if !ok || detail.Code != test.code {
				t.Fatalf("expected %s, got %v", test.code, err)
			}
		})
	}
}

func TestTypedPropertiesOnNewTopologyControl(t *testing.T) {
	base := typedFixture()
	base.Controls = nil
	form, err := CompileNew(base, 932)
	if err != nil {
		t.Fatal(err)
	}
	before, err := projection.Project(form)
	if err != nil {
		t.Fatal(err)
	}
	after := snapshotCopy(t, before)
	after.Controls = typedFixture().Controls
	result, err := CompileEdits(form, before, after, 932)
	if err != nil {
		t.Fatal(err)
	}
	got, err := projection.Project(result)
	if err != nil {
		t.Fatal(err)
	}
	if got.Controls[0].Font == nil || *got.Controls[0].Font.Name != "ＭＳ ゴシック" || got.Controls[1].PictureSizeMode == nil || *got.Controls[1].PictureSizeMode != 3 {
		t.Fatal("new topology dropped typed authoring state")
	}
}

func TestLegacyTypedPropertiesRemainReadableWithoutSourceMigration(t *testing.T) {
	input := typedFixture()
	input.Controls[0].Font = nil
	input.Controls[0].ForeColor = nil
	input.Controls[0].TextAlign = nil
	input.Controls[0].Properties = map[string]any{"Font": map[string]any{"name": "Arial", "size": 11.5, "italic": true}, "ForeColor": float64(0x80000008), "TextAlign": 3}
	form, err := CompileNew(input, 932)
	if err != nil {
		t.Fatal(err)
	}
	got, err := projection.Project(form)
	if err != nil {
		t.Fatal(err)
	}
	if got.Controls[0].Font == nil || *got.Controls[0].Font.Name != "Arial" || *got.Controls[0].TextAlign != 3 || got.Controls[0].ForeColor.Value != "windowText" {
		t.Fatal("legacy font/color/alignment not read")
	}
	// A legacy-only request must replace the projected typed baseline, rather
	// than accidentally conflicting with the compiler's own before snapshot.
	input.Controls[0].Properties = map[string]any{"ForeColor": float64(0x8000000f), "Font": map[string]any{"bold": true}}
	compiled, err := CompileTemplate(form, input, 932)
	if err != nil {
		t.Fatal(err)
	}
	updated, err := projection.Project(compiled)
	if err != nil {
		t.Fatal(err)
	}
	if updated.Controls[0].ForeColor.Value != "buttonFace" || !*updated.Controls[0].Font.Bold || *updated.Controls[0].Font.Name != "Arial" || !*updated.Controls[0].Font.Italic {
		t.Fatal("legacy template overwrote partial font state")
	}
	if input.Controls[0].ForeColor != nil || input.Controls[0].Font != nil {
		t.Fatal("compiler migrated caller authoring source")
	}
}

func TestTypedPictureEnumEditHasNoProjectedLegacyConflict(t *testing.T) {
	form, err := CompileNew(typedFixture(), 932)
	if err != nil {
		t.Fatal(err)
	}
	before, err := projection.Project(form)
	if err != nil {
		t.Fatal(err)
	}
	after := snapshotCopy(t, before)
	after.Controls[1].PictureAlignment = new(0)
	after.Controls[1].PictureSizeMode = new(1)
	result, err := CompileEdits(form, before, after, 932)
	if err != nil {
		t.Fatal(err)
	}
	if result.Controls[1].Record.Values["PictureAlignment"] != 0 || result.Controls[1].Record.Values["PictureSizeMode"] != 1 {
		t.Fatal("image enums did not persist")
	}
}

func TestPagePersistedFontIsOpaqueWithoutRuntimeProperty(t *testing.T) {
	form, err := oforms.NewForm(oforms.Definition{Name: "OpaquePage", Size: oforms.Size{Width: 5000, Height: 4000}, Controls: []oforms.ControlDefinition{{Name: "Pages", Class: 57, Visible: true, Size: oforms.Size{Width: 4000, Height: 3000}, Tabs: &oforms.TabStrip{Tabs: []oforms.Tab{{Name: "TabOne", Caption: "PageOne", Enabled: true, Visible: true}}}, Controls: []oforms.ControlDefinition{{Name: "PageOne", Class: 7, Visible: true}}}}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	form, err = oforms.ApplyEdits(form, []oforms.Edit{{Control: "PageOne", Property: "Font", Value: oforms.Font{Name: new("Tahoma"), Size: new(12.5)}}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	projected, err := projection.Project(form)
	if err != nil {
		t.Fatal(err)
	}
	for _, control := range projected.Controls {
		if control.Type != "Page" {
			continue
		}
		if control.Font != nil || control.BackColor != nil || control.ForeColor != nil || control.BorderStyle != nil {
			t.Fatal("Page exposed inapplicable typed property")
		}
		found := false
		for _, name := range control.Unsupported {
			found = found || name == "font"
		}
		if !found {
			t.Fatal("opaque Page font was not described")
		}
	}
	before, _ := oforms.SerializeForm(form, 932)
	result, err := CompileEdits(form, projected, projected, 932)
	if err != nil {
		t.Fatal(err)
	}
	after, _ := oforms.SerializeForm(result, 932)
	if !reflect.DeepEqual(before, after) {
		t.Fatal("opaque Page font changed on no-op")
	}
}
