package edit

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

const editorYAML = "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform: {name: Main}\ncontrols:\n  - id: label\n    name: Label1\n    type: Label\n    font:\n      name: 'Tahoma' # keep family comment\n      size: 8.25 # keep size comment\n      italic: false\n    left: 10 # keep geometry comment\n    width: 100\n  - {id: image, name: Image1, type: Image}\n"
const editorJSON = `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"label","name":"Label1","type":"Label","font":{"name":"Tahoma","size":8.25,"italic":false},"left":1e1,"width":100},{"id":"image","name":"Image1","type":"Image"}]}`

func TestPropertyEditorsPartialFontPreservesAuthoredSyntax(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		source := editorYAML
		if format == "json" {
			source = editorJSON
		}
		operation := Operation{Type: SetControlProperty, ControlID: "label", Field: "font", Value: map[string]any{"bold": true, "charset": json.Number("17")}}
		result := applyFormatOK(t, format, source, operation)
		font := result.Document.Controls[0].Font
		if font == nil || font.Name == nil || *font.Name != "Tahoma" || font.Size == nil || *font.Size != 8.25 || font.Italic == nil || *font.Italic || font.Bold == nil || !*font.Bold || font.Charset == nil || *font.Charset != 17 {
			t.Fatalf("partial font lost values: %+v", font)
		}
		if format == "yaml" && (!bytes.Contains(result.Source, []byte("name: 'Tahoma' # keep family comment")) || !bytes.Contains(result.Source, []byte("size: 8.25 # keep size comment"))) {
			t.Fatalf("font comments changed: %s", result.Source)
		}
		if format == "json" && !bytes.Contains(result.Source, []byte(`"left":1e1`)) {
			t.Fatalf("neighbor token changed: %s", result.Source)
		}
		second := applyFormatOK(t, format, string(result.Source), operation)
		if len(second.Edits) != 0 {
			t.Fatalf("confirmed unchanged font is not a no-op: %+v", second.Edits)
		}
	}
}

func TestPropertyEditorsMetadataNegotiationAndSuggestions(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		source := editorYAML
		if format == "json" {
			source = editorJSON
		}
		input := spec.SpecInput{Format: format}
		legacy, err := PropertyGrid(input, []byte(source))
		if err != nil {
			t.Fatal(err)
		}
		for _, version := range []int{0, 2} {
			grid, err := PropertyGrid(input, []byte(source), version)
			if err != nil {
				t.Fatal(err)
			}
			a, _ := json.Marshal(legacy)
			b, _ := json.Marshal(grid)
			if !bytes.Equal(a, b) {
				t.Fatalf("unsupported version changed legacy metadata: %s", b)
			}
		}
		if _, exists := legacy.Controls["label"].Values["font"]; exists {
			t.Fatal("typed value leaked to legacy client")
		}
		grid, err := PropertyGrid(input, []byte(source), 1)
		if err != nil {
			t.Fatal(err)
		}
		state := grid.Controls["label"].Values["font"]
		if !state.Present || state.Legacy {
			t.Fatalf("typed state: %+v", state)
		}
		if grid.Form.Values["font"].Present {
			t.Fatal("absent form font became authored")
		}
		for _, descriptor := range grid.Controls["label"].Descriptors {
			if descriptor.Field != "font" {
				continue
			}
			if descriptor.Editor != "font" {
				t.Fatalf("missing editor: %+v", descriptor)
			}
			for _, child := range descriptor.Children {
				if child.Field == "charset" && (len(child.AllowedValues) != 0 || len(child.Suggestions) == 0) {
					t.Fatalf("charset suggestions restrict arbitrary charset: %+v", child)
				}
				if child.Field == "charset" && (child.Minimum == nil || *child.Minimum != 0 || child.Maximum == nil || *child.Maximum != 255 || child.ExclusiveMinimum) {
					t.Fatalf("charset bounds missing: %+v", child)
				}
				if child.Field == "size" && (child.Minimum == nil || *child.Minimum != 0 || child.Maximum == nil || *child.Maximum != 3276.75 || !child.ExclusiveMinimum) {
					t.Fatalf("font size bounds missing: %+v", child)
				}
			}
		}
		if _, exists := grid.Controls["image"].Values["font"]; exists {
			t.Fatal("Image exposed Font")
		}
	}
}

func TestPropertyEditorsMigrateOnlyEditedLegacyKey(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		for _, bag := range []string{"single", "multiple", "flow"} {
			source := editorYAML
			properties := "    properties:\n      BackColor: -2147483633 # legacy color\n"
			if bag == "multiple" {
				properties += "      VendorState: keep # retain vendor comment\n"
			}
			if bag == "flow" {
				properties = "    properties: {BackColor: -2147483633, VendorState: keep}\n"
			}
			source = strings.Replace(source, "    left: 10", properties+"    left: 10", 1)
			if format == "json" {
				properties = `"properties":{"BackColor":-2147483633}`
				if bag != "single" {
					properties = `"properties":{"BackColor":-2147483633,"VendorState":"keep"}`
				}
				source = strings.Replace(editorJSON, `"left":1e1`, properties+`,"left":1e1`, 1)
			}
			grid, err := PropertyGrid(spec.SpecInput{Format: format}, []byte(source), 1)
			if err != nil {
				t.Fatal(err)
			}
			state := grid.Controls["label"].Values["backColor"]
			if !state.Present || !state.Legacy {
				t.Fatalf("legacy authoring state lost: %+v", state)
			}
			result := applyFormatOK(t, format, source, Operation{Type: SetControlProperty, ControlID: "label", Field: "backColor", Value: map[string]any{"kind": "rgb", "value": "#AABBCC"}})
			if bytes.Contains(result.Source, []byte("BackColor")) {
				t.Fatalf("legacy value not migrated: %s", result.Source)
			}
			if bag != "single" && !bytes.Contains(result.Source, []byte("VendorState")) {
				t.Fatalf("unrelated legacy property lost: %s", result.Source)
			}
			if result.Document.Controls[0].BackColor == nil || result.Document.Controls[0].BackColor.Value != "#AABBCC" {
				t.Fatalf("typed color lost: %s", result.Source)
			}
		}
	}
}

func TestPropertyEditorsRejectArbitraryObjectsAndPrecisionLoss(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		source := editorYAML
		if format == "json" {
			source = editorJSON
		}
		for _, op := range []Operation{
			{Type: SetControlProperty, ControlID: "label", Field: "properties", Value: map[string]any{"BackColor": 0}},
			{Type: SetControlProperty, ControlID: "label", Field: "caption", Value: map[string]any{"name": "unsafe"}},
			{Type: SetControlProperty, ControlID: "label", Field: "font", Value: map[string]any{"size": json.Number("0.1000000000000000001")}},
			{Type: SetControlProperty, ControlID: "label", Field: "font", Value: map[string]any{"bold": "false"}},
			{Type: SetControlProperty, ControlID: "label", Field: "font", Value: map[string]any{"charset": 256}},
			{Type: SetControlProperty, ControlID: "label", Field: "font", Value: map[string]any{"nested": map[string]any{}}},
			{Type: SetControlProperty, ControlID: "image", Field: "picture", ValuePresent: true},
		} {
			result, err := Apply(spec.SpecInput{Format: format}, []byte(source), []Operation{op})
			if err == nil || len(result.Source) > 0 || len(result.Edits) > 0 {
				t.Fatalf("unsafe payload accepted: %+v %v", op, err)
			}
		}
	}
}

func TestPropertyEditorsGeometryPartialAxesAndForm(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		source := editorYAML
		if format == "json" {
			source = editorJSON
		}
		result := applyFormatOK(t, format, source, Operation{Type: SetGeometry, ControlID: "label", Left: new(0.0), Height: new(17.5)})
		control := result.Document.Controls[0]
		if control.Left == nil || *control.Left != 0 || control.Height == nil || *control.Height != 17.5 || control.Width == nil || *control.Width != 100 {
			t.Fatalf("geometry lost: %+v", control)
		}
		if format == "yaml" && !bytes.Contains(result.Source, []byte("# keep geometry comment")) {
			t.Fatalf("geometry comment lost: %s", result.Source)
		}
		result = applyFormatOK(t, format, source, Operation{Type: SetGeometry, Width: new(200.0)})
		if result.Document.Form.Width == nil || *result.Document.Form.Width != 200 || result.Document.Form.Height != nil {
			t.Fatal("form unchanged axis was authored")
		}
		for _, op := range []Operation{{Type: SetGeometry}, {Type: SetGeometry, Left: new(1.0)}, {Type: SetGeometry, ControlID: "missing", Height: new(1.0)}} {
			if _, err := Apply(spec.SpecInput{Format: format}, []byte(source), []Operation{op}); err == nil {
				t.Fatalf("invalid geometry accepted: %+v", op)
			}
		}
	}
}

func TestPropertyEditorsNullAndPictureActions(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		source := editorYAML
		if format == "json" {
			source = editorJSON
		}
		null := applyFormatOK(t, format, source, Operation{Type: SetControlProperty, ControlID: "label", Field: "font", ValuePresent: true})
		grid, err := PropertyGrid(spec.SpecInput{Format: format}, null.Source, 1)
		if err != nil {
			t.Fatal(err)
		}
		if state := grid.Controls["label"].Values["font"]; !state.Present || state.Value != nil {
			t.Fatalf("font explicit null lost: %+v", state)
		}
		updated := applyFormatOK(t, format, string(null.Source), Operation{Type: SetControlProperty, ControlID: "label", Field: "font", Value: map[string]any{"bold": false}})
		if font := updated.Document.Controls[0].Font; font == nil || font.Bold == nil || *font.Bold {
			t.Fatalf("font false after null lost: %+v", font)
		}
		image := applyFormatOK(t, format, source, Operation{Type: SetControlProperty, ControlID: "image", Field: "picture", Value: map[string]any{"path": "src/forms/assets/picture.bmp"}})
		removed := applyFormatOK(t, format, string(image.Source), Operation{Type: SetControlProperty, ControlID: "image", Field: "picture", Value: map[string]any{"remove": true}})
		if picture := removed.Document.Controls[1].Picture; picture == nil || !picture.Remove || picture.Path != "" {
			t.Fatalf("picture remove retained old path: %+v", picture)
		}
		readded := applyFormatOK(t, format, string(removed.Source), Operation{Type: SetControlProperty, ControlID: "image", Field: "picture", Value: map[string]any{"path": "src/forms/assets/other.jpg"}})
		if picture := readded.Document.Controls[1].Picture; picture == nil || picture.Remove || picture.Path != "src/forms/assets/other.jpg" {
			t.Fatalf("picture update retained remove flag: %+v", picture)
		}
	}
}

func TestPropertyEditorsLegacyFontPartialMigration(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		source := strings.Replace(editorYAML, "    font:\n      name: 'Tahoma' # keep family comment\n      size: 8.25 # keep size comment\n      italic: false", "    properties: {Font: {name: Tahoma, size: 8.25, italic: false}, Vendor: keep}", 1)
		if format == "json" {
			source = strings.Replace(editorJSON, `"font":{"name":"Tahoma","size":8.25,"italic":false}`, `"properties":{"Font":{"name":"Tahoma","size":8.25,"italic":false},"Vendor":"keep"}`, 1)
		}
		result := applyFormatOK(t, format, source, Operation{Type: SetControlProperty, ControlID: "label", Field: "font", Value: map[string]any{"bold": true}})
		font := result.Document.Controls[0].Font
		if font == nil || font.Name == nil || *font.Name != "Tahoma" || font.Italic == nil || *font.Italic || font.Size == nil || *font.Size != 8.25 || font.Bold == nil || !*font.Bold {
			t.Fatalf("legacy font was not merged: %+v", font)
		}
		if _, exists := result.Document.Controls[0].Properties["Font"]; exists {
			t.Fatal("legacy Font not removed")
		}
		if result.Document.Controls[0].Properties["Vendor"] != "keep" {
			t.Fatal("unrelated property lost")
		}
	}
}

func TestPropertyEditorsTypedLegacyConflictsAndUnknownOLE(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		source := strings.Replace(editorYAML, "    left: 10", "    backColor: {kind: ole, value: '0x81234567'}\n    properties: {BackColor: 2166572391, Vendor: keep}\n    left: 10", 1)
		if format == "json" {
			source = strings.Replace(editorJSON, `"left":1e1`, `"backColor":{"kind":"ole","value":"0x81234567"},"properties":{"BackColor":2166572391,"Vendor":"keep"},"left":1e1`, 1)
		}
		result := applyFormatOK(t, format, source, Operation{Type: SetControlProperty, ControlID: "label", Field: "backColor", Value: map[string]any{"kind": "ole", "value": "0x81234567"}})
		if result.Document.Controls[0].BackColor.Kind != "ole" || result.Document.Controls[0].BackColor.Value != "0x81234567" || bytes.Contains(result.Source, []byte("BackColor")) {
			t.Fatalf("unknown OLE was not preserved/migrated: %s", result.Source)
		}
		conflict := strings.Replace(source, "2166572391", "0", 1)
		if _, err := Apply(spec.SpecInput{Format: format}, []byte(conflict), []Operation{{Type: SetControlProperty, ControlID: "label", Field: "backColor", Value: map[string]any{"kind": "rgb", "value": "#000000"}}}); err == nil {
			t.Fatal("typed/legacy conflict was silently overwritten")
		}
	}
}

func TestPropertyEditorsExcludeCustomActiveX(t *testing.T) {
	for _, format := range []string{"yaml", "json"} {
		source := strings.Replace(editorYAML, "    font:\n      name: 'Tahoma' # keep family comment\n      size: 8.25 # keep size comment\n      italic: false", "    progId: Vendor.Label.1", 1)
		if format == "json" {
			source = strings.Replace(editorJSON, `"font":{"name":"Tahoma","size":8.25,"italic":false}`, `"progId":"Vendor.Label.1"`, 1)
		}
		grid, err := PropertyGrid(spec.SpecInput{Format: format}, []byte(source), 1)
		if err != nil {
			t.Fatal(err)
		}
		for _, field := range []string{"font", "backColor", "foreColor", "textAlign", "borderStyle"} {
			if _, exists := grid.Controls["label"].Values[field]; exists {
				t.Fatalf("custom ActiveX editor exposed: %s", field)
			}
		}
		if _, err := Apply(spec.SpecInput{Format: format}, []byte(source), []Operation{{Type: SetControlProperty, ControlID: "label", Field: "font", Value: map[string]any{"bold": true}}}); err == nil {
			t.Fatal("custom ActiveX typed edit accepted")
		}
	}
}
