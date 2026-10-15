package lspserver

import (
	"encoding/json"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	formsedit "github.com/harumiWeb/xlflow/internal/vba/userforms/spec/edit"
)

func TestUserFormPropertyEditorDecoding(t *testing.T) {
	for _, raw := range []string{
		`{"type":"setControlProperty","controlId":"label","field":"font","value":{"bold":false,"size":8.25,"charset":17}}`,
		`{"type":"setFormProperty","field":"backColor","value":{"kind":"system","value":"buttonFace"}}`,
		`{"type":"setControlProperty","controlId":"image","field":"picture","value":{"remove":true}}`,
		`{"type":"setGeometry","controlId":"label","left":0,"height":15.5}`,
		`{"type":"setGeometry","width":250}`,
	} {
		operation, err := parseUserFormEditOperation([]byte(raw))
		if err != nil {
			t.Fatalf("valid editor operation %s: %+v", raw, err)
		}
		if err := validateUserFormEditTransaction([]formsedit.Operation{operation}); err != nil {
			t.Fatal(err)
		}
		if operation.Field == "font" {
			value := operation.Value.(map[string]any)
			if value["size"] != json.Number("8.25") || value["bold"] != false || value["charset"] != json.Number("17") {
				t.Fatalf("nested value precision/types lost: %+v", value)
			}
		}
	}
	for _, raw := range []string{
		`{"type":"setControlProperty","controlId":"label","field":"properties","value":{"BackColor":0}}`,
		`{"type":"setFormProperty","field":"font","value":{"bold":false,"bold":true}}`,
		`{"type":"setFormProperty","field":"font","value":[]}`,
		`{"type":"setGeometry"}`,
		`{"type":"setGeometry","width":null}`,
		`{"type":"setGeometry","controlId":null,"width":10}`,
		`{"type":"setGeometry","controlId":"","width":10}`,
		`{"type":"setGeometry","left":0}`,
		`{"type":"setGeometry","width":0.10000000000000000001}`,
		`{"type":"setGeometry","controlId":"label","left":9007199254740993}`,
		`{"type":"setGeometry","controlId":"label","width":1,"value":0}`,
	} {
		if _, err := parseUserFormEditOperation([]byte(raw)); err == nil {
			t.Fatalf("invalid editor payload accepted: %s", raw)
		}
	}
}

func TestUserFormPropertyEditorPreviewOptIn(t *testing.T) {
	root := t.TempDir()
	server := &Server{opts: Options{RootDir: root, Config: config.Default()}}
	for _, extension := range []string{"yaml", "yml", "json"} {
		source := strings.Replace(previewSource, "  width: 240", "  font: {size: 8.25, bold: false}\n  width: 240", 1)
		if extension == "json" {
			source = `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main","font":{"size":8.25,"bold":false}},"controls":[]}`
		}
		params := userFormPreviewParams{URI: pathToFileURI(filepath.Join(root, "src/forms/specs/Main."+extension)), Version: 9, Text: source}
		legacy := server.userFormPreview(params)
		if legacy.Error != nil || legacy.PropertyError != nil || legacy.PropertyGrid == nil {
			t.Fatalf("legacy preview: %+v", legacy)
		}
		if _, exists := legacy.PropertyGrid.Form.Values["font"]; exists {
			t.Fatal("legacy client received typed metadata")
		}
		params.PropertyEditorVersion = 1
		current := server.userFormPreview(params)
		if current.Error != nil || current.PropertyError != nil || current.PropertyGrid == nil || current.Version != 9 {
			t.Fatalf("editor preview: %+v", current)
		}
		if !current.PropertyGrid.Form.Values["font"].Present {
			t.Fatal("font authored value missing")
		}
		found := false
		for _, descriptor := range current.PropertyGrid.Form.Descriptors {
			if descriptor.Field == "font" {
				found = descriptor.Editor == "font" && len(descriptor.Children) > 0
			}
		}
		if !found {
			t.Fatal("font editor descriptor missing")
		}
	}
}

func TestUserFormPropertyEditorSourceTransactions(t *testing.T) {
	root := t.TempDir()
	server := &Server{opts: Options{RootDir: root, Config: config.Default()}}
	source := strings.Replace(previewSource, "controls: []", "controls:\n  - id: label\n    name: Label1\n    type: Label\n    font: {size: 8.25, italic: true}\n    properties: {BackColor: -2147483633, Vendor: keep}", 1)
	for _, raw := range []string{
		`{"type":"setControlProperty","controlId":"label","field":"font","value":{"bold":false}}`,
		`{"type":"setControlProperty","controlId":"label","field":"backColor","value":{"kind":"rgb","value":"#010203"}}`,
		`{"type":"setGeometry","controlId":"label","left":0,"height":17.5}`,
		`{"type":"setGeometry","width":300}`,
	} {
		op, err := parseUserFormEditOperation([]byte(raw))
		if err != nil {
			t.Fatal(err)
		}
		result := server.userFormEdit(userFormEditParams{URI: pathToFileURI(filepath.Join(root, "src/forms/specs/Main.yaml")), Version: 5, Text: source, Operations: []formsedit.Operation{op}})
		if result.Error != nil || result.Version != 5 || len(result.Edits) == 0 {
			t.Fatalf("editor transaction failed %s: %+v", raw, result)
		}
	}
	pageSource := strings.Replace(source, "type: Label", "type: Page", 1)
	result := server.userFormEdit(userFormEditParams{URI: pathToFileURI(filepath.Join(root, "src/forms/specs/Main.yaml")), Text: pageSource, Operations: []formsedit.Operation{{Type: formsedit.SetGeometry, ControlID: "label", Width: new(20.0)}}})
	if result.Error == nil || len(result.Edits) != 0 {
		t.Fatal("Page geometry was accepted")
	}
	// Two editor confirmations must never become an accidental multi-edit undo.
	if err := validateUserFormEditTransaction([]formsedit.Operation{{Type: formsedit.SetGeometry, Width: new(100.0)}, {Type: formsedit.SetFormProperty, Field: "font", Value: &spec.FormSpecFont{Bold: new(true)}}}); err == nil {
		t.Fatal("unrelated editor transactions were merged")
	}
}
