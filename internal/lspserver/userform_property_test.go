package lspserver

import (
	"encoding/json"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/compiler"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
)

func TestUserFormPropertyEditsAndAuthoringStates(t *testing.T) {
	s := &Server{opts: Options{RootDir: t.TempDir(), Config: config.Default()}}
	for _, format := range []string{"yaml", "json"} {
		t.Run(format, func(t *testing.T) {
			source := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform: {name: Main}\ncontrols:\n  - {id: label, type: Label, name: Label1, caption: '', enabled: false, left: 0}\n  - {id: text, type: TextBox, name: Text1, text: null}\n"
			if format == "json" {
				source = `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"label","type":"Label","name":"Label1","caption":"","enabled":false,"left":0},{"id":"text","type":"TextBox","name":"Text1","text":null}]}`
			}
			uri := pathToFileURI(filepath.Join(s.opts.RootDir, "src/forms/specs/Main."+format))
			preview := s.userFormPreview(userFormPreviewParams{URI: uri, Version: 8, Text: source})
			if preview.Error != nil || preview.PropertyError != nil || preview.PropertyGrid == nil {
				t.Fatalf("property preview: %+v", preview)
			}
			label := preview.PropertyGrid.Controls["label"]
			if !label.Values["caption"].Present || label.Values["caption"].Value != "" || !label.Values["enabled"].Present || label.Values["enabled"].Value != false || label.Values["width"].Present {
				t.Fatalf("source states: %+v", label.Values)
			}
			if !preview.PropertyGrid.Controls["text"].Values["text"].Present || preview.PropertyGrid.Controls["text"].Values["text"].Value != nil {
				t.Fatal("explicit null was lost")
			}
			for _, operation := range []map[string]any{
				{"type": "setFormProperty", "field": "build.clientWidth", "value": 200.5},
				{"type": "setControlProperty", "controlId": "label", "field": "caption", "value": "日本語 😀"},
				{"type": "setControlProperty", "controlId": "label", "field": "enabled", "value": true},
				{"type": "setControlProperty", "controlId": "text", "field": "text", "value": nil},
			} {
				body, err := json.Marshal(map[string]any{"uri": uri, "version": 8, "text": source, "operations": []any{operation}})
				if err != nil {
					t.Fatal(err)
				}
				params, requestError := parseUserFormEditParams(body)
				if requestError != nil {
					t.Fatal(requestError)
				}
				result := s.userFormEdit(params)
				if result.Error != nil || result.Version != 8 {
					t.Fatalf("property edit: %+v", result)
				}
				updated, err := applyLSPTextEdits(source, result.Edits)
				if err != nil {
					t.Fatal(err)
				}
				if _, err := spec.ParseFormSpec(spec.SpecInput{Format: format}, []byte(updated)); err != nil {
					t.Fatal(err)
				}
				if strings.Contains(updated, "observed") {
					t.Fatal("source edit materialized observed state")
				}
			}
		})
	}
}

func TestUserFormPropertyRPCNumericPrecision(t *testing.T) {
	s := &Server{opts: Options{RootDir: t.TempDir(), Config: config.Default()}}
	for _, format := range []string{"yaml", "json"} {
		source := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform: {name: Main}\ncontrols: [{id: text, type: TextBox, name: Text1, value: 0}]\n"
		if format == "json" {
			source = `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"text","type":"TextBox","name":"Text1","value":0}]}`
		}
		uri := pathToFileURI(filepath.Join(s.opts.RootDir, "src/forms/specs/Main."+format))
		for _, number := range []string{"0.1000000000000000001", "18446744073709551616", "1e-400", "0.1", "1e2", "9007199254740993"} {
			body, err := json.Marshal(map[string]any{"uri": uri, "version": 1, "text": source, "operations": []any{map[string]any{"type": "setControlProperty", "controlId": "text", "field": "value", "value": json.RawMessage(number)}}})
			if err != nil {
				t.Fatal(err)
			}
			params, requestError := parseUserFormEditParams(body)
			if requestError != nil {
				t.Fatalf("raw numeric request decode: %v", requestError)
			}
			result := s.userFormEdit(params)
			unsafe := number == "0.1000000000000000001" || number == "18446744073709551616" || number == "1e-400"
			if unsafe {
				if result.Error == nil || len(result.Edits) != 0 || len(result.Error.Diagnostics) == 0 || result.Error.Diagnostics[0].Code != "UFE003" || result.Error.Diagnostics[0].Field != "value" || result.Error.Diagnostics[0].ControlID != "text" {
					t.Fatalf("rounded RPC %s/%s: %+v", format, number, result)
				}
			} else if result.Error != nil || len(result.Edits) == 0 {
				t.Fatalf("exact RPC rejected %s/%s: %+v", format, number, result)
			}
		}
	}
}

func TestUserFormCustomControlCommonPropertyRPC(t *testing.T) {
	s := &Server{opts: Options{RootDir: t.TempDir(), Config: config.Default()}}
	for _, format := range []string{"yaml", "json"} {
		source := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform: {name: Main}\ncontrols: [{id: widget, type: VendorWidget, progId: Vendor.Widget.1, name: Widget1}]\n"
		if format == "json" {
			source = `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"widget","type":"VendorWidget","progId":"Vendor.Widget.1","name":"Widget1"}]}`
		}
		uri := pathToFileURI(filepath.Join(s.opts.RootDir, "src/forms/specs/Main."+format))
		preview := s.userFormPreview(userFormPreviewParams{URI: uri, Version: 1, Text: source})
		if preview.Error != nil || preview.PropertyError != nil || preview.PropertyGrid == nil || len(preview.PropertyGrid.Controls["widget"].Descriptors) == 0 {
			t.Fatalf("custom metadata: %+v", preview)
		}
		body, err := json.Marshal(map[string]any{"uri": uri, "version": 1, "text": source, "operations": []any{map[string]any{"type": "setControlProperty", "controlId": "widget", "field": "enabled", "value": false}}})
		if err != nil {
			t.Fatal(err)
		}
		params, requestError := parseUserFormEditParams(body)
		if requestError != nil {
			t.Fatal(requestError)
		}
		result := s.userFormEdit(params)
		if result.Error != nil || len(result.Edits) == 0 {
			t.Fatalf("custom RPC: %+v", result)
		}
		updated, err := applyLSPTextEdits(source, result.Edits)
		if err != nil {
			t.Fatal(err)
		}
		doc, err := spec.ParseFormSpec(spec.SpecInput{Format: format}, []byte(updated))
		if err != nil || doc.Controls[0].Enabled == nil || *doc.Controls[0].Enabled || doc.Controls[0].ProgID != "Vendor.Widget.1" {
			t.Fatalf("custom RPC result: %+v %v", doc, err)
		}
	}
}

func TestUserFormPropertyPayloadValidation(t *testing.T) {
	valid := `{"uri":"file:///x/src/forms/specs/Main.yaml","version":1,"text":"{}","operations":[{"type":"setControlProperty","controlId":"label","field":"caption","value":"hello"}]}`
	for _, body := range []string{
		strings.Replace(valid, `,"value":"hello"`, "", 1),
		strings.Replace(valid, `"value":"hello"`, `"value":{}`, 1),
		strings.Replace(valid, `"value":"hello"`, `"value":[]`, 1),
		strings.Replace(valid, `"value":"hello"`, `"value":1e400`, 1),
		strings.Replace(valid, `"field":"caption"`, `"field":null`, 1),
		strings.Replace(valid, `"field":"caption"`, `"field":"caption","field":"text"`, 1),
		strings.Replace(valid, `"controlId":"label"`, `"controlId":""`, 1),
		strings.Replace(valid, `"value":"hello"`, `"value":"hello","width":1`, 1),
		strings.Replace(valid, `}]}`, `},{"type":"moveControl","controlId":"label","left":1,"top":2}]}`, 1),
	} {
		if _, err := parseUserFormEditParams([]byte(body)); err == nil {
			t.Fatalf("accepted invalid payload %s", body)
		}
	}
	for _, scalar := range []string{`null`, `""`, `0`, `false`, `"false"`} {
		body := strings.Replace(valid, `"hello"`, scalar, 1)
		if _, err := parseUserFormEditParams([]byte(body)); err != nil {
			t.Fatalf("rejected scalar %s: %+v", scalar, err)
		}
	}
}

func TestUserFormPropertyValidationDiagnostics(t *testing.T) {
	s := &Server{opts: Options{RootDir: t.TempDir(), Config: config.Default()}}
	uri := pathToFileURI(filepath.Join(s.opts.RootDir, "src/forms/specs/Main.yaml"))
	source := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform: {name: Main}\ncontrols: [{id: label, type: Label, name: Label1}]\n"
	for _, op := range []string{
		`{"type":"setControlProperty","controlId":"label","field":"enabled","value":"false"}`,
		`{"type":"setControlProperty","controlId":"label","field":"tabIndex","value":1.5}`,
		`{"type":"setControlProperty","controlId":"label","field":"text","value":"unsupported"}`,
		`{"type":"setControlProperty","controlId":"label","field":"id","value":"new"}`,
		`{"type":"setFormProperty","field":"observed.width","value":1}`,
	} {
		body, _ := json.Marshal(map[string]any{"uri": uri, "version": 1, "text": source, "operations": []json.RawMessage{json.RawMessage(op)}})
		params, err := parseUserFormEditParams(body)
		if err != nil {
			t.Fatal(err)
		}
		result := s.userFormEdit(params)
		if result.Error == nil || len(result.Edits) != 0 || len(result.Error.Diagnostics) == 0 {
			t.Fatalf("missing atomic diagnostic for %s: %+v", op, result)
		}
	}
}

func TestUserFormPropertyEditsCompileAndProject(t *testing.T) {
	s := &Server{opts: Options{RootDir: t.TempDir(), Config: config.Default()}}
	for _, format := range []string{"yaml", "json"} {
		t.Run(format, func(t *testing.T) {
			source := "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform: {name: Main}\ncontrols: [{id: text, type: TextBox, name: Text1, text: Old, enabled: true}]\n"
			if format == "json" {
				source = `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"text","type":"TextBox","name":"Text1","text":"Old","enabled":true}]}`
			}
			uri := pathToFileURI(filepath.Join(s.opts.RootDir, "src/forms/specs/Main."+format))
			for index, op := range []map[string]any{
				{"type": "setFormProperty", "field": "build.caption", "value": "タイトル"},
				{"type": "setControlProperty", "controlId": "text", "field": "text", "value": "入力テキスト"},
				{"type": "setControlProperty", "controlId": "text", "field": "enabled", "value": false},
			} {
				body, err := json.Marshal(map[string]any{"uri": uri, "version": index + 1, "text": source, "operations": []any{op}})
				if err != nil {
					t.Fatal(err)
				}
				params, requestError := parseUserFormEditParams(body)
				if requestError != nil {
					t.Fatal(requestError)
				}
				result := s.userFormEdit(params)
				if result.Error != nil {
					t.Fatalf("edit %d: %+v", index, result.Error)
				}
				source, err = applyLSPTextEdits(source, result.Edits)
				if err != nil {
					t.Fatal(err)
				}
			}
			doc, err := spec.ParseFormSpec(spec.SpecInput{Format: format}, []byte(source))
			if err != nil {
				t.Fatal(err)
			}
			generated, err := compiler.CompileNew(doc, 932)
			if err != nil {
				t.Fatal(err)
			}
			projected, err := projection.Project(generated)
			if err != nil {
				t.Fatal(err)
			}
			if projected.Form.Build == nil || projected.Form.Build.Caption == nil || *projected.Form.Build.Caption != "タイトル" || len(projected.Controls) != 1 || projected.Controls[0].Text == nil || *projected.Controls[0].Text != "入力テキスト" || projected.Controls[0].Enabled == nil || *projected.Controls[0].Enabled {
				t.Fatalf("property compiler projection: %+v", projected)
			}
		})
	}
}
