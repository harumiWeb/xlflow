package lspserver

import (
	"cmp"
	"context"
	"encoding/json"
	"fmt"
	"math"
	"net"
	"os"
	"path/filepath"
	"slices"
	"strings"
	"testing"
	"unicode/utf16"
	"unicode/utf8"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/compiler"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/projection"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	formsedit "github.com/harumiWeb/xlflow/internal/vba/userforms/spec/edit"
	"github.com/sourcegraph/jsonrpc2"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

func TestUserFormEditRPCAndUTF16Ranges(t *testing.T) {
	ctx, cancel := context.WithTimeout(t.Context(), jsonrpcIntegrationTimeout)
	defer cancel()
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	serverSide, clientSide := net.Pipe()
	serverConn := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(serverSide, jsonrpc2.VSCodeObjectCodec{}), rpcHandler{handler: &s.handler, server: s, dispatch: s.dispatchRequest})
	defer func() { _ = serverConn.Close() }()
	clientConn := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(clientSide, jsonrpc2.VSCodeObjectCodec{}), &rpcRecorder{})
	defer func() { _ = clientConn.Close() }()
	var initialized struct {
		Capabilities struct {
			Experimental map[string]any `json:"experimental"`
		} `json:"capabilities"`
	}
	if err := clientConn.Call(ctx, "initialize", protocol.InitializeParams{}, &initialized); err != nil {
		t.Fatal(err)
	}
	if initialized.Capabilities.Experimental["userFormEdit"] != true || initialized.Capabilities.Experimental["userFormPropertyEdit"] != true || initialized.Capabilities.Experimental["userFormStructuralEdit"] != true || initialized.Capabilities.Experimental["userFormParentEdit"] != true {
		t.Fatalf("edit capability: %+v", initialized.Capabilities.Experimental)
	}
	if err := clientConn.Notify(ctx, "initialized", map[string]any{}); err != nil {
		t.Fatal(err)
	}

	uri := pathToFileURI(filepath.Join(root, "src/forms/specs/Main.json"))
	source := "{\r\n\"schemaVersion\":1,\r\n\"kind\":\"xlflow.userform\",\r\n\"basis\":\"designer\",\r\n\"form\":{\"name\":\"\u30e1\u30a4\u30f3\U0001F600\",\"width\":240,\"height\":180},\r\n\"controls\":[{\"id\":\"submit\",\"type\":\"CommandButton\",\"name\":\"\u9001\u4fe1\U0001F600\",\"left\":1,\"top\":2,\"width\":10,\"height\":8}]}"
	diskPath := filepath.Join(root, "src/forms/specs/Main.json")
	if err := os.MkdirAll(filepath.Dir(diskPath), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(diskPath, []byte("disk source is not the request buffer"), 0o644); err != nil {
		t.Fatal(err)
	}
	var result userFormEditResult
	params := map[string]any{
		"uri": uri, "version": 17, "text": source,
		"operations": []any{map[string]any{"type": "resizeControl", "controlId": "submit", "width": 20, "height": 30}},
	}
	if err := clientConn.Call(ctx, userFormEditMethod, params, &result); err != nil {
		t.Fatal(err)
	}
	if result.Version != 17 || result.Error != nil || len(result.Edits) != 2 {
		t.Fatalf("edit response: %+v", result)
	}
	controlsOffset := strings.Index(source, `"controls":[`)
	if controlsOffset < 0 {
		t.Fatal("fixture has no controls array")
	}
	controlWidthOffset := strings.Index(source[controlsOffset:], `"width":10`)
	if controlWidthOffset < 0 {
		t.Fatal("fixture has no target control width")
	}
	widthOffset := controlsOffset + controlWidthOffset + len(`"width":`)
	lineStart := strings.LastIndex(source[:widthOffset], "\n") + 1
	wantCharacter := len(utf16.Encode([]rune(source[lineStart:widthOffset])))
	if result.Edits[0].Range.Start.Line != 5 || int(result.Edits[0].Range.Start.Character) != wantCharacter || result.Edits[0].NewText != "20" {
		t.Fatalf("width edit position: %+v, want line 5 character %d", result.Edits[0], wantCharacter)
	}
	if result.Edits[1].Range.Start.Line != 5 || result.Edits[1].NewText != "30" {
		t.Fatalf("height edit position: %+v", result.Edits[1])
	}
	move := map[string]any{"type": "moveControl", "controlId": "submit", "left": 4, "top": 5}
	resize := map[string]any{"type": "resizeControl", "controlId": "submit", "width": 20, "height": 30}
	other := map[string]any{"type": "resizeControl", "controlId": "other", "width": 20, "height": 30}
	for _, test := range []struct {
		name       string
		operations []any
		valid      bool
	}{
		{"empty", []any{}, false},
		{"three operations", []any{move, resize, move}, false},
		{"repeated move", []any{move, move}, false},
		{"repeated resize", []any{resize, resize}, false},
		{"mixed targets", []any{move, other}, false},
		{"single move", []any{move}, true},
		{"move resize pair", []any{move, resize}, true},
		{"resize move pair", []any{resize, move}, true},
	} {
		t.Run(test.name, func(t *testing.T) {
			params["operations"] = test.operations
			var batch userFormEditResult
			if err := clientConn.Call(ctx, userFormEditMethod, params, &batch); err != nil {
				t.Fatal(err)
			}
			if batch.Version != 17 {
				t.Fatalf("version not preserved: %+v", batch)
			}
			if test.valid {
				if batch.Error != nil || len(batch.Edits) == 0 {
					t.Fatalf("valid transaction rejected: %+v", batch)
				}
			} else if batch.Error == nil || batch.Error.Code != compiler.Invalid || len(batch.Edits) != 0 {
				t.Fatalf("invalid transaction produced edits: %+v", batch)
			}
		})
	}
	params["operations"] = []any{map[string]any{"type": "setControlProperty", "controlId": "submit", "field": "caption", "value": "送信 😀"}}
	var property userFormEditResult
	if err := clientConn.Call(ctx, userFormEditMethod, params, &property); err != nil {
		t.Fatal(err)
	}
	if property.Error != nil || property.Version != 17 || len(property.Edits) == 0 {
		t.Fatalf("property RPC response: %+v", property)
	}
	edited, err := applyLSPTextEdits(source, property.Edits)
	if err != nil {
		t.Fatal(err)
	}
	doc, err := spec.ParseFormSpec(spec.SpecInput{Format: "json"}, []byte(edited))
	if err != nil || doc.Controls[0].Caption == nil || *doc.Controls[0].Caption != "送信 😀" {
		t.Fatalf("property RPC source: %s: %v", edited, err)
	}
	controlPayload := map[string]any{
		"id": "status", "name": "StatusLabel", "type": "Label",
		"left": 12, "top": 18, "width": 90, "height": 18,
	}
	params["operations"] = []any{map[string]any{"type": "addControl", "control": controlPayload}}
	var structural userFormEditResult
	if err := clientConn.Call(ctx, userFormEditMethod, params, &structural); err != nil {
		t.Fatal(err)
	}
	if structural.Error != nil || structural.Version != 17 || len(structural.Edits) == 0 {
		t.Fatalf("structural RPC response: %+v", structural)
	}
	structuralSource, err := applyLSPTextEdits(source, structural.Edits)
	if err != nil {
		t.Fatal(err)
	}
	structuralDoc, err := spec.ParseFormSpec(spec.SpecInput{Format: "json"}, []byte(structuralSource))
	if err != nil || len(structuralDoc.Controls) != 2 || structuralDoc.Controls[1].ID != "status" {
		t.Fatalf("structural RPC source: %s: %v", structuralSource, err)
	}
	structuralOperation := formsedit.Operation{
		Type: formsedit.AddControl,
		Control: &spec.FormSpecControl{
			ID: "status", Name: "StatusLabel", Type: "Label",
			Left: new(12.0), Top: new(18.0), Width: new(90.0), Height: new(18.0),
		},
	}
	wantStructural, err := formsedit.Apply(spec.SpecInput{Format: "json"}, []byte(source), []formsedit.Operation{structuralOperation})
	if err != nil || len(wantStructural.Edits) != len(structural.Edits) {
		t.Fatalf("expected structural source edits: %+v, %v", wantStructural.Edits, err)
	}
	for i, edit := range structural.Edits {
		start, err := byteOffsetAtLSPPosition(source, edit.Range.Start)
		if err != nil || start != wantStructural.Edits[i].Start {
			t.Fatalf("structural edit %d start maps to %d, want %d: %v", i, start, wantStructural.Edits[i].Start, err)
		}
		end, err := byteOffsetAtLSPPosition(source, edit.Range.End)
		if err != nil || end != wantStructural.Edits[i].End {
			t.Fatalf("structural edit %d end maps to %d, want %d: %v", i, end, wantStructural.Edits[i].End, err)
		}
	}
	diskSource, err := os.ReadFile(diskPath)
	if err != nil || string(diskSource) != "disk source is not the request buffer" {
		t.Fatalf("server changed disk source: %q, %v", diskSource, err)
	}
}

func TestUserFormEditRPCGeometryCompilesAndProjects(t *testing.T) {
	ctx, cancel := context.WithTimeout(t.Context(), jsonrpcIntegrationTimeout)
	defer cancel()
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	serverSide, clientSide := net.Pipe()
	serverConn := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(serverSide, jsonrpc2.VSCodeObjectCodec{}), rpcHandler{handler: &s.handler, server: s, dispatch: s.dispatchRequest})
	defer func() { _ = serverConn.Close() }()
	clientConn := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(clientSide, jsonrpc2.VSCodeObjectCodec{}), &rpcRecorder{})
	defer func() { _ = clientConn.Close() }()
	if err := clientConn.Call(ctx, "initialize", protocol.InitializeParams{}, new(struct{})); err != nil {
		t.Fatal(err)
	}
	if err := clientConn.Notify(ctx, "initialized", map[string]any{}); err != nil {
		t.Fatal(err)
	}

	uri := pathToFileURI(filepath.Join(root, "src/forms/specs/Main.json"))
	source := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"submit","type":"CommandButton","name":"Submit","left":1,"top":2,"width":10,"height":8}]}`
	params := map[string]any{
		"uri": uri, "version": 1, "text": source,
		"operations": []any{
			map[string]any{"type": "moveControl", "controlId": "submit", "left": 12.5, "top": 23.25},
			map[string]any{"type": "resizeControl", "controlId": "submit", "width": 45.75, "height": 16.5},
		},
	}
	var result userFormEditResult
	if err := clientConn.Call(ctx, userFormEditMethod, params, &result); err != nil {
		t.Fatal(err)
	}
	if result.Error != nil || len(result.Edits) != 4 {
		t.Fatalf("edit response: %+v", result)
	}
	editedSource, err := applyLSPTextEdits(source, result.Edits)
	if err != nil {
		t.Fatal(err)
	}
	doc, err := spec.ParseFormSpec(spec.SpecInput{Format: "json"}, []byte(editedSource))
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
	if len(projected.Controls) != 1 {
		t.Fatalf("projected controls: %+v", projected.Controls)
	}
	control := projected.Controls[0]
	for _, check := range []struct {
		name  string
		value *float64
		want  float64
	}{
		{"left", control.Left, 12.5}, {"top", control.Top, 23.25},
		{"width", control.Width, 45.75}, {"height", control.Height, 16.5},
	} {
		if check.value == nil || math.Abs(*check.value-check.want) > 0.02 {
			t.Errorf("projected %s = %v, want approximately %v", check.name, check.value, check.want)
		}
	}
}

func applyLSPTextEdits(source string, edits []protocol.TextEdit) (string, error) {
	ordered := slices.Clone(edits)
	slices.SortFunc(ordered, func(a, b protocol.TextEdit) int {
		if a.Range.Start.Line != b.Range.Start.Line {
			return cmp.Compare(b.Range.Start.Line, a.Range.Start.Line)
		}
		return cmp.Compare(b.Range.Start.Character, a.Range.Start.Character)
	})
	for _, edit := range ordered {
		start, err := byteOffsetAtLSPPosition(source, edit.Range.Start)
		if err != nil {
			return "", err
		}
		end, err := byteOffsetAtLSPPosition(source, edit.Range.End)
		if err != nil {
			return "", err
		}
		if start > end {
			return "", fmt.Errorf("edit start %d exceeds end %d", start, end)
		}
		source = source[:start] + edit.NewText + source[end:]
	}
	return source, nil
}

func byteOffsetAtLSPPosition(source string, position protocol.Position) (int, error) {
	lineStart := 0
	for line := uint32(0); line < position.Line; line++ {
		nextLine := strings.IndexByte(source[lineStart:], '\n')
		if nextLine < 0 {
			return 0, fmt.Errorf("line %d is outside source", position.Line)
		}
		lineStart += nextLine + 1
	}
	lineEnd := len(source)
	if nextLine := strings.IndexAny(source[lineStart:], "\r\n"); nextLine >= 0 {
		lineEnd = lineStart + nextLine
	}
	wanted := int(position.Character)
	if wanted == 0 {
		return lineStart, nil
	}
	units := 0
	for byteIndex, r := range source[lineStart:lineEnd] {
		units += len(utf16.Encode([]rune{r}))
		if units == wanted {
			return lineStart + byteIndex + utf8.RuneLen(r), nil
		}
		if units > wanted {
			return 0, fmt.Errorf("character %d splits a UTF-16 surrogate pair", wanted)
		}
	}
	if units == wanted {
		return lineEnd, nil
	}
	return 0, fmt.Errorf("character %d is outside line", wanted)
}

func TestUserFormEditStructuredFailuresAreAtomic(t *testing.T) {
	root := t.TempDir()
	s := &Server{opts: Options{RootDir: root, Config: config.Default()}}
	uri := pathToFileURI(filepath.Join(root, "src/forms/specs/Main.json"))
	source := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"submit","type":"CommandButton","name":"Submit","left":1,"top":2,"width":10,"height":8}]}`
	params := userFormEditParams{
		URI: uri, Version: 2, Text: source,
		Operations: []formsedit.Operation{
			{Type: formsedit.MoveControl, ControlID: "submit", Left: new(4.0), Top: new(5.0)},
			{Type: formsedit.ResizeControl, ControlID: "submit", Width: new(20.0)},
		},
	}
	result := s.userFormEdit(params)
	if result.Error == nil || result.Error.Code != compiler.Invalid || len(result.Edits) != 0 {
		t.Fatalf("partial batch result: %+v", result)
	}
	if len(result.Error.Diagnostics) == 0 {
		t.Fatalf("missing structured diagnostic: %+v", result.Error)
	}

	duplicateKey := strings.Replace(source, `"left":1`, `"left":1,"left":2`, 1)
	result = s.userFormEdit(userFormEditParams{URI: uri, Version: 3, Text: duplicateKey, Operations: params.Operations[:1]})
	if result.Error == nil || result.Error.Code != compiler.Conflict || len(result.Edits) != 0 {
		t.Fatalf("ambiguous source result: %+v", result)
	}

	yamlURI := pathToFileURI(filepath.Join(root, "src/forms/specs/Main.yaml"))
	missingCoordinates := formsedit.Operation{Type: formsedit.SetParent, ControlID: "submit"}
	result = s.userFormEdit(userFormEditParams{URI: yamlURI, Version: 4, Text: previewSource, Operations: []formsedit.Operation{missingCoordinates}})
	if result.Error == nil || result.Error.Code != compiler.Invalid || len(result.Edits) != 0 {
		t.Fatalf("typed SetParent accepted omitted coordinates: %+v", result)
	}
}

func TestUserFormEditRequestPayloadValidation(t *testing.T) {
	valid := `{"uri":"file:///x/src/forms/specs/Main.json","version":1,"text":"{}","operations":[{"type":"moveControl","controlId":"c","left":1,"top":2}]}`
	tests := []struct {
		name string
		body string
		code string
	}{
		{name: "missing request fields", body: `{}`, code: compiler.Invalid},
		{name: "unknown request field", body: strings.Replace(valid, `"text":"{}"`, `"text":"{}","other":1`, 1), code: compiler.Invalid},
		{name: "duplicate request field", body: strings.Replace(valid, `"version":1`, `"version":1,"version":2`, 1), code: compiler.Invalid},
		{name: "missing geometry", body: strings.Replace(valid, `,"top":2`, "", 1), code: compiler.Invalid},
		{name: "irrelevant operation field", body: strings.Replace(valid, `"top":2`, `"top":2,"field":"left"`, 1), code: compiler.Invalid},
		{name: "unknown operation field", body: strings.Replace(valid, `"top":2`, `"top":2,"mystery":1`, 1), code: compiler.Invalid},
		{name: "duplicate operation field", body: strings.Replace(valid, `"left":1`, `"left":1,"left":2`, 1), code: compiler.Invalid},
		{name: "nonnumeric geometry", body: strings.Replace(valid, `"left":1`, `"left":"1"`, 1), code: compiler.Invalid},
		{name: "nonfinite geometry", body: strings.Replace(valid, `"left":1`, `"left":1e400`, 1), code: compiler.Invalid},
		{name: "backend-only operation", body: strings.Replace(valid, `"type":"moveControl"`, `"type":"reorderControl"`, 1), code: compiler.Unsupported},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			_, requestError := parseUserFormEditParams([]byte(test.body))
			if requestError == nil || requestError.Code != test.code {
				t.Fatalf("parse result: %+v", requestError)
			}
		})
	}
}

func TestUserFormEditStructuralPayloadValidation(t *testing.T) {
	validAdd := `{"type":"addControl","control":{"id":"new","name":"NewLabel","type":"Label"}}`
	for _, test := range []struct {
		name string
		raw  string
		code string
	}{
		{name: "valid add", raw: validAdd},
		{name: "property bag remains open", raw: strings.Replace(validAdd, `"type":"Label"`, `"type":"Label","properties":{"Vendor.Key":null}`, 1)},
		{name: "Frame parent insertion", raw: strings.Replace(validAdd, `"type":"Label"`, `"type":"Label","parentId":"frame"`, 1)},
		{name: "root Page insertion remains unsupported", raw: strings.Replace(validAdd, `"type":"Label"`, `"type":"pAgE"`, 1), code: compiler.Unsupported},
		{name: "custom insertion remains unsupported", raw: strings.Replace(validAdd, `"type":"Label"`, `"type":"VendorControl","progId":"Vendor.Control"`, 1), code: compiler.Unsupported},
		{name: "unknown operation field", raw: strings.Replace(validAdd, `"control":`, `"extra":true,"control":`, 1), code: compiler.Invalid},
		{name: "unknown control field", raw: strings.Replace(validAdd, `"type":"Label"`, `"type":"Label","mystery":1`, 1), code: compiler.Invalid},
		{name: "duplicate control field", raw: strings.Replace(validAdd, `"name":"NewLabel"`, `"name":"NewLabel","name":"Other"`, 1), code: compiler.Invalid},
		{name: "case alias for id", raw: strings.Replace(validAdd, `"id":"new"`, `"id":"new","ID":"other"`, 1), code: compiler.Invalid},
		{name: "case alias for name", raw: strings.Replace(validAdd, `"name":"NewLabel"`, `"name":"NewLabel","Name":"Other"`, 1), code: compiler.Invalid},
		{name: "case alias for type", raw: strings.Replace(validAdd, `"type":"Label"`, `"type":"Label","Type":"TextBox"`, 1), code: compiler.Invalid},
		{name: "nested duplicate field", raw: strings.Replace(validAdd, `"type":"Label"`, `"type":"Label","tabs":[{"name":"A","name":"B"}]`, 1), code: compiler.Invalid},
		{name: "nested noncanonical tab field", raw: strings.Replace(validAdd, `"type":"Label"`, `"type":"Label","tabs":[{"name":"A","Caption":"A"}]`, 1), code: compiler.Invalid},
		{name: "nested noncanonical observed field", raw: strings.Replace(validAdd, `"type":"Label"`, `"type":"Label","observed":{"Width":10}`, 1), code: compiler.Invalid},
		{name: "null control", raw: `{"type":"addControl","control":null}`, code: compiler.Invalid},
		{name: "missing id", raw: `{"type":"addControl","control":{"name":"NewLabel","type":"Label"}}`, code: compiler.Invalid},
		{name: "null id", raw: `{"type":"addControl","control":{"id":null,"name":"NewLabel","type":"Label"}}`, code: compiler.Invalid},
		{name: "empty name", raw: `{"type":"addControl","control":{"id":"new","name":"","type":"Label"}}`, code: compiler.Invalid},
		{name: "null type", raw: `{"type":"addControl","control":{"id":"new","name":"NewLabel","type":null}}`, code: compiler.Invalid},
		{name: "null optional field", raw: strings.Replace(validAdd, `"type":"Label"`, `"type":"Label","caption":null`, 1), code: compiler.Invalid},
		{name: "null tabs", raw: strings.Replace(validAdd, `"type":"Label"`, `"type":"Label","tabs":null`, 1), code: compiler.Invalid},
		{name: "null selected index", raw: strings.Replace(validAdd, `"type":"Label"`, `"type":"Label","selectedIndex":null`, 1), code: compiler.Invalid},
		{name: "nested controls", raw: strings.Replace(validAdd, `"type":"Label"`, `"type":"Label","controls":[{"id":"child","name":"Child","type":"Label"}]`, 1), code: compiler.Invalid},
		{name: "remove null cascade", raw: `{"type":"removeControl","controlId":"frame","cascade":null}`, code: compiler.Invalid},
		{name: "remove non-boolean cascade", raw: `{"type":"removeControl","controlId":"frame","cascade":1}`, code: compiler.Invalid},
		{name: "set parent requires coordinates", raw: `{"type":"setParent","controlId":"new","parentId":"frame"}`, code: compiler.Invalid},
		{name: "reorder remains unsupported", raw: `{"type":"reorderControl","controlId":"new","index":0}`, code: compiler.Unsupported},
	} {
		t.Run(test.name, func(t *testing.T) {
			operation, requestError := parseUserFormEditOperation([]byte(test.raw))
			if test.code != "" {
				if requestError == nil || requestError.Code != test.code {
					t.Fatalf("parse result = %+v, want error code %s (operation %+v)", requestError, test.code, operation)
				}
				return
			}
			if requestError != nil || operation.Type != formsedit.AddControl || operation.Control == nil {
				t.Fatalf("parse result = %+v, %+v", operation, requestError)
			}
		})
	}

	for _, controlType := range []string{"MultiPage", "TabStrip"} {
		t.Run("empty "+controlType+" explicit payload", func(t *testing.T) {
			extra := `,"selectedIndex":-1`
			if controlType == "TabStrip" {
				extra = `,"tabs":[]` + extra
			}
			raw := []byte(fmt.Sprintf(`{"type":"addControl","control":{"id":"empty","name":"Empty","type":%q%s}}`, controlType, extra))
			operation, requestError := parseUserFormEditOperation(raw)
			if requestError != nil || operation.Control == nil {
				t.Fatalf("parse result = %+v, %+v", operation, requestError)
			}
			if (controlType == "TabStrip" && operation.Control.Tabs == nil) || (controlType == "MultiPage" && operation.Control.Tabs != nil) || operation.Control.SelectedIndex == nil || *operation.Control.SelectedIndex != -1 {
				t.Fatalf("empty %s payload lost canonical defaults: %+v", controlType, operation.Control)
			}
		})
	}
	for _, controlType := range []string{"MultiPage", "TabStrip"} {
		raw := []byte(fmt.Sprintf(`{"type":"addControl","control":{"id":"empty","name":"Empty","type":%q}}`, controlType))
		operation, requestError := parseUserFormEditOperation(raw)
		if requestError != nil || operation.Control == nil || operation.Control.SelectedIndex != nil || operation.Control.Tabs != nil {
			t.Fatalf("omitted %s fields were rewritten: %+v, %+v", controlType, operation.Control, requestError)
		}
	}
	for _, operations := range [][]formsedit.Operation{
		{{Type: formsedit.AddControl, Control: &spec.FormSpecControl{ID: "new", Name: "New", Type: "Label"}}, {Type: formsedit.MoveControl, ControlID: "old", Left: new(1.0), Top: new(2.0)}},
		{{Type: formsedit.AddControl, Control: &spec.FormSpecControl{ID: "new", Name: "New", Type: "Label"}}, {Type: formsedit.RemoveControl, ControlID: "old"}},
		{{Type: formsedit.RemoveControl, ControlID: "old"}, {Type: formsedit.SetControlProperty, ControlID: "other", Field: "visible", Value: true, ValuePresent: true}},
	} {
		if requestError := validateUserFormEditTransaction(operations); requestError == nil || requestError.Code != compiler.Invalid {
			t.Fatalf("mixed structural transaction accepted: %+v", requestError)
		}
	}
}

func TestUserFormEditInitialMultiPageAndTabStrip(t *testing.T) {
	root := t.TempDir()
	server := &Server{opts: Options{RootDir: root, Config: config.Default()}}
	fixtures := []struct {
		name   string
		format string
		source string
	}{
		{
			name: "yaml", format: "yaml",
			source: "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: Main\ncontrols: []\n",
		},
		{
			name: "json", format: "json",
			source: `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[]}`,
		},
	}
	for _, fixture := range fixtures {
		t.Run(fixture.name+"/MultiPage", func(t *testing.T) {
			uri := pathToFileURI(filepath.Join(root, "src/forms/specs/Main."+fixture.format))
			operations := []any{
				map[string]any{"type": "addControl", "control": map[string]any{
					"id": "multi", "name": "MultiPage1", "type": "MultiPage", "selectedIndex": -1,
				}},
				map[string]any{"type": "addPage", "control": map[string]any{
					"id": "page-one", "name": "PageOne", "type": "Page", "parentId": "multi",
				}},
				map[string]any{"type": "addPage", "control": map[string]any{
					"id": "page-two", "name": "PageTwo", "type": "Page", "parentId": "multi",
				}},
			}
			params := parseUserFormEditTestParams(t, uri, fixture.source, operations)
			result := server.userFormEdit(params)
			if result.Error != nil || result.Version != params.Version || len(result.Edits) == 0 {
				t.Fatalf("initial MultiPage transaction failed: %+v", result)
			}
			updated, err := applyLSPTextEdits(fixture.source, result.Edits)
			if err != nil {
				t.Fatal(err)
			}
			document, err := spec.ParseFormSpec(spec.SpecInput{Format: fixture.format}, []byte(updated))
			if err != nil {
				t.Fatalf("initial MultiPage transaction produced invalid %s: %v\n%s", fixture.format, err, updated)
			}
			if len(document.Controls) != 3 {
				t.Fatalf("MultiPage controls = %+v, want owner and two Pages", document.Controls)
			}
			owner := document.Controls[0]
			if owner.ID != "multi" || owner.ParentID != "" || !strings.EqualFold(owner.Type, "MultiPage") ||
				owner.SelectedIndex == nil || *owner.SelectedIndex != 0 || len(owner.Tabs) != 0 {
				t.Fatalf("canonical MultiPage owner = %+v", owner)
			}
			for index, want := range []struct {
				id, name string
			}{{"page-one", "PageOne"}, {"page-two", "PageTwo"}} {
				page := document.Controls[index+1]
				if page.ID != want.id || page.Name != want.name || !strings.EqualFold(page.Type, "Page") || page.ParentID != owner.ID ||
					page.Left != nil || page.Top != nil || page.Width != nil || page.Height != nil {
					t.Errorf("canonical Page %d = %+v, want parent %q and no geometry", index, page, owner.ID)
				}
			}
		})

		t.Run(fixture.name+"/TabStrip", func(t *testing.T) {
			uri := pathToFileURI(filepath.Join(root, "src/forms/specs/Main."+fixture.format))
			operations := []any{map[string]any{"type": "addControl", "control": map[string]any{
				"id": "tabs", "name": "TabStrip1", "type": "TabStrip", "selectedIndex": 0,
				"tabs": []any{
					map[string]any{"name": "Tab1", "caption": "Tab1"},
					map[string]any{"name": "Tab2", "caption": "Tab2"},
				},
			}}}
			params := parseUserFormEditTestParams(t, uri, fixture.source, operations)
			result := server.userFormEdit(params)
			if result.Error != nil || len(result.Edits) == 0 {
				t.Fatalf("initial TabStrip insertion failed: %+v", result)
			}
			updated, err := applyLSPTextEdits(fixture.source, result.Edits)
			if err != nil {
				t.Fatal(err)
			}
			document, err := spec.ParseFormSpec(spec.SpecInput{Format: fixture.format}, []byte(updated))
			if err != nil || len(document.Controls) != 1 {
				t.Fatalf("initial TabStrip result = %+v, %v\n%s", document.Controls, err, updated)
			}
			control := document.Controls[0]
			if control.ID != "tabs" || !strings.EqualFold(control.Type, "TabStrip") || control.SelectedIndex == nil || *control.SelectedIndex != 0 || len(control.Tabs) != 2 {
				t.Fatalf("canonical TabStrip = %+v", control)
			}
			for index, want := range []string{"Tab1", "Tab2"} {
				tab := control.Tabs[index]
				if tab.Name != want || tab.Caption == nil || *tab.Caption != want || tab.ControlTipText != nil || tab.Tag != nil ||
					tab.Accelerator != nil || tab.Enabled != nil || tab.Visible != nil {
					t.Errorf("canonical TabStrip tab %d = %+v", index, tab)
				}
			}
		})

		for _, test := range []struct {
			name    string
			control map[string]any
		}{
			{name: "empty-MultiPage-selected-minus-one", control: map[string]any{"id": "empty-multi", "name": "EmptyMulti", "type": "MultiPage", "selectedIndex": -1}},
			{name: "empty-MultiPage-omitted-selection", control: map[string]any{"id": "empty-multi-omitted", "name": "EmptyMultiOmitted", "type": "MultiPage"}},
			{name: "empty-TabStrip-selected-minus-one", control: map[string]any{"id": "empty-tabs", "name": "EmptyTabs", "type": "TabStrip", "selectedIndex": -1, "tabs": []any{}}},
			{name: "empty-TabStrip-omitted-selection-and-tabs", control: map[string]any{"id": "empty-tabs-omitted", "name": "EmptyTabsOmitted", "type": "TabStrip"}},
		} {
			t.Run(fixture.name+"/legacy-"+test.name, func(t *testing.T) {
				uri := pathToFileURI(filepath.Join(root, "src/forms/specs/Main."+fixture.format))
				params := parseUserFormEditTestParams(t, uri, fixture.source, []any{
					map[string]any{"type": "addControl", "control": test.control},
				})
				result := server.userFormEdit(params)
				if result.Error != nil || len(result.Edits) == 0 {
					t.Fatalf("legacy %s insertion failed: %+v", test.name, result)
				}
				updated, err := applyLSPTextEdits(fixture.source, result.Edits)
				if err != nil {
					t.Fatal(err)
				}
				if _, err := spec.ParseFormSpec(spec.SpecInput{Format: fixture.format}, []byte(updated)); err != nil {
					t.Fatalf("legacy %s result is invalid: %v\n%s", test.name, err, updated)
				}
			})
		}
	}
}

func TestUserFormEditInitialMultiPageSupportsExistingContainers(t *testing.T) {
	root := t.TempDir()
	server := &Server{opts: Options{RootDir: root, Config: config.Default()}}
	fixtures := []struct {
		name   string
		format string
		source string
	}{
		{
			name: "yaml", format: "yaml",
			source: "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: Main\ncontrols:\n  - id: existing-multi\n    name: ExistingMulti\n    type: MultiPage\n    selectedIndex: 0\n  - id: existing-page\n    name: ExistingPage\n    type: Page\n    parentId: existing-multi\n  - id: existing-frame\n    name: ExistingFrame\n    type: Frame\n    parentId: existing-page\n",
		},
		{
			name: "json", format: "json",
			source: `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"existing-multi","name":"ExistingMulti","type":"MultiPage","selectedIndex":0},{"id":"existing-page","name":"ExistingPage","type":"Page","parentId":"existing-multi"},{"id":"existing-frame","name":"ExistingFrame","type":"Frame","parentId":"existing-page"}]}`,
		},
	}
	for _, fixture := range fixtures {
		for _, parent := range []struct {
			name string
			id   string
		}{{name: "Frame", id: "existing-frame"}, {name: "Page", id: "existing-page"}} {
			t.Run(fixture.name+"/inside-"+parent.name, func(t *testing.T) {
				uri := pathToFileURI(filepath.Join(root, "src/forms/specs/Main."+fixture.format))
				operations := []any{
					map[string]any{"type": "addControl", "control": map[string]any{
						"id": "multi", "name": "MultiPage1", "type": "MultiPage", "selectedIndex": -1, "parentId": parent.id,
					}},
					map[string]any{"type": "addPage", "control": map[string]any{
						"id": "page-one", "name": "PageOne", "type": "Page", "parentId": "multi",
					}},
					map[string]any{"type": "addPage", "control": map[string]any{
						"id": "page-two", "name": "PageTwo", "type": "Page", "parentId": "multi",
					}},
				}
				params := parseUserFormEditTestParams(t, uri, fixture.source, operations)
				result := server.userFormEdit(params)
				if result.Error != nil || len(result.Edits) == 0 {
					t.Fatalf("MultiPage insertion inside %s failed: %+v", parent.name, result)
				}
				updated, err := applyLSPTextEdits(fixture.source, result.Edits)
				if err != nil {
					t.Fatal(err)
				}
				document, err := spec.ParseFormSpec(spec.SpecInput{Format: fixture.format}, []byte(updated))
				if err != nil {
					t.Fatalf("nested MultiPage transaction produced invalid %s: %v\n%s", fixture.format, err, updated)
				}
				controls := make(map[string]spec.FormSpecControl, len(document.Controls))
				for _, control := range document.Controls {
					controls[control.ID] = control
				}
				owner := controls["multi"]
				if owner.ID != "multi" || owner.ParentID != parent.id || owner.SelectedIndex == nil || *owner.SelectedIndex != 0 {
					t.Fatalf("canonical nested MultiPage = %+v, want parent %q and selectedIndex 0", owner, parent.id)
				}
				for _, pageID := range []string{"page-one", "page-two"} {
					page := controls[pageID]
					if page.ID != pageID || !strings.EqualFold(page.Type, "Page") || page.ParentID != owner.ID {
						t.Errorf("canonical Page %q = %+v, want parent %q", pageID, page, owner.ID)
					}
				}
			})
		}
	}
}

func TestUserFormEditInitialMultiPageRejectsMalformedTransactions(t *testing.T) {
	root := t.TempDir()
	server := &Server{opts: Options{RootDir: root, Config: config.Default()}}
	source := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[]}`
	uri := pathToFileURI(filepath.Join(root, "src/forms/specs/Main.json"))
	baseOperations := func() []any {
		return []any{
			map[string]any{"type": "addControl", "control": map[string]any{
				"id": "multi", "name": "MultiPage1", "type": "MultiPage", "selectedIndex": -1,
			}},
			map[string]any{"type": "addPage", "control": map[string]any{
				"id": "page-one", "name": "PageOne", "type": "Page", "parentId": "multi",
			}},
			map[string]any{"type": "addPage", "control": map[string]any{
				"id": "page-two", "name": "PageTwo", "type": "Page", "parentId": "multi",
			}},
		}
	}
	tests := []struct {
		name     string
		mutate   func([]any)
		source   string
		wantCode string
	}{
		{name: "wrong parent", mutate: func(operations []any) {
			operations[2].(map[string]any)["control"].(map[string]any)["parentId"] = "other"
		}, wantCode: compiler.Invalid},
		{name: "invalid child", mutate: func(operations []any) { operations[1].(map[string]any)["control"].(map[string]any)["type"] = "Label" }, wantCode: compiler.Invalid},
		{name: "duplicate IDs", mutate: func(operations []any) { operations[2].(map[string]any)["control"].(map[string]any)["id"] = "page-one" }, wantCode: compiler.Invalid},
		{name: "duplicate names ignoring case", mutate: func(operations []any) { operations[2].(map[string]any)["control"].(map[string]any)["name"] = "pageone" }, wantCode: compiler.Invalid},
		{name: "unsupported owner ProgID", mutate: func(operations []any) {
			operations[0].(map[string]any)["control"].(map[string]any)["progId"] = "Vendor.MultiPage"
		}, wantCode: compiler.Unsupported},
		{name: "mixed arbitrary batch", mutate: func(operations []any) {
			operations[0] = map[string]any{"type": "setFormProperty", "field": "caption", "value": "Changed"}
		}, wantCode: compiler.Invalid},
		{name: "existing ID collision", mutate: func([]any) {}, source: `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"multi","name":"Existing","type":"Label"}]}`},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			operations := baseOperations()
			test.mutate(operations)
			requestSource := source
			if test.source != "" {
				requestSource = test.source
			}
			payload := marshalUserFormEditTestRequest(t, uri, requestSource, operations)
			params, requestError := parseUserFormEditParams(payload)
			if requestError != nil {
				if requestError.Code != test.wantCode {
					t.Fatalf("request error = %+v, want %s", requestError, test.wantCode)
				}
				return
			}
			result := server.userFormEdit(params)
			if result.Error == nil || (test.wantCode != "" && result.Error.Code != test.wantCode) || len(result.Edits) != 0 {
				t.Fatalf("malformed transaction result = %+v, want %s and no edits", result, test.wantCode)
			}
			if params.Text != requestSource {
				t.Fatalf("failed transaction changed source: %q", params.Text)
			}
		})
	}
}

func TestUserFormEditInitialTabStripPayloadIsNarrow(t *testing.T) {
	valid := `{"type":"addControl","control":{"id":"tabs","name":"Tabs","type":"TabStrip","selectedIndex":0,"tabs":[{"name":"Tab1","caption":"Tab1"},{"name":"Tab2","caption":"Tab2"}]}}`
	emptyCaption := strings.Replace(valid, `"caption":"Tab1"`, `"caption":""`, 1)
	if _, requestError := parseUserFormEditOperation([]byte(emptyCaption)); requestError != nil {
		t.Fatalf("valid empty TabStrip caption was rejected: %+v", requestError)
	}
	for _, test := range []struct {
		name string
		raw  string
	}{
		{name: "extra tab field", raw: strings.Replace(valid, `"caption":"Tab1"`, `"caption":"Tab1","enabled":false`, 1)},
		{name: "duplicate tab names ignoring case", raw: strings.Replace(valid, `"name":"Tab2"`, `"name":"tab1"`, 1)},
		{name: "invalid selected index", raw: strings.Replace(valid, `"selectedIndex":0`, `"selectedIndex":1`, 1)},
		{name: "missing selected index", raw: strings.Replace(valid, `,"selectedIndex":0`, ``, 1)},
		{name: "non-string caption", raw: strings.Replace(valid, `"caption":"Tab1"`, `"caption":1`, 1)},
		{name: "only one initial tab", raw: strings.Replace(valid, `,{"name":"Tab2","caption":"Tab2"}`, ``, 1)},
		{name: "missing caption", raw: strings.Replace(valid, `,"caption":"Tab1"`, ``, 1)},
		{name: "custom ProgID", raw: strings.Replace(valid, `"type":"TabStrip"`, `"type":"TabStrip","progId":"Vendor.Tabs"`, 1)},
	} {
		t.Run(test.name, func(t *testing.T) {
			if _, requestError := parseUserFormEditOperation([]byte(test.raw)); requestError == nil {
				t.Fatalf("malformed initial TabStrip payload was accepted: %s", test.raw)
			}
		})
	}
}

func parseUserFormEditTestParams(t *testing.T, uri, source string, operations []any) userFormEditParams {
	t.Helper()
	params, requestError := parseUserFormEditParams(marshalUserFormEditTestRequest(t, uri, source, operations))
	if requestError != nil {
		t.Fatalf("parse test request: %+v", requestError)
	}
	return params
}

func marshalUserFormEditTestRequest(t *testing.T, uri, source string, operations []any) []byte {
	t.Helper()
	request, err := json.Marshal(map[string]any{
		"uri": uri, "version": 23, "text": source, "operations": operations,
	})
	if err != nil {
		t.Fatal(err)
	}
	return request
}

func TestUserFormEditStructuralYAMLAndJSON(t *testing.T) {
	root := t.TempDir()
	server := &Server{opts: Options{RootDir: root, Config: config.Default()}}
	fixtures := []struct {
		name   string
		format string
		source string
	}{
		{
			name: "yaml", format: "yaml",
			source: "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: Main\n  width: 320\n  height: 240\ncontrols:\n  - id: frame\n    name: Frame1\n    type: Frame\n    left: 0\n    top: 0\n    width: 160\n    height: 100\n  - id: child\n    parentId: frame\n    name: ChildLabel\n    type: Label\n    left: 4\n    top: 4\n    width: 80\n    height: 16\n",
		},
		{
			name: "json", format: "json",
			source: `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main","width":320,"height":240},"controls":[{"id":"frame","name":"Frame1","type":"Frame","left":0,"top":0,"width":160,"height":100},{"id":"child","parentId":"frame","name":"ChildLabel","type":"Label","left":4,"top":4,"width":80,"height":16}]}`,
		},
	}
	for _, fixture := range fixtures {
		t.Run(fixture.name, func(t *testing.T) {
			uri := pathToFileURI(filepath.Join(root, "src/forms/specs/Main."+fixture.format))
			add, requestError := parseUserFormEditOperation([]byte(`{"type":"addControl","control":{"id":"status","name":"StatusLabel","type":"Label","left":12,"top":18,"width":90,"height":18}}`))
			if requestError != nil {
				t.Fatal(requestError)
			}
			added := server.userFormEdit(userFormEditParams{URI: uri, Version: 31, Text: fixture.source, Operations: []formsedit.Operation{add}})
			if added.Error != nil || added.Version != 31 || len(added.Edits) == 0 {
				t.Fatalf("add control result: %+v", added)
			}
			addedSource, err := applyLSPTextEdits(fixture.source, added.Edits)
			if err != nil {
				t.Fatal(err)
			}
			input := spec.SpecInput{Format: fixture.format}
			document, err := spec.ParseFormSpec(input, []byte(addedSource))
			if err != nil || len(document.Controls) != 3 || document.Controls[2].ID != "status" {
				t.Fatalf("added %s document: %+v, %v\n%s", fixture.format, document.Controls, err, addedSource)
			}

			remove, requestError := parseUserFormEditOperation([]byte(`{"type":"removeControl","controlId":"frame","cascade":true}`))
			if requestError != nil {
				t.Fatal(requestError)
			}
			removed := server.userFormEdit(userFormEditParams{URI: uri, Version: 32, Text: addedSource, Operations: []formsedit.Operation{remove}})
			if removed.Error != nil || removed.Version != 32 || len(removed.Edits) == 0 {
				t.Fatalf("cascade remove result: %+v", removed)
			}
			removedSource, err := applyLSPTextEdits(addedSource, removed.Edits)
			if err != nil {
				t.Fatal(err)
			}
			document, err = spec.ParseFormSpec(input, []byte(removedSource))
			if err != nil || len(document.Controls) != 1 || document.Controls[0].ID != "status" {
				t.Fatalf("cascade result: %+v, %v\n%s", document.Controls, err, removedSource)
			}
		})
	}
}

func TestUserFormEditStructuralScope(t *testing.T) {
	root := t.TempDir()
	server := &Server{opts: Options{RootDir: root, Config: config.Default()}}
	fixtures := []struct {
		format string
		source string
	}{
		{format: "json", source: `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"frame","name":"Frame1","type":"Frame"},{"id":"multi","name":"MultiPage1","type":"MultiPage","selectedIndex":0},{"id":"page1","name":"Page1","type":"Page","parentId":"multi"},{"id":"page2","name":"Page2","type":"pAgE","parentId":"multi"},{"id":"leaf","name":"Label1","type":"Label","parentId":"page2"}]}`},
		{format: "yaml", source: "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: Main\ncontrols:\n  - id: frame\n    name: Frame1\n    type: Frame\n  - id: multi\n    name: MultiPage1\n    type: MultiPage\n    selectedIndex: 0\n    controls:\n      - id: page1\n        name: Page1\n        type: Page\n      - id: page2\n        name: Page2\n        type: pAgE\n        controls:\n          - id: leaf\n            name: Label1\n            type: Label\n"},
	}
	for _, fixture := range fixtures {
		t.Run(fixture.format, func(t *testing.T) {
			uri := pathToFileURI(filepath.Join(root, "src/forms/specs/Main."+fixture.format))
			for _, test := range []struct {
				name    string
				op      formsedit.Operation
				allowed bool
			}{
				{name: "Frame child insertion", op: formsedit.Operation{Type: formsedit.AddControl, Control: &spec.FormSpecControl{ID: "new", Name: "Label2", Type: "Label", ParentID: "frame"}}, allowed: true},
				{name: "Page insertion", op: formsedit.Operation{Type: formsedit.AddControl, Control: &spec.FormSpecControl{ID: "new", Name: "Page3", Type: "pAgE", ParentID: "multi"}}},
				{name: "custom insertion", op: formsedit.Operation{Type: formsedit.AddControl, Control: &spec.FormSpecControl{ID: "new", Name: "Custom1", Type: "VendorControl", ProgID: "Vendor.Control"}}},
				{name: "Page direct removal", op: formsedit.Operation{Type: formsedit.RemoveControl, ControlID: "page2", Cascade: true}},
			} {
				t.Run(test.name, func(t *testing.T) {
					// These are valid generic semantic edits. Frame children are
					// supported by the Designer boundary; Page and custom edits are not.
					if _, err := formsedit.Apply(spec.SpecInput{Format: fixture.format}, []byte(fixture.source), []formsedit.Operation{test.op}); err != nil {
						t.Fatalf("generic edit unexpectedly rejected: %v", err)
					}
					params := userFormEditParams{URI: uri, Version: 41, Text: fixture.source, Operations: []formsedit.Operation{test.op}}
					result := server.userFormEdit(params)
					if test.allowed {
						if result.Error != nil || result.Version != 41 || len(result.Edits) == 0 {
							t.Fatalf("supported handler edit failed: %+v", result)
						}
						updated, err := applyLSPTextEdits(fixture.source, result.Edits)
						if err != nil {
							t.Fatal(err)
						}
						document, err := spec.ParseFormSpec(spec.SpecInput{Format: fixture.format}, []byte(updated))
						if err != nil || len(document.Controls) == 0 || document.Controls[len(document.Controls)-1].ParentID != "frame" {
							t.Fatalf("Frame child parent was not retained: controls=%+v err=%v", document.Controls, err)
						}
					} else if result.Error == nil || result.Error.Code != compiler.Unsupported || result.Version != 41 || len(result.Edits) != 0 {
						t.Fatalf("out-of-scope handler edit: %+v", result)
					}
					raw, err := json.Marshal(map[string]any{"uri": uri, "version": 41, "text": fixture.source, "operations": params.Operations})
					if err != nil {
						t.Fatal(err)
					}
					decoded, requestError := parseUserFormEditParams(raw)
					if requestError != nil {
						if requestError.Code != compiler.Unsupported {
							t.Fatalf("wrong RPC rejection: %+v", requestError)
						}
						return
					}
					result = server.userFormEdit(decoded)
					if test.allowed {
						if result.Error != nil || len(result.Edits) == 0 {
							t.Fatalf("supported raw RPC edit failed: %+v", result)
						}
					} else if result.Error == nil || result.Error.Code != compiler.Unsupported || len(result.Edits) != 0 {
						t.Fatalf("out-of-scope raw RPC edit: %+v", result)
					}
				})
			}
			for _, id := range []string{"leaf", "multi"} {
				result := server.userFormEdit(userFormEditParams{URI: uri, Version: 41, Text: fixture.source, Operations: []formsedit.Operation{{Type: formsedit.RemoveControl, ControlID: id, Cascade: true}}})
				if result.Error != nil || len(result.Edits) == 0 {
					t.Fatalf("supported deletion of %s failed: %+v", id, result)
				}
				updated, err := applyLSPTextEdits(fixture.source, result.Edits)
				if err != nil {
					t.Fatal(err)
				}
				if _, err := spec.ParseFormSpec(spec.SpecInput{Format: fixture.format}, []byte(updated)); err != nil {
					t.Fatalf("supported deletion produced invalid source: %v", err)
				}
			}
		})
	}
}

func TestUserFormEditDesignerAdditionProgIDValidation(t *testing.T) {
	builtInProgID, ok := spec.BuiltInControlProgID("Label")
	if !ok {
		t.Fatal("Label built-in ProgID is missing")
	}
	formats := []struct {
		name   string
		input  spec.SpecInput
		source string
	}{
		{
			name: "yaml", input: spec.SpecInput{Format: "yaml"},
			source: "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: Main\ncontrols: []\n",
		},
		{
			name: "json", input: spec.SpecInput{Format: "json"},
			source: `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[]}`,
		},
	}
	for _, format := range formats {
		t.Run(format.name, func(t *testing.T) {
			root := t.TempDir()
			server := &Server{opts: Options{RootDir: root, Config: config.Default()}}
			uri := pathToFileURI(filepath.Join(root, "src/forms/specs/Main."+format.input.Format))
			marshalAdd := func(progID *string) []byte {
				control := map[string]any{"id": "new", "name": "NewLabel", "type": "Label"}
				if progID != nil {
					control["progId"] = *progID
				}
				payload, err := json.Marshal(map[string]any{"type": "addControl", "control": control})
				if err != nil {
					t.Fatal(err)
				}
				return payload
			}

			for _, invalid := range []struct {
				name   string
				progID string
			}{
				{name: "custom", progID: "Vendor.Control"},
				{name: "whitespace-padded built-in", progID: " " + builtInProgID + " "},
			} {
				t.Run("reject "+invalid.name+" ProgID at Designer boundary", func(t *testing.T) {
					control := &spec.FormSpecControl{ID: "new", Name: "NewLabel", Type: "Label", ProgID: invalid.progID}
					operation := formsedit.Operation{Type: formsedit.AddControl, Control: control}
					if _, requestError := parseUserFormEditOperation(marshalAdd(&control.ProgID)); requestError == nil || requestError.Code != compiler.Unsupported {
						t.Errorf("raw decoder error = %+v, want %s", requestError, compiler.Unsupported)
					}

					generic, err := formsedit.Apply(format.input, []byte(format.source), []formsedit.Operation{operation})
					if err != nil {
						t.Fatalf("generic Apply rejected the semantic edit: %v", err)
					}
					if _, err := compiler.CompileNew(generic.Document, 932); err == nil {
						t.Error("CompileNew accepted a built-in Label with an unsupported ProgID")
					}

					result := server.userFormEdit(userFormEditParams{URI: uri, Version: 1, Text: format.source, Operations: []formsedit.Operation{operation}})
					if result.Error == nil || result.Error.Code != compiler.Unsupported || len(result.Edits) != 0 {
						t.Errorf("direct handler result = %+v, want unsupported with no edits", result)
					}
				})
			}

			for _, variant := range []struct {
				name   string
				progID *string
			}{
				{name: "omitted"},
				{name: "exact", progID: &builtInProgID},
				{name: "case-insensitive", progID: new(strings.ToLower(builtInProgID))},
			} {
				t.Run("accept "+variant.name, func(t *testing.T) {
					operation, requestError := parseUserFormEditOperation(marshalAdd(variant.progID))
					if requestError != nil {
						t.Fatalf("raw decoder rejected built-in ProgID variant: %+v", requestError)
					}
					result := server.userFormEdit(userFormEditParams{URI: uri, Version: 2, Text: format.source, Operations: []formsedit.Operation{operation}})
					if result.Error != nil || len(result.Edits) == 0 {
						t.Fatalf("direct handler rejected built-in ProgID variant: %+v", result)
					}
					edited, err := applyLSPTextEdits(format.source, result.Edits)
					if err != nil {
						t.Fatal(err)
					}
					document, err := spec.ParseFormSpec(format.input, []byte(edited))
					if err != nil {
						t.Fatalf("ParseFormSpec: %v", err)
					}
					if variant.progID != nil && document.Controls[0].ProgID != *variant.progID {
						t.Errorf("parsed ProgID = %q, want %q", document.Controls[0].ProgID, *variant.progID)
					}
					if _, err := compiler.CompileNew(document, 932); err != nil {
						t.Fatalf("CompileNew: %v", err)
					}
				})
			}
		})
	}
}

func TestUserFormEditNewBuiltInControlsCompile(t *testing.T) {
	formats := []struct {
		name   string
		input  spec.SpecInput
		source string
	}{
		{
			name: "yaml", input: spec.SpecInput{Format: "yaml"},
			source: "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: Main\n  width: 320\n  height: 240\ncontrols: []\n",
		},
		{
			name: "json", input: spec.SpecInput{Format: "json"},
			source: `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main","width":320,"height":240},"controls":[]}`,
		},
	}
	types := []struct {
		name          string
		width, height float64
		caption       bool
		selectedIndex bool
		emptyTabStrip bool
	}{
		{name: "Label", width: 72, height: 18, caption: true},
		{name: "TextBox", width: 120, height: 18},
		{name: "ComboBox", width: 120, height: 18},
		{name: "ListBox", width: 120, height: 72},
		{name: "CommandButton", width: 72, height: 24, caption: true},
		{name: "CheckBox", width: 72, height: 18, caption: true},
		{name: "OptionButton", width: 72, height: 18, caption: true},
		{name: "ToggleButton", width: 72, height: 18, caption: true},
		{name: "SpinButton", width: 18, height: 36},
		{name: "ScrollBar", width: 120, height: 18},
		{name: "Image", width: 72, height: 72},
		{name: "Frame", width: 144, height: 108, caption: true},
		{name: "MultiPage", width: 240, height: 180, selectedIndex: true},
		{name: "TabStrip", width: 240, height: 48, selectedIndex: true, emptyTabStrip: true},
	}
	for _, format := range formats {
		for _, controlType := range types {
			t.Run(format.name+"/"+controlType.name, func(t *testing.T) {
				id := "new-" + strings.ToLower(controlType.name)
				name := controlType.name + "1"
				payloadControl := map[string]any{
					"id": id, "name": name, "type": controlType.name,
					"left": 0, "top": 0, "width": controlType.width, "height": controlType.height,
				}
				if controlType.caption {
					payloadControl["caption"] = name
				}
				if controlType.selectedIndex {
					payloadControl["selectedIndex"] = -1
				}
				if controlType.emptyTabStrip {
					payloadControl["tabs"] = []any{}
				}
				payload, err := json.Marshal(map[string]any{
					"type":    "addControl",
					"control": payloadControl,
				})
				if err != nil {
					t.Fatal(err)
				}
				operation, requestError := parseUserFormEditOperation(payload)
				if requestError != nil {
					t.Fatal(requestError)
				}
				result, err := formsedit.Apply(format.input, []byte(format.source), []formsedit.Operation{operation})
				if err != nil {
					t.Fatalf("Apply: %v", err)
				}
				if len(result.Document.Controls) != 1 {
					t.Fatalf("Apply controls = %+v", result.Document.Controls)
				}
				control := result.Document.Controls[0]
				if control.ID != id || control.Name != name || control.Type != controlType.name || control.Width == nil || *control.Width != controlType.width || control.Height == nil || *control.Height != controlType.height {
					t.Fatalf("Apply control payload changed: %+v", control)
				}
				if controlType.selectedIndex {
					if control.SelectedIndex == nil || *control.SelectedIndex != -1 {
						t.Fatalf("empty %s defaults: %+v", controlType.name, control)
					}
				}
				if controlType.emptyTabStrip && (control.Tabs == nil || len(control.Tabs) != 0) {
					t.Fatalf("empty TabStrip tabs were not preserved: %+v", control)
				}
				if _, err := compiler.CompileNew(result.Document, 932); err != nil {
					t.Fatalf("CompileNew: %v", err)
				}
			})
		}
	}
}

func TestUserFormEditURIEligibilityMatchesPreview(t *testing.T) {
	root := t.TempDir()
	s := &Server{opts: Options{RootDir: root, Config: config.Default()}}
	jsonSource := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"submit","type":"CommandButton","name":"Submit"}]}`
	yamlSource := strings.Replace(previewSource, "controls: []", "controls:\n  - id: submit\n    type: CommandButton\n    name: Submit", 1)
	for _, test := range []struct {
		path   string
		source string
		ok     bool
	}{
		{path: "src/forms/specs/Main.json", source: jsonSource, ok: true},
		{path: "src/forms/specs/Main.yaml", source: yamlSource, ok: true},
		{path: "src/forms/specs/nested/Main.json", source: jsonSource},
		{path: "src/forms/Main.json", source: jsonSource},
		{path: "src/forms/specs/Main.txt", source: jsonSource},
	} {
		uri := pathToFileURI(filepath.Join(root, filepath.FromSlash(test.path)))
		preview := s.userFormPreview(userFormPreviewParams{URI: uri, Version: 5, Text: test.source})
		edit := s.userFormEdit(userFormEditParams{
			URI: uri, Version: 5, Text: test.source,
			Operations: []formsedit.Operation{{Type: formsedit.MoveControl, ControlID: "submit", Left: new(4.0), Top: new(5.0)}},
		})
		if test.ok && (preview.Error != nil || edit.Error != nil) {
			t.Fatalf("%s eligible mismatch: preview=%+v edit=%+v", test.path, preview.Error, edit.Error)
		}
		if !test.ok && (preview.Error == nil || edit.Error == nil) {
			t.Fatalf("%s was not rejected by both methods: preview=%+v edit=%+v", test.path, preview.Error, edit.Error)
		}
	}
}

func TestUserFormEditResponseHasEmptyEditArrayOnFailure(t *testing.T) {
	result := userFormEditResult{Version: 9, Edits: []protocol.TextEdit{}, Error: &userFormEditError{Code: compiler.Invalid, Message: "invalid request"}}
	encoded, err := json.Marshal(result)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(encoded), `"edits":[]`) {
		t.Fatalf("empty edits missing from failure response: %s", encoded)
	}
}
