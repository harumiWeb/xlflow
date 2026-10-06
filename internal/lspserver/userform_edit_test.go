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
	if initialized.Capabilities.Experimental["userFormEdit"] != true || initialized.Capabilities.Experimental["userFormPropertyEdit"] != true {
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

	unsupported := params.Operations[:1]
	unsupported[0].Type = formsedit.SetParent
	yamlURI := pathToFileURI(filepath.Join(root, "src/forms/specs/Main.yaml"))
	result = s.userFormEdit(userFormEditParams{URI: yamlURI, Version: 4, Text: previewSource, Operations: unsupported})
	if result.Error == nil || result.Error.Code != compiler.Unsupported || len(result.Edits) != 0 {
		t.Fatalf("unsupported operation result: %+v", result)
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
		{name: "backend-only operation", body: strings.Replace(valid, `"type":"moveControl"`, `"type":"setParent"`, 1), code: compiler.Unsupported},
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
