package lspserver

import (
	"encoding/json"
	"math"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/compiler"
	"github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	formsedit "github.com/harumiWeb/xlflow/internal/vba/userforms/spec/edit"
)

func TestUserFormEditAtomicRootResizePreservesYAMLAndJSONSource(t *testing.T) {
	tests := []struct {
		name       string
		format     string
		source     string
		want       string
		operations []any
	}{
		{
			name:   "yaml",
			format: "yaml",
			source: "# keep header\nschemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: Main\n  caption: \"Keep this caption\" # keep caption comment\n  width: 200.0 # keep width comment\n  height: 150 # keep height comment\ncontrols: []\n",
			want:   "# keep header\nschemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: Main\n  caption: \"Keep this caption\" # keep caption comment\n  width: 280 # keep width comment\n  height: 240 # keep height comment\ncontrols: []\n",
			operations: []any{
				map[string]any{"type": formsedit.SetFormProperty, "field": "width", "value": 280},
				map[string]any{"type": formsedit.SetFormProperty, "field": "height", "value": 240},
			},
		},
		{
			name:   "json",
			format: "json",
			source: "{\n  \"schemaVersion\": 1,\n  \"kind\": \"xlflow.userform\",\n  \"basis\": \"designer\",\n  \"form\": {\n    \"name\": \"Main\",\n    \"caption\": \"Keep this caption\",\n    \"width\": 200.0,\n    \"height\": 150\n  },\n  \"controls\": []\n}\n",
			want:   "{\n  \"schemaVersion\": 1,\n  \"kind\": \"xlflow.userform\",\n  \"basis\": \"designer\",\n  \"form\": {\n    \"name\": \"Main\",\n    \"caption\": \"Keep this caption\",\n    \"width\": 280,\n    \"height\": 240\n  },\n  \"controls\": []\n}\n",
			operations: []any{
				map[string]any{"type": formsedit.SetFormProperty, "field": "height", "value": 240},
				map[string]any{"type": formsedit.SetFormProperty, "field": "width", "value": 280},
			},
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			root := t.TempDir()
			uri := pathToFileURI(filepath.Join(root, "src/forms/specs/Main."+test.format))
			params := parseFormResizeRequest(t, uri, test.source, test.operations)
			server := &Server{opts: Options{RootDir: root, Config: config.Default()}}
			result := server.userFormEdit(params)
			if result.Error != nil || result.Version != 7 || len(result.Edits) != 2 {
				t.Fatalf("atomic root resize failed: %+v", result)
			}

			updated, err := applyLSPTextEdits(test.source, result.Edits)
			if err != nil {
				t.Fatal(err)
			}
			if updated != test.want {
				t.Fatalf("source-preserving resize mismatch:\n got: %q\nwant: %q", updated, test.want)
			}
			document, err := spec.ParseFormSpec(spec.SpecInput{Format: test.format}, []byte(updated))
			if err != nil {
				t.Fatalf("final FormSpec validation failed: %v", err)
			}
			if document.Form.Width == nil || *document.Form.Width != 280 || document.Form.Height == nil || *document.Form.Height != 240 {
				t.Fatalf("final dimensions = (%v, %v), want (280, 240)", document.Form.Width, document.Form.Height)
			}

			invalidParams := parseFormResizeRequest(t, uri, test.source, []any{
				map[string]any{"type": formsedit.SetFormProperty, "field": "width", "value": -1},
				map[string]any{"type": formsedit.SetFormProperty, "field": "height", "value": 240},
			})
			invalidResult := server.userFormEdit(invalidParams)
			if invalidResult.Error == nil || invalidResult.Error.Code != compiler.Invalid || len(invalidResult.Edits) != 0 {
				t.Fatalf("canonical final validation did not reject invalid dimensions: %+v", invalidResult)
			}
		})
	}
}

func TestUserFormEditAtomicRootResizeRejectsMalformedPairs(t *testing.T) {
	uri := "file:///project/src/forms/specs/Main.json"
	source := `{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main","width":200,"height":150},"controls":[]}`
	width := map[string]any{"type": formsedit.SetFormProperty, "field": "width", "value": 280}
	height := map[string]any{"type": formsedit.SetFormProperty, "field": "height", "value": 240}
	controlWidth := map[string]any{"type": formsedit.SetControlProperty, "controlId": "submit", "field": "width", "value": 40}
	resizeControl := map[string]any{"type": formsedit.ResizeControl, "controlId": "submit", "width": 40, "height": 20}
	caption := map[string]any{"type": formsedit.SetFormProperty, "field": "caption", "value": "Title"}
	stringWidth := map[string]any{"type": formsedit.SetFormProperty, "field": "width", "value": "280"}
	nullHeight := map[string]any{"type": formsedit.SetFormProperty, "field": "height", "value": nil}

	for _, test := range []struct {
		name       string
		operations []any
	}{
		{name: "duplicate width", operations: []any{width, width}},
		{name: "caption mixed with width", operations: []any{caption, width}},
		{name: "control property mixed with form property", operations: []any{controlWidth, height}},
		{name: "control resize mixed with form property", operations: []any{resizeControl, width}},
		{name: "string dimension", operations: []any{stringWidth, height}},
		{name: "null dimension", operations: []any{width, nullHeight}},
	} {
		t.Run(test.name, func(t *testing.T) {
			_, requestError := parseUserFormEditParams(marshalFormResizeRequest(t, uri, source, test.operations))
			if requestError == nil || requestError.Code != compiler.Invalid {
				t.Fatalf("malformed pair was accepted: %+v", requestError)
			}
		})
	}

	for _, test := range []struct {
		name       string
		operations []formsedit.Operation
	}{
		{
			name: "missing value presence",
			operations: []formsedit.Operation{
				{Type: formsedit.SetFormProperty, Field: "width", Value: 280, ValuePresent: true},
				{Type: formsedit.SetFormProperty, Field: "height", Value: 240},
			},
		},
		{
			name: "nonfinite NaN",
			operations: []formsedit.Operation{
				{Type: formsedit.SetFormProperty, Field: "width", Value: math.NaN(), ValuePresent: true},
				{Type: formsedit.SetFormProperty, Field: "height", Value: 240, ValuePresent: true},
			},
		},
		{
			name: "nonfinite infinity",
			operations: []formsedit.Operation{
				{Type: formsedit.SetFormProperty, Field: "width", Value: 280, ValuePresent: true},
				{Type: formsedit.SetFormProperty, Field: "height", Value: math.Inf(1), ValuePresent: true},
			},
		},
	} {
		t.Run(test.name, func(t *testing.T) {
			if requestError := validateUserFormEditTransaction(test.operations); requestError == nil || requestError.Code != compiler.Invalid {
				t.Fatalf("invalid typed pair was accepted: %+v", requestError)
			}
		})
	}
}

func TestUserFormEditTransactionRetainsSingleAndGeometryShapes(t *testing.T) {
	move := formsedit.Operation{Type: formsedit.MoveControl, ControlID: "submit", Left: new(1.0), Top: new(2.0)}
	resize := formsedit.Operation{Type: formsedit.ResizeControl, ControlID: "submit", Width: new(20.0), Height: new(10.0)}
	for _, test := range []struct {
		name       string
		operations []formsedit.Operation
	}{
		{name: "single form property", operations: []formsedit.Operation{{Type: formsedit.SetFormProperty, Field: "caption", Value: "Title", ValuePresent: true}}},
		{name: "single move", operations: []formsedit.Operation{move}},
		{name: "single resize", operations: []formsedit.Operation{resize}},
		{name: "move then resize", operations: []formsedit.Operation{move, resize}},
		{name: "resize then move", operations: []formsedit.Operation{resize, move}},
		{
			name: "root resize properties",
			operations: []formsedit.Operation{
				{Type: formsedit.SetFormProperty, Field: "width", Value: json.Number("280.5"), ValuePresent: true},
				{Type: formsedit.SetFormProperty, Field: "height", Value: float64(240), ValuePresent: true},
			},
		},
	} {
		t.Run(test.name, func(t *testing.T) {
			if requestError := validateUserFormEditTransaction(test.operations); requestError != nil {
				t.Fatalf("valid transaction rejected: %+v", requestError)
			}
		})
	}
}

func parseFormResizeRequest(t *testing.T, uri, source string, operations []any) userFormEditParams {
	t.Helper()
	params, requestError := parseUserFormEditParams(marshalFormResizeRequest(t, uri, source, operations))
	if requestError != nil {
		t.Fatalf("valid request was rejected: %+v", requestError)
	}
	params.Version = 7
	return params
}

func marshalFormResizeRequest(t *testing.T, uri, source string, operations []any) []byte {
	t.Helper()
	body, err := json.Marshal(map[string]any{
		"uri": uri, "version": 7, "text": source, "operations": operations,
	})
	if err != nil {
		t.Fatal(err)
	}
	return body
}
