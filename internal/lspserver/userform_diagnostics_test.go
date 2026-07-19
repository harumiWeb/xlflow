package lspserver

import (
	"context"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	formsintel "github.com/harumiWeb/xlflow/internal/excel/forms/intel"
	vbaintel "github.com/harumiWeb/xlflow/internal/vba/intel"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

func TestUserFormYAMLSemanticDiagnosticsUseSharedValidationAndPreciseRanges(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	source := `schemaVersion: "one"
kind: not-xlflow
basis: runtime
unknownRoot: true
form:
  name: UserForm1
  width: 240
controls:
  - id: label_parent
    type: Label
    list: []
  - id: label_parent
    name: LabelDuplicate
    type: Label
    parentId: missing_parent
  - id: self_frame
    name: SelfFrame
    type: Frame
    parentId: self_frame
  - id: label_child
    name: LabelChild
    type: Label
    parentId: label_parent
  - id: unknown_type
    name: UnknownType
    type: Widget
  - id: mismatch
    name: Mismatch
    type: TextBox
    progId: Forms.Label.1
  - id: custom
    name: Custom
    type: VendorWidget
    progId: Vendor.Widget.1
  - id: frame_a
    name: FrameA
    type: Frame
    parentId: frame_b
  - id: frame_b
    name: FrameB
    type: Frame
    parentId: frame_a
  - id: legacy_frame
    name: LegacyFrame
    type: Frame
    controls: []
  - id: combo
    name: Combo1
    type: ComboBox
    list: []
`
	path := filepath.Join(root, "src", "forms", "specs", "UserForm1.yaml")
	doc, err := s.docs.open(pathToFileURI(path), source)
	if err != nil {
		t.Fatal(err)
	}
	diagnostics := s.documentDiagnostics(context.Background(), doc)
	syntax := formsintel.ParseYAML(source)
	if syntax.ParseError != nil {
		t.Fatal(syntax.ParseError)
	}

	for _, want := range []struct {
		code     string
		severity string
		field    string
		key      bool
	}{
		{"UFV001", "error", "unknownRoot", true},
		{"UFV002", "error", "schemaVersion", false},
		{"UFV003", "error", "kind", false},
		{"UFV004", "error", "controls[0].id", true},
		{"UFV005", "error", "controls[0].list", true},
		{"UFV006", "error", "controls[4].type", false},
		{"UFV007", "error", "controls[1].id", false},
		{"UFV008", "error", "controls[1].parentId", false},
		{"UFV009", "error", "controls[2].parentId", false},
		{"UFV010", "error", "controls[7].parentId", false},
		{"UFV011", "error", "controls[3].parentId", false},
		{"UFV012", "error", "controls[5].progId", false},
		{"UFV013", "warning", "form.width", true},
		{"UFV013", "warning", "controls[9].controls", true},
		{"UFV013", "warning", "controls[10].list", true},
		{"UFV014", "warning", "controls[6].progId", false},
	} {
		field, ok := syntax.Field(want.field)
		if !ok {
			t.Fatalf("syntax field %q was not indexed", want.field)
		}
		rangeWant := field.ValueRange
		if want.key {
			rangeWant = field.KeyRange
		}
		requireUserFormDiagnostic(t, diagnostics, want.code, want.severity, rangeWant)
	}
}

func TestUserFormYAMLDiagnosticsPublishChangesAndClearAfterCorrection(t *testing.T) {
	s, timers, cleanup := newDiagnosticsTestServer(t)
	defer cleanup()
	notifications := &diagnosticNotificationRecorder{}
	ctx := diagnosticTestContext(notifications)
	uri := pathToFileURI(filepath.Join(s.opts.RootDir, "src", "forms", "specs", "UserForm1.yaml"))

	invalid := `schemaVersion: 1
kind: xlflow.userform
basis: designer
form:
  name: UserForm1
controls:
  - id: label
    name: Label1
    type: Label
    list: []
`
	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: protocol.DocumentUri(uri), Version: 1, Text: invalid,
	}}); err != nil {
		t.Fatal(err)
	}
	initial := notifications.waitForCount(t, 1)
	if len(initial[0].Diagnostics) != 1 || initial[0].Diagnostics[0].Code == nil || initial[0].Diagnostics[0].Code.Value != "UFV005" {
		t.Fatalf("open diagnostics = %+v, want UFV005", initial)
	}

	notifications.clear()
	if err := s.didChange(ctx, &protocol.DidChangeTextDocumentParams{
		TextDocument: protocol.VersionedTextDocumentIdentifier{
			TextDocumentIdentifier: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
			Version:                2,
		},
		ContentChanges: []any{protocol.TextDocumentContentChangeEventWhole{Text: "controls:\n  - type: Label\n    broken\n"}},
	}); err != nil {
		t.Fatal(err)
	}
	timers.snapshot()[0].Fire()
	parse := notifications.waitForCount(t, 1)
	if len(parse[0].Diagnostics) != 1 || parse[0].Diagnostics[0].Code == nil || parse[0].Diagnostics[0].Code.Value != "UFY001" {
		t.Fatalf("parse diagnostics = %+v, want UFY001 only", parse)
	}

	notifications.clear()
	valid := `schemaVersion: 1
kind: xlflow.userform
basis: designer
form:
  name: UserForm1
controls: []
`
	if err := s.didChange(ctx, &protocol.DidChangeTextDocumentParams{
		TextDocument: protocol.VersionedTextDocumentIdentifier{
			TextDocumentIdentifier: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
			Version:                3,
		},
		ContentChanges: []any{protocol.TextDocumentContentChangeEventWhole{Text: valid}},
	}); err != nil {
		t.Fatal(err)
	}
	timers.snapshot()[1].Fire()
	fixed := notifications.waitForCount(t, 1)
	if len(fixed[0].Diagnostics) != 0 {
		t.Fatalf("fixed diagnostics = %+v, want empty publish", fixed)
	}
}

func TestUserFormYAMLStructuralDiagnosticUsesNearestCollectionKey(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	source := `schemaVersion: 1
kind: xlflow.userform
basis: designer
form:
  name: UserForm1
controls:
  - not-a-control
`
	doc, err := s.docs.open(pathToFileURI(filepath.Join(root, "src", "forms", "specs", "UserForm1.yaml")), source)
	if err != nil {
		t.Fatal(err)
	}
	diagnostics := s.documentDiagnostics(context.Background(), doc)
	syntax := formsintel.ParseYAML(source)
	controls, ok := syntax.Field("controls")
	if !ok {
		t.Fatal("controls was not indexed")
	}
	requireUserFormDiagnostic(t, diagnostics, "UFV002", "error", controls.KeyRange)
}

func TestUserFormParentPath(t *testing.T) {
	for _, test := range []struct {
		path string
		want string
	}{
		{"controls[0].list[0]", "controls[0].list"},
		{"controls[0].controls[1].id", "controls[0].controls[1]"},
		{"controls[0]", "controls"},
		{"schemaVersion", ""},
	} {
		if got := userFormParentPath(test.path); got != test.want {
			t.Errorf("userFormParentPath(%q) = %q, want %q", test.path, got, test.want)
		}
	}
}

func requireUserFormDiagnostic(t *testing.T, diagnostics []vbaintel.Diagnostic, code, severity string, want formsintel.Range) {
	t.Helper()
	for _, diagnostic := range diagnostics {
		if diagnostic.Code != code || diagnostic.Severity != severity {
			continue
		}
		if diagnostic.Source != "xlflow" {
			t.Fatalf("%s source = %q, want xlflow", code, diagnostic.Source)
		}
		if diagnostic.Range.Start.Line == want.Start.Line && diagnostic.Range.Start.Character == want.Start.Character &&
			diagnostic.Range.End.Line == want.End.Line && diagnostic.Range.End.Character == want.End.Character {
			return
		}
	}
	t.Fatalf("missing %s %s at %#v in diagnostics %#v", code, severity, want, diagnostics)
}
