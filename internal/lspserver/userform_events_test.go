package lspserver

import (
	"encoding/json"
	"net"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	formevents "github.com/harumiWeb/xlflow/internal/vba/userforms/events"
	forms "github.com/harumiWeb/xlflow/internal/vba/userforms/spec"
	"github.com/sourcegraph/jsonrpc2"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

const eventFormSource = "schemaVersion: 1\nkind: xlflow.userform\nbasis: designer\nform:\n  name: EventForm\ncontrols: []\n"

func newUserFormEventTestServer(t *testing.T, codeSource string) (*Server, string, string) {
	t.Helper()
	root := t.TempDir()
	config := config.Default()
	config.Src.Forms = "custom/forms"
	config.UserForm.CodeSource = codeSource
	s := &Server{opts: Options{RootDir: root, Config: config}}
	specPath := filepath.Join(root, "custom", "forms", "specs", "EventForm.yaml")
	return s, pathToFileURI(specPath), root
}

func TestUserFormEventsUsesVBEObjectNamesAndResolvesPageParent(t *testing.T) {
	t.Parallel()
	s, uri, _ := newUserFormEventTestServer(t, "sidecar")
	root := s.userFormEvents(userFormEventsParams{URI: uri, Version: 17, Text: eventFormSource})
	if root.Error != "" || root.Version != 17 || root.ObjectName != "UserForm" {
		t.Fatalf("root events: %+v", root)
	}
	if event := findUserFormEvent(root.Events, "Click"); event == nil || event.Signature != "Private Sub UserForm_Click()" || !event.IsDefault {
		t.Fatalf("root Click descriptor: %+v", event)
	}

	controlSource := strings.Replace(eventFormSource, "controls: []", "controls:\n  - {id: label, name: Label1, type: Label}\n", 1)
	control := s.userFormEvents(userFormEventsParams{URI: uri, Version: 18, Text: controlSource, TargetID: "label"})
	if control.Error != "" || control.ObjectName != "Label1" {
		t.Fatalf("control events: %+v", control)
	}
	if event := findUserFormEvent(control.Events, "Click"); event == nil || event.Signature != "Private Sub Label1_Click()" || !event.IsDefault {
		t.Fatalf("control Click descriptor: %+v", event)
	}

	pageSource := strings.Replace(eventFormSource, "controls: []", "controls:\n  - {id: multi, name: MultiPage1, type: MultiPage, selectedIndex: 0}\n  - {id: page, name: Page1, type: Page, parentId: multi}\n", 1)
	page := s.userFormEvents(userFormEventsParams{URI: uri, Version: 19, Text: pageSource, TargetID: "page"})
	if page.Error != "" || page.ObjectName != "MultiPage1" || page.TargetID != "page" {
		t.Fatalf("Page events did not resolve to its parent: %+v", page)
	}
	if event := findUserFormEvent(page.Events, "Change"); event == nil || event.Signature != "Private Sub MultiPage1_Change()" || !event.IsDefault {
		t.Fatalf("MultiPage Change descriptor: %+v", event)
	}
}

func TestUserFormEventsRejectsUnknownAndMismatchedControls(t *testing.T) {
	t.Parallel()
	s, uri, _ := newUserFormEventTestServer(t, "sidecar")
	custom := strings.Replace(eventFormSource, "controls: []", "controls:\n  - {id: custom, name: Vendor1, type: VendorControl, progId: Vendor.Control.1}\n", 1)
	result := s.userFormEvents(userFormEventsParams{URI: uri, Text: custom, TargetID: "custom"})
	if result.Error == "" {
		t.Fatal("custom control was accepted as an event host")
	}

	mismatch := strings.Replace(eventFormSource, "controls: []", "controls:\n  - {id: button, name: Button1, type: TextBox, progId: Forms.CommandButton.1}\n", 1)
	result = s.userFormEvents(userFormEventsParams{URI: uri, Text: mismatch, TargetID: "button"})
	if result.Error == "" {
		t.Fatal("control with a mismatched ProgID was accepted")
	}
}

func TestUserFormEventsRejectsSpecFilenameMismatch(t *testing.T) {
	t.Parallel()
	s, uri, _ := newUserFormEventTestServer(t, "sidecar")
	otherURI := strings.Replace(uri, "EventForm.yaml", "OtherForm.yaml", 1)
	result := s.userFormEvents(userFormEventsParams{URI: otherURI, Version: 1, Text: eventFormSource})
	if result.Error == "" {
		t.Fatal("spec whose basename differs from form.name was accepted")
	}
}

func TestUserFormEventEditCreatesMinimalSidecarWithFixedRootPrefix(t *testing.T) {
	t.Parallel()
	s, uri, root := newUserFormEventTestServer(t, "sidecar")
	codeURI := pathToFileURI(filepath.Join(root, "custom", "forms", "code", "EventForm.bas"))
	result := s.userFormEventEdit(userFormEventEditParams{
		userFormEventsParams: userFormEventsParams{URI: uri, Version: 23, Text: eventFormSource},
		EventName:            "Click", CodeURI: codeURI, CodeVersion: 8, CodeSource: "sidecar", Indent: "    ", Newline: "\r\n",
	})
	if result.Error != "" || !result.Created || result.Version != 23 || result.CodeVersion != 8 || len(result.Edits) != 1 {
		t.Fatalf("new sidecar handler: %+v", result)
	}
	want := "Option Explicit\r\n\r\nPrivate Sub UserForm_Click()\r\n    \r\nEnd Sub\r\n"
	if result.Edits[0].NewText != want {
		t.Fatalf("insertion = %q, want %q", result.Edits[0].NewText, want)
	}
	if result.Edits[0].Range.Start != (protocol.Position{}) || result.Edits[0].Range.End != (protocol.Position{}) {
		t.Fatalf("empty file edit range = %+v", result.Edits[0].Range)
	}
	if result.Selection.Start != (protocol.Position{Line: 3, Character: 4}) || result.Selection.End != result.Selection.Start {
		t.Fatalf("body selection = %+v, want line 3 character 4", result.Selection)
	}
}

func TestUserFormEventEditUsesControlNameAndUTF16EOF(t *testing.T) {
	t.Parallel()
	s, uri, root := newUserFormEventTestServer(t, "sidecar")
	text := "Option Explicit\n' 😀"
	source := strings.Replace(eventFormSource, "controls: []", "controls:\n  - {id: text, name: TextBox1, type: TextBox}\n", 1)
	result := s.userFormEventEdit(userFormEventEditParams{
		userFormEventsParams: userFormEventsParams{URI: uri, Version: 1, Text: source, TargetID: "text"},
		EventName:            "Change", CodeURI: pathToFileURI(filepath.Join(root, "custom", "forms", "code", "EventForm.bas")),
		CodeVersion: 2, CodeText: text, CodeSource: "sidecar", Newline: "\n", Indent: "\t",
	})
	if result.Error != "" || !result.Created || len(result.Edits) != 1 {
		t.Fatalf("control handler: %+v", result)
	}
	if !strings.Contains(result.Edits[0].NewText, "Private Sub TextBox1_Change()\n\t\nEnd Sub") {
		t.Fatalf("control handler signature/body: %q", result.Edits[0].NewText)
	}
	if result.Edits[0].Range.Start != (protocol.Position{Line: 1, Character: 4}) {
		t.Fatalf("EOF range = %+v, want UTF-16 character 4", result.Edits[0].Range.Start)
	}
}

func TestUserFormEventEditNavigatesExistingProcedureWithoutEditingSignature(t *testing.T) {
	t.Parallel()
	s, uri, root := newUserFormEventTestServer(t, "sidecar")
	code := "' UserForm_Click is only a comment\nPrivate Sub userform_CLICK()\n    Dim label As String\n    label = \"UserForm_Click\"\nEnd Sub\n"
	original := code
	result := s.userFormEventEdit(userFormEventEditParams{
		userFormEventsParams: userFormEventsParams{URI: uri, Version: 3, Text: eventFormSource},
		EventName:            "Click", CodeURI: pathToFileURI(filepath.Join(root, "custom", "forms", "code", "EventForm.bas")),
		CodeVersion: 4, CodeText: code, CodeSource: "sidecar", Newline: "\n", Indent: "    ",
	})
	if result.Error != "" || result.Created || len(result.Edits) != 0 || code != original {
		t.Fatalf("existing handler edit result=%+v source changed=%t", result, code != original)
	}
	if result.Selection.Start.Line < 2 || result.Selection.End != result.Selection.Start {
		t.Fatalf("existing procedure body selection = %+v", result.Selection)
	}
}

func TestUserFormEventEditRejectsSyntaxDuplicateConditionalAndModuleDeclarationCollisions(t *testing.T) {
	t.Parallel()
	s, uri, root := newUserFormEventTestServer(t, "sidecar")
	codeURI := pathToFileURI(filepath.Join(root, "custom", "forms", "code", "EventForm.bas"))
	cases := []struct {
		name string
		code string
	}{
		{name: "syntax error", code: "Private Sub Broken(\n"},
		{name: "duplicate procedures", code: "Private Sub UserForm_Click()\nEnd Sub\nPrivate Sub userform_click()\nEnd Sub\n"},
		{name: "conditional procedures", code: "#If VBA7 Then\nPrivate Sub UserForm_Click()\nEnd Sub\n#Else\nPrivate Sub UserForm_Click()\nEnd Sub\n#End If\n"},
		{name: "module variable", code: "Private UserForm_Click As Long\n"},
		{name: "module constant", code: "Private Const UserForm_Click As Long = 1\n"},
		{name: "module declare", code: "Private Declare PtrSafe Function UserForm_Click Lib \"kernel32\" () As Long\n"},
		{name: "module event", code: "Public Event UserForm_Click()\n"},
	}
	for _, test := range cases {
		t.Run(test.name, func(t *testing.T) {
			result := s.userFormEventEdit(userFormEventEditParams{
				userFormEventsParams: userFormEventsParams{URI: uri, Version: 1, Text: eventFormSource},
				EventName:            "Click", CodeURI: codeURI, CodeVersion: 1, CodeText: test.code,
				CodeSource: "sidecar", Newline: "\n", Indent: "",
			})
			if result.Error == "" || result.Created || len(result.Edits) != 0 {
				t.Fatalf("ambiguous code was accepted: %+v", result)
			}
		})
	}
}

func TestUserFormEventEditPreservesFRMHeaderAndChecksVBName(t *testing.T) {
	t.Parallel()
	s, uri, root := newUserFormEventTestServer(t, "frm")
	codeURI := pathToFileURI(filepath.Join(root, "custom", "forms", "EventForm.frm"))
	frm := "VERSION 5.00\r\nBegin VB.UserForm EventForm\r\nEnd\r\nAttribute VB_Name = \"EventForm\"\r\nAttribute VB_GlobalNameSpace = False\r\n\r\nOption Explicit\r\n"
	result := s.userFormEventEdit(userFormEventEditParams{
		userFormEventsParams: userFormEventsParams{URI: uri, Version: 10, Text: eventFormSource},
		EventName:            "Click", CodeURI: codeURI, CodeVersion: 11, CodeText: frm,
		CodeSource: "frm", Newline: "\r\n", Indent: "  ",
	})
	if result.Error != "" || !result.Created || len(result.Edits) != 1 {
		t.Fatalf("FRM edit: %+v", result)
	}
	if result.Edits[0].Range.Start != sourcePositionAtUTF16([]byte(frm), len(frm)) {
		t.Fatalf("FRM edit was not appended at EOF: %+v", result.Edits[0].Range)
	}
	if !strings.HasPrefix(result.Edits[0].NewText, "\r\nPrivate Sub UserForm_Click()") {
		t.Fatalf("FRM append = %q", result.Edits[0].NewText)
	}
	bad := strings.Replace(frm, `VB_Name = "EventForm"`, `VB_Name = "OtherForm"`, 1)
	result = s.userFormEventEdit(userFormEventEditParams{
		userFormEventsParams: userFormEventsParams{URI: uri, Version: 10, Text: eventFormSource},
		EventName:            "Click", CodeURI: codeURI, CodeVersion: 11, CodeText: bad,
		CodeSource: "frm", Newline: "\r\n", Indent: "  ",
	})
	if result.Error == "" || len(result.Edits) != 0 {
		t.Fatalf("mismatched VB_Name was accepted: %+v", result)
	}
	duplicate := frm + "Attribute VB_Name = \"EventForm\"\r\n"
	result = s.userFormEventEdit(userFormEventEditParams{
		userFormEventsParams: userFormEventsParams{URI: uri, Version: 10, Text: eventFormSource},
		EventName:            "Click", CodeURI: codeURI, CodeVersion: 11, CodeText: duplicate,
		CodeSource: "frm", Newline: "\r\n", Indent: "  ",
	})
	if result.Error == "" || len(result.Edits) != 0 {
		t.Fatalf("duplicate VB_Name was accepted: %+v", result)
	}
}

func TestUserFormEventEditRejectsCodeURIAuthorityMismatch(t *testing.T) {
	t.Parallel()
	s, uri, root := newUserFormEventTestServer(t, "sidecar")
	params := userFormEventEditParams{
		userFormEventsParams: userFormEventsParams{URI: uri, Version: 1, Text: eventFormSource},
		EventName:            "Click", CodeVersion: 1, CodeSource: "sidecar", Newline: "\n", Indent: "",
	}
	params.CodeURI = pathToFileURI(filepath.Join(root, "custom", "forms", "OtherForm.bas"))
	if result := s.userFormEventEdit(params); result.Error == "" {
		t.Fatal("code URI outside the sidecar authority was accepted")
	}
	params.CodeURI = pathToFileURI(filepath.Join(root, "custom", "forms", "code", "EventForm.bas"))
	params.CodeSource = "frm"
	if result := s.userFormEventEdit(params); result.Error == "" {
		t.Fatal("request source mode differing from project configuration was accepted")
	}
}

func TestUserFormEventEditRejectsLinkedCodeDirectories(t *testing.T) {
	t.Parallel()
	for _, mode := range []string{"sidecar", "frm"} {
		t.Run(mode, func(t *testing.T) {
			s, uri, root := newUserFormEventTestServer(t, mode)
			formsRoot := filepath.Join(root, "custom", "forms")
			if err := os.MkdirAll(formsRoot, 0o755); err != nil {
				t.Fatal(err)
			}
			external := t.TempDir()
			linked := filepath.Join(formsRoot, "code")
			filename := "EventForm.bas"
			if mode == "frm" {
				linked = filepath.Join(formsRoot, "nested")
				filename = "EventForm.frm"
			}
			const original = "' external file must stay unchanged\n"
			if err := os.WriteFile(filepath.Join(external, filename), []byte(original), 0o644); err != nil {
				t.Fatal(err)
			}
			if runtime.GOOS == "windows" {
				if out, err := exec.Command("cmd", "/c", "mklink", "/J", linked, external).CombinedOutput(); err != nil {
					t.Fatalf("create junction: %v: %s", err, out)
				}
			} else if err := os.Symlink(external, linked); err != nil {
				t.Fatal(err)
			}
			if _, err := os.Stat(filepath.Join(linked, filename)); err != nil {
				t.Fatalf("junction fixture must expose the external code file: %v", err)
			}
			codeText := "Option Explicit\n"
			if mode == "frm" {
				codeText = "VERSION 5.00\nBegin VB.UserForm EventForm\nEnd\nAttribute VB_Name = \"EventForm\"\nOption Explicit\n"
			}
			result := s.userFormEventEdit(userFormEventEditParams{
				userFormEventsParams: userFormEventsParams{URI: uri, Version: 1, Text: eventFormSource},
				EventName:            "Click", CodeURI: pathToFileURI(filepath.Join(linked, filename)),
				CodeVersion: 1, CodeSource: mode, Newline: "\n", CodeText: codeText,
			})
			if result.Error == "" || len(result.Edits) != 0 {
				t.Fatalf("linked code directory was accepted: %+v", result)
			}
			got, err := os.ReadFile(filepath.Join(external, filename))
			if err != nil || string(got) != original {
				t.Fatalf("external file changed: %q, %v", got, err)
			}
		})
	}
}

func TestUserFormEventRequestIsScopedToTargetNotOtherInvalidProjectModules(t *testing.T) {
	t.Parallel()
	s, uri, root := newUserFormEventTestServer(t, "sidecar")
	invalidModule := filepath.Join(root, "src", "modules", "Broken.bas")
	if err := os.MkdirAll(filepath.Dir(invalidModule), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(invalidModule, []byte("Private Sub Broken(\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	result := s.userFormEventEdit(userFormEventEditParams{
		userFormEventsParams: userFormEventsParams{URI: uri, Version: 1, Text: eventFormSource},
		EventName:            "Click", CodeURI: pathToFileURI(filepath.Join(root, "custom", "forms", "code", "EventForm.bas")),
		CodeVersion: 1, CodeSource: "sidecar", Newline: "\n", Indent: "",
	})
	if result.Error != "" || !result.Created {
		t.Fatalf("unrelated invalid project module blocked target edit: %+v", result)
	}
}

func TestUserFormEventParamsAreStrict(t *testing.T) {
	t.Parallel()
	validList := `{"uri":"file:///p/src/forms/specs/EventForm.yaml","version":1,"text":""}`
	if _, err := parseUserFormEventsParams([]byte(validList)); err != nil {
		t.Fatalf("valid list params: %v", err)
	}
	for _, raw := range []string{
		`{"uri":"file:///p/src/forms/specs/EventForm.yaml","version":1,"text":"","unexpected":true}`,
		`{"uri":"file:///p/src/forms/specs/EventForm.yaml","version":1,"version":2,"text":""}`,
		`{"uri":"file:///p/src/forms/specs/EventForm.yaml","version":-1,"text":""}`,
		`{"uri":"file:///p/src/forms/specs/EventForm.yaml","version":1,"text":"","targetId":null}`,
	} {
		if _, err := parseUserFormEventsParams([]byte(raw)); err == nil {
			t.Errorf("accepted malformed list params %s", raw)
		}
	}
	validEdit := `{"uri":"file:///p/src/forms/specs/EventForm.yaml","version":1,"text":"","eventName":"Click","codeURI":"file:///p/src/forms/code/EventForm.bas","codeVersion":2,"codeText":"","indent":"","newline":"\n","codeSource":"sidecar"}`
	if _, err := parseUserFormEventEditParams([]byte(validEdit)); err != nil {
		t.Fatalf("valid edit params: %v", err)
	}
	if _, err := parseUserFormEventEditParams([]byte(strings.Replace(validEdit, `"codeSource":"sidecar"`, `"codeSource":"sidecar","extra":1`, 1))); err == nil {
		t.Fatal("accepted unknown edit parameter")
	}
}

func TestUserFormEventsRPCAdvertisesCapabilityAndReturnsContract(t *testing.T) {
	t.Parallel()
	ctx := t.Context()
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
	if initialized.Capabilities.Experimental["userFormEvents"] != float64(1) {
		t.Fatalf("UserForm event capability: %#v", initialized.Capabilities.Experimental)
	}
	if err := clientConn.Notify(ctx, "initialized", map[string]any{}); err != nil {
		t.Fatal(err)
	}
	var result userFormEventsResult
	params := userFormEventsParams{URI: pathToFileURI(filepath.Join(root, "src", "forms", "specs", "EventForm.yaml")), Version: 5, Text: eventFormSource}
	if err := clientConn.Call(ctx, userFormEventsMethod, params, &result); err != nil {
		t.Fatal(err)
	}
	if result.Error != "" || result.Version != 5 || result.ObjectName != "UserForm" || len(result.Events) != 22 {
		t.Fatalf("events RPC result: %+v", result)
	}
}

func findUserFormEvent(descriptors []formevents.Descriptor, name string) *formevents.Descriptor {
	for index := range descriptors {
		if strings.EqualFold(descriptors[index].Name, name) {
			return &descriptors[index]
		}
	}
	return nil
}

func TestUserFormEventEditResultJSONUsesWireNames(t *testing.T) {
	t.Parallel()
	encoded, err := json.Marshal(userFormEventEditResult{
		Version: 1, CodeVersion: 2, Edits: []protocol.TextEdit{}, Selection: protocol.Range{}, Created: false,
	})
	if err != nil {
		t.Fatal(err)
	}
	for _, key := range []string{`"version"`, `"codeVersion"`, `"edits"`, `"selection"`, `"created"`} {
		if !strings.Contains(string(encoded), key) {
			t.Errorf("wire result is missing %s: %s", key, encoded)
		}
	}
}

func TestUserFormEventTargetRootObjectNameDoesNotFollowComponentName(t *testing.T) {
	t.Parallel()
	form, err := forms.ParseFormSpec(forms.SpecInput{Format: "yaml"}, []byte(eventFormSource))
	if err != nil {
		t.Fatal(err)
	}
	target, err := formevents.ResolveTarget(form, "")
	if err != nil || target.ObjectName != "UserForm" {
		t.Fatalf("root target = %+v, %v", target, err)
	}
	got, found := formevents.ForType(target.Type, target.ObjectName)
	click := findUserFormEvent(got, "Click")
	if !found || click == nil || click.Signature != "Private Sub UserForm_Click()" {
		t.Fatalf("root event signature = %+v, found=%t", got, found)
	}
}
