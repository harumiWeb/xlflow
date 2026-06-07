package output

import (
	"bytes"
	"encoding/json"
	"errors"
	"strings"
	"testing"
)

func TestExitCode(t *testing.T) {
	if got := ExitCode(nil); got != ExitSuccess {
		t.Fatalf("nil error exit code = %d, want %d", got, ExitSuccess)
	}
	if got := ExitCode(WithExitCode(ExitEnvironment, errors.New("boom"))); got != ExitEnvironment {
		t.Fatalf("classified exit code = %d, want %d", got, ExitEnvironment)
	}
	if got := ExitCode(errors.New("boom")); got != ExitConfig {
		t.Fatalf("default exit code = %d, want %d", got, ExitConfig)
	}
}

func TestWriteJSONEnvelope(t *testing.T) {
	env := New("lint")
	env.Issues = []string{"x"}
	var buf bytes.Buffer
	if err := Write(&buf, env, true); err != nil {
		t.Fatal(err)
	}
	var decoded Envelope
	if err := json.Unmarshal(buf.Bytes(), &decoded); err != nil {
		t.Fatal(err)
	}
	if decoded.Status != StatusOK || decoded.Command != "lint" || decoded.Error != nil {
		t.Fatalf("decoded envelope = %+v", decoded)
	}
}

func TestWriteJSONEnvelopeIncludesTests(t *testing.T) {
	env := New("test")
	env.Tests = []map[string]any{
		{"name": "TestCreateReport", "module": "ReportTests", "status": "passed", "duration_ms": 12},
	}
	var buf bytes.Buffer
	if err := Write(&buf, env, true); err != nil {
		t.Fatal(err)
	}
	var decoded map[string]any
	if err := json.Unmarshal(buf.Bytes(), &decoded); err != nil {
		t.Fatal(err)
	}
	tests, ok := decoded["tests"].([]any)
	if !ok || len(tests) != 1 {
		t.Fatalf("expected one test result in JSON envelope: %s", buf.String())
	}
}

func TestWriteJSONEnvelopeIncludesDiff(t *testing.T) {
	env := New("diff")
	env.Diff = map[string]any{
		"summary": map[string]any{"total_diffs": 1},
	}
	var buf bytes.Buffer
	if err := Write(&buf, env, true); err != nil {
		t.Fatal(err)
	}
	var decoded map[string]any
	if err := json.Unmarshal(buf.Bytes(), &decoded); err != nil {
		t.Fatal(err)
	}
	if _, ok := decoded["diff"].(map[string]any); !ok {
		t.Fatalf("expected diff result in JSON envelope: %s", buf.String())
	}
}

func TestWriteJSONEnvelopeIncludesAnalysisCheckAndRunDiagnostic(t *testing.T) {
	env := New("check")
	env.Analysis = []map[string]any{{"code": "VBA101"}}
	env.Check = map[string]any{"analyze": map[string]any{"status": "failed", "count": 1}}
	env.RunDiagnostic = map[string]any{"likely_cause": "missing Set"}
	var buf bytes.Buffer
	if err := Write(&buf, env, true); err != nil {
		t.Fatal(err)
	}
	var decoded map[string]any
	if err := json.Unmarshal(buf.Bytes(), &decoded); err != nil {
		t.Fatal(err)
	}
	for _, key := range []string{"analysis", "check", "run_diagnostic"} {
		if _, ok := decoded[key]; !ok {
			t.Fatalf("expected %s in JSON envelope: %s", key, buf.String())
		}
	}
}

func TestWriteJSONEnvelopeIncludesPushDiagnostic(t *testing.T) {
	env := New("push")
	env.PushDiagnostic = map[string]any{
		"kind": "compile",
		"location": map[string]any{
			"source_path": "src/modules/Main.bas",
			"line":        6,
			"text":        "  x =",
		},
	}
	var buf bytes.Buffer
	if err := Write(&buf, env, true); err != nil {
		t.Fatal(err)
	}
	var decoded map[string]any
	if err := json.Unmarshal(buf.Bytes(), &decoded); err != nil {
		t.Fatal(err)
	}
	if _, ok := decoded["push_diagnostic"]; !ok {
		t.Fatalf("expected push_diagnostic in JSON envelope: %s", buf.String())
	}
}

func TestWriteWithOptionsCompactsRunJSONByDefault(t *testing.T) {
	env := Failure("run", Error{Code: "macro_failed", Message: "boom", Number: 9, Phase: "invoke_macro"})
	env.Logs = []string{"verbose log"}
	env.Workbook = map[string]any{"path": `C:\temp\Book.xlsm`}
	env.Bridge = map[string]any{"host": "dotnet"}
	env.Runtime = map[string]any{"mode": "headless"}
	env.Target = map[string]any{"kind": "live_session", "path": `C:\temp\Book.xlsm`}
	env.Session = map[string]any{
		"active":          true,
		"mode":            "explicit",
		"dirty":           true,
		"save_required":   true,
		"source_of_truth": "live_workbook",
		"workbook_path":   `C:\temp\Book.xlsm`,
		"extra":           "hidden",
	}
	env.Warnings = []map[string]any{{"code": "save_required", "message": "save it"}}
	env.Macro = map[string]any{"name": "Main.Run", "duration_ms": 42, "arguments": []any{"x"}, "error": map[string]any{"message": "dup"}}
	env.RunDiagnostic = map[string]any{
		"kind":             "runtime",
		"suggestion":       "Inspect src/modules/Main.bas:12.",
		"location":         map[string]any{"source_path": "src/modules/Main.bas", "component": "Main", "component_type": "module", "procedure": "Run", "line": 12, "end_line": 12, "text": "    Debug.Print foo", "confidence": "high", "method": "vbe.selection"},
		"dialogs":          []map[string]any{{"title": "Microsoft Visual Basic"}},
		"location_capture": map[string]any{"attempts": []map[string]any{{"timing": "before_dialog_action"}}},
		"worker":           map[string]any{"pid": 1234},
	}

	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{JSON: true}); err != nil {
		t.Fatal(err)
	}
	var decoded map[string]any
	if err := json.Unmarshal(buf.Bytes(), &decoded); err != nil {
		t.Fatal(err)
	}
	for _, key := range []string{"workbook", "bridge", "runtime", "logs", "run_diagnostic"} {
		if _, ok := decoded[key]; ok {
			t.Fatalf("did not expect %s in compact run JSON: %s", key, buf.String())
		}
	}
	if got := decoded["suggestion"]; got != "Inspect src/modules/Main.bas:12." {
		t.Fatalf("suggestion = %#v, want promoted suggestion", got)
	}
	location, ok := decoded["location"].(map[string]any)
	if !ok || location["source_path"] != "src/modules/Main.bas" {
		t.Fatalf("expected promoted location in compact run JSON: %s", buf.String())
	}
	macro, ok := decoded["macro"].(map[string]any)
	if !ok || macro["name"] != "Main.Run" || macro["duration_ms"] != float64(42) {
		t.Fatalf("unexpected compact macro payload: %s", buf.String())
	}
	if _, ok := macro["arguments"]; ok {
		t.Fatalf("did not expect macro arguments in compact run JSON: %s", buf.String())
	}
	session, ok := decoded["session"].(map[string]any)
	if !ok || session["workbook_path"] != `C:\temp\Book.xlsm` {
		t.Fatalf("unexpected compact session payload: %s", buf.String())
	}
	if _, ok := session["extra"]; ok {
		t.Fatalf("did not expect extra session fields in compact run JSON: %s", buf.String())
	}
}

func TestWriteWithOptionsIncludesVerboseRunJSONDiagnostics(t *testing.T) {
	env := Failure("run", Error{Code: "macro_failed", Message: "boom", Number: 9, Phase: "invoke_macro"})
	env.Logs = []string{"verbose log"}
	env.Workbook = map[string]any{"path": `C:\temp\Book.xlsm`}
	env.Bridge = map[string]any{"host": "dotnet"}
	env.Runtime = map[string]any{"mode": "headless"}
	env.Macro = map[string]any{"name": "Main.Run", "duration_ms": 42, "arguments": []any{"x"}}
	env.RunDiagnostic = map[string]any{
		"kind":    "runtime",
		"dialog":  map[string]any{"title": "Microsoft Visual Basic"},
		"dialogs": []map[string]any{{"title": "Microsoft Visual Basic"}},
	}

	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{JSON: true, Verbose: true}); err != nil {
		t.Fatal(err)
	}
	var decoded map[string]any
	if err := json.Unmarshal(buf.Bytes(), &decoded); err != nil {
		t.Fatal(err)
	}
	if _, ok := decoded["workbook"].(map[string]any); !ok {
		t.Fatalf("expected workbook in verbose run JSON: %s", buf.String())
	}
	if _, ok := decoded["bridge"].(map[string]any); !ok {
		t.Fatalf("expected bridge in verbose run JSON: %s", buf.String())
	}
	diag, ok := decoded["run_diagnostic"].(map[string]any)
	if !ok {
		t.Fatalf("expected run_diagnostic in verbose run JSON: %s", buf.String())
	}
	if _, ok := diag["dialogs"]; !ok {
		t.Fatalf("expected dialogs in verbose run JSON: %s", buf.String())
	}
	if _, ok := diag["dialog"]; ok {
		t.Fatalf("did not expect duplicate dialog field in verbose run JSON: %s", buf.String())
	}
}

func TestWriteWithOptionsKeepsRunPreflightDiagnosticsByDefault(t *testing.T) {
	env := Failure("run", Error{Code: "source_preflight_failed", Message: "preflight failed", Phase: "preflight"})
	env.Issues = []map[string]any{{"code": "VB001", "file": "src/modules/Main.bas", "line": 10}}
	env.Analysis = []map[string]any{{"code": "VBA101", "file": "src/modules/Main.bas", "line": 12}}
	env.GUIBoundaries = []map[string]any{{"file": "src/modules/Main.bas", "line": 14, "kind": "modal_dialog"}}
	env.Workbook = map[string]any{"path": `C:\temp\Book.xlsm`}
	env.Runtime = map[string]any{"mode": "headless"}
	env.Logs = []string{"blocked before Excel automation"}

	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{JSON: true}); err != nil {
		t.Fatal(err)
	}
	var decoded map[string]any
	if err := json.Unmarshal(buf.Bytes(), &decoded); err != nil {
		t.Fatal(err)
	}
	for _, key := range []string{"issues", "analysis", "gui_boundaries"} {
		if _, ok := decoded[key]; !ok {
			t.Fatalf("expected %s in compact run preflight JSON: %s", key, buf.String())
		}
	}
	for _, key := range []string{"workbook", "runtime", "logs"} {
		if _, ok := decoded[key]; ok {
			t.Fatalf("did not expect %s in compact run preflight JSON: %s", key, buf.String())
		}
	}
}

func TestPushHumanOutputRendersDiagnosticSourcePathAndText(t *testing.T) {
	env := Failure("push", Error{Code: "vba_compile_failed", Message: "Compile error", Phase: "compile_vba"})
	env.Workbook = map[string]any{"path": "build/Book.xlsm"}
	env.PushDiagnostic = map[string]any{
		"kind": "compile",
		"location": map[string]any{
			"component":   "Main",
			"procedure":   "CompileError",
			"source_path": "src/modules/Main.bas",
			"line":        6,
			"column":      3,
			"text":        "  x =",
		},
	}

	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	text := buf.String()
	for _, want := range []string{"Diagnostic", "src/modules/Main.bas", "line 6", "column 3", "x ="} {
		if !strings.Contains(text, want) {
			t.Fatalf("human output missing %q:\n%s", want, text)
		}
	}
}

func TestRunHumanOutputRendersDiagnosticSourcePathAndText(t *testing.T) {
	env := Failure("run", Error{Code: "vba_compile_failed", Message: "Compile error", Phase: "compile_vba"})
	env.Macro = map[string]any{"name": "Main.Run", "duration_ms": 0}
	env.RunDiagnostic = map[string]any{
		"kind": "compile",
		"location": map[string]any{
			"component":   "Main",
			"procedure":   "Run",
			"source_path": "src/modules/Main.bas",
			"line":        12,
			"column":      5,
			"text":        "    Debug.Print foo",
		},
	}

	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	text := buf.String()
	for _, want := range []string{"src/modules/Main.bas", "line 12", "column 5", "Debug.Print foo"} {
		if !strings.Contains(text, want) {
			t.Fatalf("human output missing %q:\n%s", want, text)
		}
	}
}

func TestWriteJSONEnvelopeIncludesExportImageFields(t *testing.T) {
	env := New("export-image")
	env.Target = map[string]any{"kind": "live_session", "path": "build/Book.xlsm", "sheet": "QR", "range": "A1:AE31"}
	env.Session = map[string]any{"active": true, "workbook_path": "build/Book.xlsm", "dirty": true, "save_required": true}
	env.Output = map[string]any{"path": ".xlflow/artifacts/images/Book/qr.png", "format": "png", "default": true}
	env.Warnings = []map[string]any{{"code": "temporary_object_cleanup_failed", "message": "cleanup failed"}}
	env.Hints = []map[string]any{{"code": "next_step", "message": "Run `xlflow save --session`."}}
	var buf bytes.Buffer
	if err := Write(&buf, env, true); err != nil {
		t.Fatal(err)
	}
	var decoded map[string]any
	if err := json.Unmarshal(buf.Bytes(), &decoded); err != nil {
		t.Fatal(err)
	}
	for _, key := range []string{"target", "session", "output", "warnings", "hints"} {
		if _, ok := decoded[key]; !ok {
			t.Fatalf("expected %s in JSON envelope: %s", key, buf.String())
		}
	}
}

func TestWriteJSONEnvelopeIncludesFormSnapshotFields(t *testing.T) {
	env := New("form snapshot")
	env.Target = map[string]any{"kind": "live_session", "path": "build/Book.xlsm"}
	env.Session = map[string]any{"active": true, "workbook_path": "build/Book.xlsm", "dirty": true, "save_required": true}
	env.Forms = map[string]any{"name": "UserForm1", "basis": "designer", "coordinate_system": "parent-relative", "control_count": 3}
	env.Output = map[string]any{"path": "artifacts/UserForm1.form.yaml", "format": "yaml"}
	env.Warnings = []map[string]any{{"code": "save_required", "message": "Run `xlflow save --session` to persist workbook changes."}}
	var buf bytes.Buffer
	if err := Write(&buf, env, true); err != nil {
		t.Fatal(err)
	}
	var decoded map[string]any
	if err := json.Unmarshal(buf.Bytes(), &decoded); err != nil {
		t.Fatal(err)
	}
	for _, key := range []string{"target", "session", "forms", "output", "warnings"} {
		if _, ok := decoded[key]; !ok {
			t.Fatalf("expected %s in JSON envelope: %s", key, buf.String())
		}
	}
}

func TestWriteJSONEnvelopeIncludesEditFields(t *testing.T) {
	env := New("edit")
	env.Target = map[string]any{"kind": "live_session", "path": "build/Book.xlsm"}
	env.Session = map[string]any{"active": true, "workbook_path": "build/Book.xlsm", "dirty": true, "save_required": true}
	env.Edit = map[string]any{
		"kind":  "cell",
		"sheet": "Input",
		"cell":  "B2",
		"mutation": map[string]any{
			"value": map[string]any{"before": "", "after": "ABC123"},
		},
	}
	var buf bytes.Buffer
	if err := Write(&buf, env, true); err != nil {
		t.Fatal(err)
	}
	var decoded map[string]any
	if err := json.Unmarshal(buf.Bytes(), &decoded); err != nil {
		t.Fatal(err)
	}
	for _, key := range []string{"target", "session", "edit"} {
		if _, ok := decoded[key]; !ok {
			t.Fatalf("expected %s in JSON envelope: %s", key, buf.String())
		}
	}
}

func TestWriteJSONEnvelopeIncludesBackupAndRollbackFields(t *testing.T) {
	env := New("rollback")
	env.Backups = []map[string]any{{"id": "20260518-100000-push", "path": ".xlflow/backups/20260518-100000-push/Book.xlsm"}}
	env.Rollback = map[string]any{
		"restored_from": map[string]any{"id": "20260518-100000-push"},
		"safety_backup": map[string]any{"id": "20260518-110000-pre-rollback"},
	}
	var buf bytes.Buffer
	if err := Write(&buf, env, true); err != nil {
		t.Fatal(err)
	}
	var decoded map[string]any
	if err := json.Unmarshal(buf.Bytes(), &decoded); err != nil {
		t.Fatal(err)
	}
	for _, key := range []string{"backups", "rollback"} {
		if _, ok := decoded[key]; !ok {
			t.Fatalf("expected %s in JSON envelope: %s", key, buf.String())
		}
	}
}

func TestWriteWithOptionsRendersBackupListSummary(t *testing.T) {
	env := New("backup list")
	env.Backups = []map[string]any{
		{"id": "20260518-100000-push", "created_at": "2026-05-18T10:00:00+09:00", "reason": "before-push", "workbook": "build/Book.xlsm", "path": ".xlflow/backups/20260518-100000-push/Book.xlsm"},
		{"id": "20260518-110000-pre-rollback", "created_at": "2026-05-18T11:00:00+09:00", "reason": "pre-rollback", "workbook": "build/Book.xlsm", "path": ".xlflow/backups/20260518-110000-pre-rollback/Book.xlsm"},
	}
	env.Warnings = []map[string]any{{"code": "stale_index", "message": "refresh backup metadata"}}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Backups:", "2", "20260518-100000-push", "before-push", "20260518-110000-pre-rollback", "Warnings:", "refresh backup metadata"} {
		if !strings.Contains(got, want) {
			t.Fatalf("backup list output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersRollbackSummary(t *testing.T) {
	env := New("rollback")
	env.Target = map[string]any{"kind": "file", "path": "build/Book.xlsm"}
	env.Rollback = map[string]any{
		"restored_from": map[string]any{
			"id":         "20260518-100000-push",
			"path":       ".xlflow/backups/20260518-100000-push/Book.xlsm",
			"reason":     "before-push",
			"created_at": "2026-05-18T10:00:00+09:00",
		},
		"safety_backup": map[string]any{
			"id":   "20260518-110000-pre-rollback",
			"path": ".xlflow/backups/20260518-110000-pre-rollback/Book.xlsm",
		},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Target:", "build/Book.xlsm", "Backup ID:", "20260518-100000-push", "Safety backup:", "20260518-110000-pre-rollback"} {
		if !strings.Contains(got, want) {
			t.Fatalf("rollback output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsInspectMarkdownUsesUnnamedPlaceholder(t *testing.T) {
	env := New("inspect")
	payload := map[string]any{
		"target": "form",
		"forms": map[string]any{
			"designer": map[string]any{
				"name":  "UserForm1",
				"basis": "designer",
				"controls": []map[string]any{
					{"name": "", "type": "Label", "caption": "Hello"},
				},
			},
		},
	}
	text := renderer{}.renderInspectMarkdown(env, payload)
	if !strings.Contains(text, "<unnamed>") {
		t.Fatalf("expected unnamed placeholder in markdown output: %s", text)
	}
}

func TestRenderInspectFormShowsUnavailableWhenBothPayloadMissing(t *testing.T) {
	env := New("inspect")
	env.Warnings = []map[string]any{{"code": "runtime_unavailable", "message": "runtime snapshot unavailable"}}
	text := renderer{}.renderInspectForm(env, map[string]any{"target": "form"})
	if !strings.Contains(text, "unavailable") {
		t.Fatalf("expected unavailable forms summary: %s", text)
	}
	if !strings.Contains(text, "runtime snapshot unavailable") {
		t.Fatalf("expected warnings to be preserved: %s", text)
	}
}

func TestWriteWithOptionsInspectJSONIncludesTargetStateAndWarnings(t *testing.T) {
	env := New("inspect")
	env.Target = map[string]any{"kind": "file", "path": "build/Book.xlsm", "description": "Saved workbook file on disk"}
	env.Session = map[string]any{"active": true, "workbook_path": "build/Book.xlsm", "dirty": true, "save_required": true}
	env.Warnings = []map[string]any{{"code": "live_session_dirty", "message": "stale"}}
	env.Hints = []map[string]any{{"code": "next_step", "message": "save first"}}
	env.Inspect = map[string]any{
		"target": "workbook",
		"format": "json",
		"workbook": map[string]any{
			"path":   "build/Book.xlsm",
			"name":   "Book.xlsm",
			"sheets": []map[string]any{},
		},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	var decoded map[string]any
	if err := json.Unmarshal(buf.Bytes(), &decoded); err != nil {
		t.Fatal(err)
	}
	if _, ok := decoded["target_state"].(map[string]any); !ok {
		t.Fatalf("expected target_state in inspect json: %s", buf.String())
	}
	if _, ok := decoded["session"].(map[string]any); !ok {
		t.Fatalf("expected session in inspect json: %s", buf.String())
	}
	if _, ok := decoded["warnings"].([]any); !ok {
		t.Fatalf("expected warnings in inspect json: %s", buf.String())
	}
	if _, ok := decoded["hints"].([]any); !ok {
		t.Fatalf("expected hints in inspect json: %s", buf.String())
	}
}

func TestWritePlainFailureIncludesLogsBeforeError(t *testing.T) {
	env := Failure("run", Error{Message: "macro failed"})
	env.Logs = []string{"[2026-04-29 21:12:03] start"}
	var buf bytes.Buffer
	if err := Write(&buf, env, false); err != nil {
		t.Fatal(err)
	}
	if got := buf.String(); got != "[2026-04-29 21:12:03] start\nmacro failed\n" {
		t.Fatalf("plain failure output = %q", got)
	}
}

func TestWriteJSONEnvelopeIncludesErrorLine(t *testing.T) {
	env := Failure("run", Error{
		Code:    "macro_failed",
		Message: "inputPath is required",
		Source:  "Main",
		Number:  5,
		Line:    10,
	})
	var buf bytes.Buffer
	if err := Write(&buf, env, true); err != nil {
		t.Fatal(err)
	}
	var decoded map[string]any
	if err := json.Unmarshal(buf.Bytes(), &decoded); err != nil {
		t.Fatal(err)
	}
	errorMap, ok := decoded["error"].(map[string]any)
	if !ok {
		t.Fatalf("error payload = %#v", decoded["error"])
	}
	if errorMap["line"] != float64(10) {
		t.Fatalf("error line = %#v", errorMap["line"])
	}
}

func TestWriteWithOptionsRendersDoctorChecklist(t *testing.T) {
	env := New("doctor")
	env.Diagnostics = map[string]any{
		"excel_installed":   true,
		"workbook_openable": true,
		"vbide_access":      false,
		"fix":               "Enable Trust access.",
	}
	env.Workbook = map[string]any{"path": "build/Book.xlsm"}
	env.Status = StatusFailed
	env.Error = &Error{Code: "vbide_access_denied", Message: "VBIDE access is not available.", Source: "Excel"}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"xlflow doctor", "Excel automation", "Workbook", "VBIDE access", "Enable Trust access.", "vbide_access_denied"} {
		if !strings.Contains(got, want) {
			t.Fatalf("doctor output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersDoctorChecklistFromDotNetBridge(t *testing.T) {
	env := New("doctor")
	env.Diagnostics = map[string]any{
		"requested_bridge": "auto",
		"selected_bridge":  "dotnet",
		"fallback":         false,
		"legacy":           false,
		"protocol_version": float64(1),
		"runtime":          map[string]any{"os": "Windows 11"},
		"excel": map[string]any{
			"com_activation":      true,
			"version":             "16.0",
			"build":               "12345",
			"vbide_access":        true,
			"automation_security": float64(1),
			"trust_vba_access":    nil,
		},
	}
	env.Status = StatusOK
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"xlflow doctor", "Selected bridge:", "dotnet", "Requested bridge:", "auto", "Fallback:", "no", "Bridge role:", "primary", "Excel automation", "VBIDE access"} {
		if !strings.Contains(got, want) {
			t.Fatalf("doctor output missing %q:\n%s", want, got)
		}
	}
	if strings.Contains(got, "FAILED") {
		t.Fatalf("doctor output should not contain FAILED for dotnet bridge:\n%s", got)
	}
}

func TestWriteWithOptionsRendersRunSummaryAndDebug(t *testing.T) {
	env := New("run")
	env.Macro = map[string]any{"name": "Main.Run", "duration_ms": 42}
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "saved": false}
	env.Target = map[string]any{"kind": "file", "path": "build/Book.xlsm", "description": "Saved workbook file on disk"}
	env.Session = map[string]any{"active": false, "workbook_path": "build/Book.xlsm", "dirty": false, "save_required": false}
	env.Debug = map[string]any{"events": []map[string]any{{"message": "starting run", "runtime_mode": "headless"}}, "count": 1}
	env.UI = map[string]any{"events": []map[string]any{{"kind": "msgbox", "dialog_id": "confirm-save", "response_source": "default", "resolved_result": "yes"}, {"kind": "inputbox", "dialog_id": "customer-name", "response_source": "default", "resolved_value": "[redacted]"}}}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Target:", "Saved workbook file on disk", "Session state:", "inactive", "Main.Run", "42ms", "left unchanged", "Debug", "log message=starting run mode=headless", "UI", "Events:", "msgbox id=confirm-save source=default result=yes", "inputbox id=customer-name source=default value=[redacted]"} {
		if !strings.Contains(got, want) {
			t.Fatalf("run output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersTestUISummary(t *testing.T) {
	env := New("test")
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "saved": false}
	env.Tests = []map[string]any{{"name": "TestDialogDefaults", "module": "DialogTests", "status": "passed", "duration_ms": 12}}
	env.Debug = map[string]any{"events": []map[string]any{{"message": "test emitted", "runtime_mode": "test"}}, "count": 1}
	env.UI = map[string]any{"events": []map[string]any{{"kind": "msgbox", "dialog_id": "test-confirm", "response_source": "scripted", "resolved_result": "ok"}}}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Summary:", "1 passed, 0 failed, 1 total", "Debug", "log message=test emitted mode=test", "UI", "Events:", "msgbox id=test-confirm source=scripted result=ok"} {
		if !strings.Contains(got, want) {
			t.Fatalf("test output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersRunUIOnFailure(t *testing.T) {
	env := Failure("run", Error{Code: "macro_timeout", Message: "Macro timed out."})
	env.Debug = map[string]any{"events": []map[string]any{{"message": "before timeout", "runtime_mode": "headless"}}, "count": 1}
	env.UI = map[string]any{"events": []map[string]any{{"kind": "msgbox", "dialog_id": "confirm-save", "response_source": "default", "resolved_result": "yes"}}}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Macro timed out.", "Debug", "log message=before timeout mode=headless", "UI", "Events:", "msgbox id=confirm-save source=default result=yes"} {
		if !strings.Contains(got, want) {
			t.Fatalf("failed run output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersTestUIOnFailure(t *testing.T) {
	env := Failure("test", Error{Code: "test_environment_failed", Message: "VBIDE access is not available."})
	env.Debug = map[string]any{"events": []map[string]any{{"message": "before failure", "runtime_mode": "test"}}, "count": 1}
	env.UI = map[string]any{"events": []map[string]any{{"kind": "inputbox", "dialog_id": "customer-name", "response_source": "default", "resolved_value": "[redacted]"}}}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"VBIDE access is not available.", "Debug", "log message=before failure mode=test", "UI", "Events:", "inputbox id=customer-name source=default value=[redacted]"} {
		if !strings.Contains(got, want) {
			t.Fatalf("failed test output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersDebugTruncationHint(t *testing.T) {
	env := New("run")
	env.Debug = map[string]any{
		"events":    []map[string]any{{"message": "recent line", "runtime_mode": "headless"}},
		"count":     42,
		"truncated": true,
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Debug", "Events:", "42", "Retention:", "truncated to recent events", "log message=recent line mode=headless"} {
		if !strings.Contains(got, want) {
			t.Fatalf("debug output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersMacrosHintsAndWarnings(t *testing.T) {
	env := New("macros")
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "session": true, "session_mode": "explicit", "needs_save": true}
	env.Target = map[string]any{"kind": "live_session", "path": "build/Book.xlsm", "description": "Workbook currently open through xlflow session"}
	env.Session = map[string]any{"active": true, "workbook_path": "build/Book.xlsm", "dirty": true, "save_required": true}
	env.Macros = []map[string]any{}
	env.Warnings = []map[string]any{{"code": "save_required", "message": "The live session workbook differs from disk. Run `xlflow save --session` to persist workbook changes."}}
	env.Hints = []map[string]any{
		{"code": "macros_empty_before_push", "message": "If you edited source files, run `xlflow push --session` before `xlflow macros --session`."},
		{"code": "macros_read_from_workbook", "message": "`macros` discovers procedures from the workbook, not directly from source files."},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"live session", "SAVE REQUIRED", "Warnings:", "Hints:", "macros_empty_before_push", "macros_read_from_workbook"} {
		if !strings.Contains(got, want) {
			t.Fatalf("macros output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersListFormsSummary(t *testing.T) {
	env := New("list")
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "session": true, "session_mode": "auto", "needs_save": true}
	env.Target = map[string]any{"kind": "live_session", "path": "build/Book.xlsm"}
	env.Forms = []map[string]any{
		{"name": "CustomerForm", "component_type": "MSForm", "has_frx": true, "source_path": "src/forms/CustomerForm.frm"},
		{"name": "OrderForm", "component_type": "MSForm", "has_frx": false, "source_path": "src/forms/Sales/OrderForm.frm"},
	}
	env.Warnings = []map[string]any{{"code": "save_required", "message": "The live session workbook differs from disk. Run `xlflow save --session` to persist workbook changes."}}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"xlflow list", "Forms:", "2", "CustomerForm", "has .frx", "src/forms/Sales/OrderForm.frm", "SAVE REQUIRED"} {
		if !strings.Contains(got, want) {
			t.Fatalf("list output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsDoesNotRenderZeroFormsWhenListFailed(t *testing.T) {
	env := Failure("list", Error{Code: "vbproject_access_denied", Message: "VBProject access is denied."})
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "session": true, "session_mode": "auto", "needs_save": true}
	env.Session = map[string]any{"active": true, "workbook_path": "build/Book.xlsm", "dirty": true, "save_required": true}
	env.Warnings = []map[string]any{{"code": "save_required", "message": "The live session workbook differs from disk. Run `xlflow save --session` to persist workbook changes."}}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Forms:", "unavailable", "SAVE REQUIRED", "Warnings:", "save_required"} {
		if !strings.Contains(got, want) {
			t.Fatalf("failed list output missing %q:\n%s", want, got)
		}
	}
	if strings.Contains(got, "Forms:          0") || strings.Contains(got, "Forms:\t0") {
		t.Fatalf("failed list output should not render zero forms:\n%s", got)
	}
}

func TestWriteWithOptionsRendersRunPreflightAnalysis(t *testing.T) {
	env := Failure("run", Error{Code: "analyze_failed", Message: "1 source issue(s) must be fixed before run", Source: "xlflow", Phase: "preflight"})
	env.Analysis = []map[string]any{{
		"code":       "VBA105",
		"severity":   "error",
		"file":       "src/modules/Main.bas",
		"line":       3,
		"message":    "XlflowLog is called but no Public standard-module definition was found in source.",
		"suggestion": "Replace the legacy helper with `XlflowDebug.Log` and inspect output via `xlflow run --json`.",
	}}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"VBA105", "src/modules/Main.bas:3", "Suggestion:", "XlflowDebug.Log", "xlflow run --json"} {
		if !strings.Contains(got, want) {
			t.Fatalf("run preflight output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersRunDiagnostic(t *testing.T) {
	env := Failure("run", Error{Code: "macro_failed", Message: "Main Err 450", Source: "Main", Number: 450, Phase: "invoke_macro"})
	env.RunDiagnostic = map[string]any{
		"location":     map[string]any{"file": "src/modules/Main.bas", "module": "Main", "procedure": "Run", "line": 4},
		"likely_cause": "missing Set",
		"suggestion":   "Use Set result = ...",
		"nearby_code":  []string{"> 4 | result = FindRange()"},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Diagnostic", "missing Set", "Use Set result", "result = FindRange()"} {
		if !strings.Contains(got, want) {
			t.Fatalf("run diagnostic output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersInspectSnapshotMetadata(t *testing.T) {
	env := New("inspect")
	env.Target = map[string]any{"kind": "file", "path": "build/Book.xlsm", "description": "Saved workbook file on disk"}
	env.Session = map[string]any{"active": true, "workbook_path": "build/Book.xlsm", "dirty": true, "save_required": true}
	env.Warnings = []map[string]any{{"code": "live_session_dirty", "message": "A live session exists and has unsaved changes. This command inspected the saved file, not the live workbook."}}
	env.Hints = []map[string]any{{"code": "userform_planned_commands", "message": "Related commands for deeper UserForm inspection include `xlflow form snapshot <name> --out <path>`."}}
	env.Inspect = map[string]any{
		"target": "range",
		"target_info": map[string]any{
			"kind": "file",
			"note": "This command inspected the saved workbook file, not an unsaved live Excel session.",
		},
		"range": map[string]any{
			"sheet":          "Visible",
			"range":          "A1:B2",
			"row_count":      2,
			"column_count":   2,
			"style_included": true,
			"values":         [][]any{{"A1", nil}, {"A2", "B2"}},
		},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Target:", "Session state:", "Warnings:", "Hints:", "live_session_dirty", "userform_planned_commands", "Snapshot", "saved workbook file", "Style:         included"} {
		if !strings.Contains(got, want) {
			t.Fatalf("inspect output missing %q:\n%s", want, got)
		}
	}
	if strings.Count(got, "Target:") != 1 || strings.Count(got, "Session state:") != 1 {
		t.Fatalf("inspect output should render target/session once:\n%s", got)
	}
}

func TestWriteWithOptionsRendersInspectFormSummary(t *testing.T) {
	env := New("inspect")
	env.Target = map[string]any{"kind": "live_session", "path": "build/Book.xlsm", "description": "Workbook currently open through xlflow session"}
	env.Session = map[string]any{"active": true, "workbook_path": "build/Book.xlsm", "dirty": false, "save_required": false}
	env.Warnings = []map[string]any{{"code": "runtime_form_loads_initialize", "message": "Runtime inspection loads the form and executes UserForm_Initialize."}}
	env.Inspect = map[string]any{
		"target": "form",
		"format": "text",
		"source": "excel_com",
		"form": map[string]any{
			"name":              "UserForm1",
			"basis":             "runtime",
			"caption":           "Order Entry Form",
			"width":             308,
			"height":            372,
			"coordinate_system": "parent-relative",
			"controls": []map[string]any{
				{"name": "txtOrderDate", "type": "TextBox", "value": "2026/05/11", "left": 108, "top": 15},
				{"name": "cmdRegister", "type": "CommandButton", "caption": "Save Order", "left": 72, "top": 231},
			},
		},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"xlflow inspect", "Basis:", "runtime", "Form:", "UserForm1", "txtOrderDate", "value=2026/05/11", "cmdRegister", "caption=Save Order", "Warnings:"} {
		if !strings.Contains(got, want) {
			t.Fatalf("inspect form output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersInspectEmptyRangeWarnings(t *testing.T) {
	env := New("inspect")
	env.Target = map[string]any{"kind": "file", "path": "build/Book.xlsm", "description": "Saved workbook file on disk"}
	env.Session = map[string]any{"active": true, "workbook_path": "build/Book.xlsm", "dirty": true, "save_required": true}
	env.Warnings = []map[string]any{{"code": "live_session_dirty", "message": "A live session exists and has unsaved changes. This command inspected the saved file, not the live workbook."}}
	env.Inspect = map[string]any{
		"target": "range",
		"range": map[string]any{
			"sheet":        "Visible",
			"range":        "A1:A1",
			"row_count":    0,
			"column_count": 0,
			"values":       [][]any{},
		},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Values: <empty>", "Warnings:", "live_session_dirty"} {
		if !strings.Contains(got, want) {
			t.Fatalf("inspect empty range output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersCompileDiagnostic(t *testing.T) {
	env := Failure("run", Error{Code: "vba_compile_failed", Message: "Compile error", Source: "Main", Phase: "compile_vba", Line: 8})
	env.RunDiagnostic = map[string]any{
		"kind":        "compile",
		"message":     []string{"Compile error:", "Method or data member not found"},
		"location":    map[string]any{"module": "Main", "line": 8, "column": 5, "token": "DisplayGridlines"},
		"nearby_code": []string{"> 8 |   .DisplayGridlines = False", "    |     ^^^^^^^^^^^^^^^^"},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Kind", "compile", "Method or data member not found", "Main line 8 column 5 DisplayGridlines", ".DisplayGridlines"} {
		if !strings.Contains(got, want) {
			t.Fatalf("compile diagnostic output missing %q:\n%s", want, got)
		}
	}
	if strings.Count(got, "Message") != 1 {
		t.Fatalf("compile diagnostic should render one message block:\n%s", got)
	}
}

func TestWriteWithOptionsRendersSessionOnlyPushResult(t *testing.T) {
	env := New("push")
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "saved": false, "session": true, "session_mode": "auto", "needs_save": true}
	env.Source = map[string]any{"changed_only": true, "changed": true}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"SAVE REQUIRED", "xlflow save before session stop", "auto-reused matching xlflow session workbook"} {
		if !strings.Contains(got, want) {
			t.Fatalf("push output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersRunSessionUnsavedWarning(t *testing.T) {
	env := New("run")
	env.Macro = map[string]any{"name": "Main.Run", "duration_ms": 42}
	env.Runtime = map[string]any{"mode": "headless", "source": "command", "injected": true}
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "saved": false, "session": true, "session_mode": "explicit", "needs_save": true}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Runtime:", "headless", "workbook marker injected", "SAVE REQUIRED", "xlflow save before session stop", "explicit xlflow session workbook"} {
		if !strings.Contains(got, want) {
			t.Fatalf("run output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersRunFailureSessionSaveRequirement(t *testing.T) {
	env := Failure("run", Error{Code: "macro_failed", Message: "macro failed"})
	env.Macro = map[string]any{"name": "Main.Run", "duration_ms": 42}
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "saved": false, "session": true, "session_mode": "explicit", "needs_save": true}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"macro failed", "SAVE REQUIRED", "xlflow save before session stop"} {
		if !strings.Contains(got, want) {
			t.Fatalf("run failure output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersRunSaveAsAndSaveRequirement(t *testing.T) {
	env := New("run")
	env.Macro = map[string]any{"name": "Main.Run", "duration_ms": 42}
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "saved": false, "save_as": "build/Copy.xlsm", "session": true, "session_mode": "explicit", "needs_save": true}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"copied to build/Copy.xlsm", "SAVE REQUIRED", "xlflow save before session stop"} {
		if !strings.Contains(got, want) {
			t.Fatalf("run save-as output missing %q:\n%s", want, got)
		}
	}
	if saveIdx, resultIdx := strings.Index(got, "Save:"), strings.Index(got, "Result:"); saveIdx == -1 || resultIdx == -1 || saveIdx > resultIdx {
		t.Fatalf("expected Save warning before Result summary:\n%s", got)
	}
}

func TestWriteWithOptionsRendersVersionVerboseDetails(t *testing.T) {
	env := New("version")
	env.Version = map[string]any{
		"version":         "1.2.3",
		"commit":          "abc123",
		"date":            "2026-05-02T00:00:00Z",
		"executable_path": `C:\tools\xlflow.exe`,
		"go_version":      "go1.25.0",
		"module_path":     "github.com/harumiWeb/xlflow",
		"build_settings":  []map[string]any{{"key": "vcs.revision", "value": "abc123"}},
		"scripts":         []map[string]any{{"command": "run", "source": "embedded"}, {"command": "push", "source": "override", "path": `C:\dev\go\xlflow\scripts\push.ps1`}},
		"features":        []map[string]any{{"name": "run-entry-fallback", "description": "Use project.entry when the macro is omitted."}},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Executable:", `C:\tools\xlflow.exe`, "Build settings", "Scripts", "run-entry-fallback", "Use project.entry"} {
		if !strings.Contains(got, want) {
			t.Fatalf("version verbose output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersSessionStatusSaveRequirement(t *testing.T) {
	env := New("session")
	env.Session = map[string]any{"running": true, "workbook_open": true, "needs_save": true, "live_newer_than_disk": true, "source_of_truth": "live_workbook", "userforms_present": true, "userform_count": 2}
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "needs_save": true}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Running:", "Workbook open:", "SAVE REQUIRED", "live workbook is newer than disk", "Source of truth:", "live_workbook", "UserForms:", "true (2)", "xlflow save before session stop"} {
		if !strings.Contains(got, want) {
			t.Fatalf("session status output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersUnknownUserFormsState(t *testing.T) {
	env := New("session")
	env.Session = map[string]any{"running": true, "workbook_open": true, "userforms_known": false}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	if !strings.Contains(got, "UserForms:") || !strings.Contains(got, "unknown") {
		t.Fatalf("expected unknown UserForms state:\n%s", got)
	}
}

func TestWriteWithOptionsRendersFormBuildSpecFailureMetadata(t *testing.T) {
	env := Failure("form build", Error{Code: "spec_parse_failed", Message: "yaml: line 6: did not find expected node content"})
	env.Spec = map[string]any{
		"path":       "src/forms/specs/UserForm1.yaml",
		"format":     "yaml",
		"line":       6,
		"suggestion": `Try quoting scalar strings or use JSON if YAML syntax is uncertain. For an empty caption, use caption: "" rather than caption: -.`,
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Spec:", "src/forms/specs/UserForm1.yaml", "Spec format:", "YAML", "Spec location:", "line 6", "Remediation:", `caption: ""`} {
		if !strings.Contains(got, want) {
			t.Fatalf("form build failure output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersUserFormStateNote(t *testing.T) {
	env := New("form snapshot")
	env.Target = map[string]any{"kind": "live_session", "path": "build/Book.xlsm"}
	env.Session = map[string]any{"active": true, "dirty": true, "save_required": true, "live_newer_than_disk": true, "userforms_present": true}
	env.Forms = map[string]any{"name": "UserForm1", "basis": "designer", "control_count": 1}
	env.Output = map[string]any{"path": "src/forms/specs/UserForm1.yaml", "format": "yaml"}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	if !strings.Contains(got, "State note:") || !strings.Contains(got, "UserForm project: save before disk inspect/pull review.") {
		t.Fatalf("missing userform state note:\n%s", got)
	}
}

func TestWriteWithOptionsRendersTargetNoteStateNote(t *testing.T) {
	env := New("inspect form")
	env.Target = map[string]any{"kind": "file", "path": "build/Book.xlsm", "note": "Strict designer inspection used a temporary workbook copy plus helper module to recover concrete control types."}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	if !strings.Contains(got, "State note:") || !strings.Contains(got, "Strict designer inspection used a temporary workbook copy plus helper module to recover concrete control types.") {
		t.Fatalf("missing target note state note:\n%s", got)
	}
}

func TestWriteWithOptionsRendersExportImageSummary(t *testing.T) {
	env := New("export-image")
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "session": true, "session_mode": "auto", "needs_save": true}
	env.Target = map[string]any{"kind": "live_session", "path": "build/Book.xlsm", "sheet": "QR", "range": "A1:AE31"}
	env.Output = map[string]any{"path": ".xlflow/artifacts/images/Book/QR_A1-AE31_20260509-083012.png", "format": "png", "default": true, "width_px": 620, "height_px": 620}
	env.Warnings = []map[string]any{{"code": "temporary_object_cleanup_failed", "message": "The image was exported, but xlflow could not delete a temporary chart object."}}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Export target:", "live session workbook", "QR", "Selection:", "A1:AE31", "PNG", "620 x 620 px", "SAVE REQUIRED", "temporary_object_cleanup_failed"} {
		if !strings.Contains(got, want) {
			t.Fatalf("export-image output missing %q:\n%s", want, got)
		}
	}
	if strings.Count(got, "Target:") != 1 {
		t.Fatalf("export-image output should render one Target label:\n%s", got)
	}
}

func TestWriteWithOptionsRendersFormSnapshotSummary(t *testing.T) {
	env := New("form snapshot")
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "session": true, "session_mode": "explicit", "needs_save": true}
	env.Target = map[string]any{"kind": "live_session", "path": "build/Book.xlsm"}
	env.Forms = map[string]any{"name": "UserForm1", "basis": "designer", "caption": "Order Entry", "coordinate_system": "parent-relative", "control_count": 3}
	env.Output = map[string]any{"path": "artifacts/UserForm1.form.yaml", "format": "yaml"}
	env.Warnings = []map[string]any{{"code": "save_required", "message": "Run `xlflow save --session` to persist workbook changes."}}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Snapshot target:", "live session", "Form:", "UserForm1", "Basis:", "designer", "Coordinates:", "parent-relative", "Controls:", "3", "Output:", "artifacts/UserForm1.form.yaml", "Format:", "YAML", "SAVE REQUIRED", "save_required"} {
		if !strings.Contains(got, want) {
			t.Fatalf("form snapshot output missing %q:\n%s", want, got)
		}
	}
	if strings.Count(got, "Target:") != 1 {
		t.Fatalf("form snapshot output should render one Target label:\n%s", got)
	}
}

func TestWriteWithOptionsRendersFormExportImageSummary(t *testing.T) {
	env := New("form export-image")
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "session": true, "session_mode": "auto", "needs_save": true}
	env.Target = map[string]any{"kind": "live_session", "path": "build/Book.xlsm", "form": "UserForm1", "capture_state": "temporary_copy"}
	env.Forms = map[string]any{"name": "UserForm1", "basis": "runtime", "initializer": "InitializeForm"}
	env.Output = map[string]any{"path": "artifacts/UserForm1.png", "format": "png", "width_px": 308, "height_px": 372}
	env.Warnings = []map[string]any{{"code": "userform_image_export_experimental", "message": "experimental"}}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Export target:", "live session", "Form:", "UserForm1", "Basis:", "runtime", "Initializer:", "InitializeForm", "Capture:", "temporary_copy", "Output:", "artifacts/UserForm1.png", "Format:", "PNG", "308 x 372 px", "SAVE REQUIRED", "userform_image_export_experimental"} {
		if !strings.Contains(got, want) {
			t.Fatalf("form export-image output missing %q:\n%s", want, got)
		}
	}
	if strings.Count(got, "Target:") != 1 {
		t.Fatalf("form export-image output should render one Target label:\n%s", got)
	}
}

func TestWriteWithOptionsRendersFormWriteSummary(t *testing.T) {
	env := New("form apply")
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "session": true, "session_mode": "explicit", "needs_save": true}
	env.Target = map[string]any{"kind": "live_session", "path": "build/Book.xlsm"}
	env.Forms = map[string]any{
		"name":              "UserForm1",
		"basis":             "designer",
		"action":            "apply",
		"coordinate_system": "parent-relative",
		"control_count":     4,
		"spec_path":         "src/forms/specs/UserForm1.yaml",
		"overwrite":         false,
	}
	env.Warnings = []map[string]any{{"code": "save_required", "message": "Run `xlflow save --session` to persist workbook changes."}}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Write target:", "live session", "Action:", "apply", "Form:", "UserForm1", "Basis:", "designer", "Coordinates:", "parent-relative", "Controls:", "4", "Spec:", "src/forms/specs/UserForm1.yaml", "Overwrite:", "false", "SAVE REQUIRED"} {
		if !strings.Contains(got, want) {
			t.Fatalf("form write output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersEditSummary(t *testing.T) {
	env := New("edit")
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "session": true, "session_mode": "explicit", "needs_save": true}
	env.Target = map[string]any{"kind": "live_session", "path": "build/Book.xlsm", "description": "Workbook currently open through xlflow session"}
	env.Session = map[string]any{"active": true, "workbook_path": "build/Book.xlsm", "dirty": true, "save_required": true}
	env.Edit = map[string]any{
		"kind":  "cell",
		"sheet": "Input",
		"cell":  "B2",
		"mutation": map[string]any{
			"value": map[string]any{"before": "", "after": "ABC123"},
		},
		"events": map[string]any{
			"mode":                 "on",
			"enable_events_before": true,
			"enable_events_after":  true,
			"restored":             true,
		},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"xlflow edit", "Edit target:", "Selection:", "Input!B2", "Mutation:", "ABC123", "Events:", "mode=on", "SAVE REQUIRED"} {
		if !strings.Contains(got, want) {
			t.Fatalf("edit output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersFormulaMutationAheadOfCalculatedValue(t *testing.T) {
	env := New("edit")
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "session": true, "session_mode": "explicit", "needs_save": true}
	env.Target = map[string]any{"kind": "live_session", "path": "build/Book.xlsm"}
	env.Session = map[string]any{"active": true, "workbook_path": "build/Book.xlsm", "dirty": true, "save_required": true}
	env.Edit = map[string]any{
		"kind":  "cell",
		"sheet": "Input",
		"cell":  "B2",
		"mutation": map[string]any{
			"value":   map[string]any{"before": "1", "after": "3"},
			"formula": map[string]any{"before": "=1", "after": "=1+2"},
		},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	if !strings.Contains(got, "formula -> =1+2") {
		t.Fatalf("expected formula summary, got:\n%s", got)
	}
}

func TestWriteWithOptionsRendersBridgeHost(t *testing.T) {
	env := New("run")
	env.Bridge = map[string]any{"host": "powershell.exe", "edition": "Desktop", "version": "5.1"}
	env.Macro = map[string]any{"name": "Main.Run", "duration_ms": 42}
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "saved": false}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Bridge:", "powershell.exe", "Desktop", "5.1"} {
		if !strings.Contains(got, want) {
			t.Fatalf("bridge output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersTestFailures(t *testing.T) {
	env := Failure("test", Error{Code: "test_failed", Message: "1 of 2 test(s) failed"})
	env.Runtime = map[string]any{"mode": "test", "source": "command", "injected": true}
	env.Tests = []map[string]any{
		{"name": "TestOk", "module": "Tests", "status": "passed", "duration_ms": 3},
		{"name": "TestBad", "module": "Tests", "status": "failed", "duration_ms": 5, "error": map[string]any{"message": "expected 1"}},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"1 passed, 1 failed, 2 total", "Tests.TestOk", "Tests.TestBad", "expected 1"} {
		if !strings.Contains(got, want) {
			t.Fatalf("test output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersTestRuntimeSummary(t *testing.T) {
	env := New("test")
	env.Runtime = map[string]any{"mode": "test", "source": "command", "injected": true}
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "saved": true}
	env.Tests = []map[string]any{{"name": "TestRun", "module": "MainTests", "status": "passed", "duration_ms": 5}}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Runtime:", "test", "source=command", "workbook marker injected", "1 passed, 0 failed, 1 total"} {
		if !strings.Contains(got, want) {
			t.Fatalf("test runtime output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersDiscoveredButUnrunTestsAsNotRun(t *testing.T) {
	env := Failure("test", Error{Code: "duplicate_test_name", Message: "duplicate VBA test name(s): TestSame"})
	env.Tests = []map[string]any{
		{"name": "TestSame", "module": "TestsA"},
		{"name": "TestSame", "module": "TestsB"},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"0 passed, 0 failed, 2 not run, 2 total", "[-] TestsA.TestSame", "[-] TestsB.TestSame"} {
		if !strings.Contains(got, want) {
			t.Fatalf("test output missing %q:\n%s", want, got)
		}
	}
	if strings.Contains(got, "[x] TestsA.TestSame") || strings.Contains(got, "[x] TestsB.TestSame") {
		t.Fatalf("unrun tests should not be marked failed:\n%s", got)
	}
}

func TestWriteWithOptionsRendersLintIssues(t *testing.T) {
	env := Failure("lint", Error{Code: "lint_failed", Message: "1 lint issue(s) found"})
	env.Issues = []map[string]any{{
		"code":     "VB001",
		"severity": "warning",
		"file":     "src/modules/Main.bas",
		"line":     1,
		"message":  "missing Option Explicit",
	}}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"VB001", "src/modules/Main.bas:1", "missing Option Explicit"} {
		if !strings.Contains(got, want) {
			t.Fatalf("lint output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersDiffSummary(t *testing.T) {
	env := New("diff")
	env.Diff = map[string]any{"summary": map[string]any{"total_diffs": 2, "sheet_diffs": 1, "cell_diffs": 1, "vba_diffs": 0}}
	env.Logs = []string{"Sheet: + Result", "A1 value: \"old\" -> \"new\""}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Total diffs", "2", "Sheet Diffs", "1", "A1 value"} {
		if !strings.Contains(got, want) {
			t.Fatalf("diff output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersFmtSummary(t *testing.T) {
	env := New("fmt")
	env.Target = map[string]any{
		"kind":        "source",
		"path":        "src",
		"description": "source files",
	}
	env.Output = map[string]any{
		"mode":            "check",
		"changed":         1,
		"unchanged":       1,
		"skipped":         1,
		"total":           3,
		"changed_paths":   []any{"src/modules/Main.bas"},
		"skipped_paths":   []any{"src/forms/UserForm1.frm"},
		"skipped_reasons": []map[string]any{{"path": "src/forms/UserForm1.frm", "reason": "unsupported extension: .frm"}},
	}
	env.Warnings = []map[string]any{
		{"code": "fmt_skipped_unsupported_extension", "message": "Skipped unsupported file: src/forms/UserForm1.frm"},
	}
	env.Hints = []map[string]any{
		{"code": "fmt_write_hint", "message": "Run `xlflow fmt --write` to apply formatting changes."},
	}

	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{
		"Target:",
		"source files",
		"Summary:",
		"1 not formatted",
		"3 total",
		"src/modules/Main.bas",
		"src/forms/UserForm1.frm",
		"Warnings:",
		"fmt_skipped_unsupported_extension",
		"Hints:",
		"xlflow fmt --write",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("fmt output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersFmtInspectSummary(t *testing.T) {
	env := New("fmt")
	env.Target = map[string]any{
		"kind":        "source",
		"path":        "src",
		"description": "source files",
	}
	env.Output = map[string]any{
		"mode":      "inspect",
		"changed":   1,
		"unchanged": 0,
		"skipped":   0,
		"total":     1,
	}
	env.Logs = []string{"1 file(s) would be formatted"}

	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{
		"Summary:",
		"1 would be formatted",
		"1 total",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("fmt inspect output missing %q:\n%s", want, got)
		}
	}
	if strings.Contains(got, "1 changed") {
		t.Fatalf("fmt inspect output should not say changed:\n%s", got)
	}
}

func TestWriteWithOptionsRendersFmtWriteSummary(t *testing.T) {
	env := New("fmt")
	env.Target = map[string]any{
		"kind":        "source",
		"path":        "src",
		"description": "source files",
	}
	env.Output = map[string]any{
		"mode":      "write",
		"changed":   1,
		"unchanged": 0,
		"skipped":   0,
		"total":     1,
	}
	env.Logs = []string{"1 file(s) formatted"}

	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{
		"Summary:",
		"1 formatted",
		"1 total",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("fmt write output missing %q:\n%s", want, got)
		}
	}
	if strings.Contains(got, "1 changed") {
		t.Fatalf("fmt write output should not say changed:\n%s", got)
	}
}

func TestWriteWithOptionsRendersStatusBaseline(t *testing.T) {
	env := New("status")
	env.Project = map[string]any{
		"root":          ".",
		"workbook_path": "build/Book.xlsm",
		"src_paths":     []any{"src/modules", "src/classes", "src/forms", "src/workbook"},
		"project_name":  "sample",
	}
	env.Session = map[string]any{
		"active":               false,
		"dirty":                false,
		"save_required":        false,
		"live_newer_than_disk": false,
	}
	env.State = map[string]any{
		"src_newer_than_workbook":      false,
		"live_session_newer_than_disk": false,
		"workbook_saved":               true,
		"source_of_truth":              "saved_workbook",
		"workbook_last_modified_at":    "2026-05-23T10:00:00Z",
		"latest_source_modified_at":    "2026-05-22T10:00:00Z",
	}
	env.Logs = []string{"status reported"}

	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{
		"status",
		"Root:",
		"Workbook:",
		"build/Book.xlsm",
		"Source:",
		"src/modules",
		"Session:",
		"inactive",
		"Source newer:",
		"false",
		"Workbook saved:",
		"true",
		"Source of truth:",
		"saved_workbook",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("status output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersStatusSessionActiveDirty(t *testing.T) {
	env := New("status")
	env.Project = map[string]any{
		"root":          ".",
		"workbook_path": "build/Book.xlsm",
		"src_paths":     []any{"src/modules"},
	}
	env.Session = map[string]any{
		"active":               true,
		"dirty":                true,
		"save_required":        true,
		"live_newer_than_disk": true,
	}
	env.State = map[string]any{
		"src_newer_than_workbook":      false,
		"live_session_newer_than_disk": true,
		"workbook_saved":               false,
		"source_of_truth":              "live_workbook",
	}
	env.Warnings = []map[string]any{
		{"code": "session_dirty", "message": "The live session workbook has unsaved changes."},
	}
	env.Hints = []map[string]any{
		{"code": "save_session", "message": "Run `xlflow save --session` to persist the live workbook to disk."},
	}
	env.Logs = []string{"status reported"}

	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{
		"Session:",
		"active, dirty",
		"Dirty:",
		"true",
		"Live newer:",
		"true",
		"Source of truth:",
		"live_workbook",
		"Warnings:",
		"session_dirty",
		"Hints:",
		"save_session",
		"xlflow save --session",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("status output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersStatusSourceNewerThanWorkbook(t *testing.T) {
	env := New("status")
	env.Project = map[string]any{
		"root":          ".",
		"workbook_path": "build/Book.xlsm",
		"src_paths":     []any{"src/modules"},
	}
	env.Session = map[string]any{
		"active":               false,
		"dirty":                false,
		"save_required":        false,
		"live_newer_than_disk": false,
	}
	env.State = map[string]any{
		"src_newer_than_workbook":      true,
		"live_session_newer_than_disk": false,
		"workbook_saved":               true,
		"source_of_truth":              "saved_workbook",
		"workbook_last_modified_at":    "2026-05-22T10:00:00Z",
		"latest_source_modified_at":    "2026-05-23T10:00:00Z",
	}
	env.Warnings = []map[string]any{
		{"code": "source_newer_than_workbook", "message": "Source files are newer than the saved workbook."},
	}
	env.Hints = []map[string]any{
		{"code": "push_source", "message": "Run `xlflow push` to import the latest source into the workbook."},
	}
	env.Logs = []string{"status reported"}

	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{
		"Source newer:",
		"true",
		"Warnings:",
		"source_newer_than_workbook",
		"Hints:",
		"push_source",
		"xlflow push",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("status output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersStatusSessionLiveNewerThanDisk(t *testing.T) {
	env := New("status")
	env.Project = map[string]any{
		"root":          ".",
		"workbook_path": "build/Book.xlsm",
		"src_paths":     []any{"src/modules"},
	}
	env.Session = map[string]any{
		"active":               true,
		"dirty":                true,
		"save_required":        true,
		"live_newer_than_disk": true,
	}
	env.State = map[string]any{
		"src_newer_than_workbook":      false,
		"live_session_newer_than_disk": true,
		"workbook_saved":               false,
		"source_of_truth":              "live_workbook",
	}
	env.Warnings = []map[string]any{
		{"code": "session_dirty", "message": "The live session workbook has unsaved changes."},
		{"code": "live_session_newer_than_disk", "message": "The live session workbook is newer than the saved workbook on disk."},
	}
	env.Hints = []map[string]any{
		{"code": "save_session", "message": "Run `xlflow save --session` to persist the live workbook to disk."},
		{"code": "save_before_push", "message": "Run `xlflow save --session` before `xlflow push` when the live session has unsaved changes you want to keep."},
	}
	env.Logs = []string{"status reported"}

	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{
		"Session:",
		"Status:",
		"active, dirty",
		"Dirty:",
		"true",
		"State:",
		"Live newer:",
		"true",
		"Workbook saved:",
		"false",
		"Source of truth:",
		"live_workbook",
		"Warnings:",
		"session_dirty",
		"live_session_newer_than_disk",
		"Hints:",
		"save_session",
		"save_before_push",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("status output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersStatusSessionActiveDirtyUnknown(t *testing.T) {
	env := New("status")
	env.Project = map[string]any{
		"root":          ".",
		"workbook_path": "build/Book.xlsm",
		"src_paths":     []any{"src/modules"},
	}
	env.Session = map[string]any{
		"active":               true,
		"dirty":                nil,
		"save_required":        true,
		"live_newer_than_disk": true,
	}
	env.State = map[string]any{
		"src_newer_than_workbook":      false,
		"live_session_newer_than_disk": true,
		"workbook_saved":               false,
		"source_of_truth":              "live_workbook",
	}
	env.Warnings = []map[string]any{
		{"code": "session_dirty", "message": "The live session workbook has unsaved changes."},
	}
	env.Logs = []string{"status reported"}

	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{
		"Session:",
		"Status:",
		"active",
		"Warnings:",
		"session_dirty",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("status output missing %q:\n%s", want, got)
		}
	}
	if strings.Contains(got, "clean") {
		t.Fatalf("dirty unknown should not render as clean:\n%s", got)
	}
	if strings.Contains(got, "Dirty:") {
		t.Fatalf("dirty unknown should not render a dirty line:\n%s", got)
	}
	if strings.Contains(got, "Session state:") {
		t.Fatalf("status should not render target/session state from renderTargetSession:\n%s", got)
	}
}

func TestWriteWithOptionsRendersStatusSectionHeaders(t *testing.T) {
	env := New("status")
	env.Project = map[string]any{
		"root":          ".",
		"workbook_path": "build/Book.xlsm",
		"src_paths":     []any{"src/modules"},
	}
	env.Session = map[string]any{
		"active": false,
		"dirty":  false,
	}
	env.State = map[string]any{
		"src_newer_than_workbook": false,
		"workbook_saved":          true,
	}
	env.Logs = []string{"status reported"}

	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{
		"Project:",
		"Session:",
		"State:",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("status output missing section header %q:\n%s", want, got)
		}
	}
	if strings.Count(got, "Project:") != 1 {
		t.Fatalf("expected exactly one 'Project:' header, got %d:\n%s", strings.Count(got, "Project:"), got)
	}
	if strings.Count(got, "Session:") != 1 {
		t.Fatalf("expected exactly one 'Session:' header, got %d:\n%s", strings.Count(got, "Session:"), got)
	}
	if strings.Count(got, "State:") != 1 {
		t.Fatalf("expected exactly one 'State:' header, got %d:\n%s", strings.Count(got, "State:"), got)
	}
}

func TestWriteJSONEnvelopeIncludesProcessFields(t *testing.T) {
	env := New("process list")
	env.Process = []map[string]any{
		{"pid": 1234, "has_workbook": true},
		{"pid": 5678, "has_workbook": false},
	}
	var buf bytes.Buffer
	if err := Write(&buf, env, true); err != nil {
		t.Fatal(err)
	}
	var decoded map[string]any
	if err := json.Unmarshal(buf.Bytes(), &decoded); err != nil {
		t.Fatal(err)
	}
	processList, ok := decoded["process"].([]any)
	if !ok || len(processList) != 2 {
		t.Fatalf("expected 2 process entries in JSON envelope: %s", buf.String())
	}
	first, _ := processList[0].(map[string]any)
	if first["pid"] != float64(1234) || first["has_workbook"] != true {
		t.Fatalf("unexpected first process entry: %+v", first)
	}
}

func TestWriteJSONEnvelopeIncludesProcessCleanupResults(t *testing.T) {
	env := New("process cleanup")
	env.Process = map[string]any{
		"action":  "cleanup",
		"mode":    "auto",
		"total":   2,
		"results": []map[string]any{{"pid": 1234, "terminated": true, "method": "graceful"}},
	}
	var buf bytes.Buffer
	if err := Write(&buf, env, true); err != nil {
		t.Fatal(err)
	}
	var decoded map[string]any
	if err := json.Unmarshal(buf.Bytes(), &decoded); err != nil {
		t.Fatal(err)
	}
	processResult, ok := decoded["process"].(map[string]any)
	if !ok {
		t.Fatalf("expected process result in JSON envelope: %s", buf.String())
	}
	if processResult["action"] != "cleanup" || processResult["mode"] != "auto" {
		t.Fatalf("unexpected process cleanup result: %+v", processResult)
	}
}

func TestWriteWithOptionsRendersProcessListSummary(t *testing.T) {
	env := New("process list")
	env.Process = []map[string]any{
		{"pid": 1234, "has_workbook": true},
		{"pid": 5678, "has_workbook": false},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"xlflow process list", "1234", "5678", "has workbook"} {
		if !strings.Contains(got, want) {
			t.Fatalf("process list output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersProcessListEmptyResult(t *testing.T) {
	env := New("process list")
	env.Process = []map[string]any{}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	if !strings.Contains(got, "0 process") {
		t.Fatalf("expected empty process list to show 0 count:\n%s", got)
	}
}

func TestWriteWithOptionsRendersProcessListShowsUnknownForNullWorkbookState(t *testing.T) {
	env := New("process list")
	env.Process = []map[string]any{
		{"pid": float64(1234), "has_workbook": nil},
		{"pid": float64(5678), "has_workbook": true},
		{"pid": float64(9012), "has_workbook": false},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"unknown", "has workbook", "no workbook"} {
		if !strings.Contains(got, want) {
			t.Fatalf("process list output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersProcessCleanupSummary(t *testing.T) {
	env := New("process cleanup")
	env.Process = map[string]any{
		"action": "cleanup",
		"mode":   "pid",
		"total":  1,
		"results": []map[string]any{
			{"pid": 1234, "terminated": true, "method": "graceful"},
		},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"xlflow process cleanup", "PID:", "1234", "graceful", "1 terminated", "0 failed"} {
		if !strings.Contains(got, want) {
			t.Fatalf("process cleanup output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersProcessCleanupAllMode(t *testing.T) {
	env := New("process cleanup")
	env.Process = map[string]any{
		"action": "cleanup",
		"mode":   "all",
		"total":  3,
		"results": []map[string]any{
			{"pid": 1234, "terminated": true, "method": "force"},
			{"pid": 5678, "terminated": true, "method": "force"},
			{"pid": 9012, "terminated": false, "method": "force"},
		},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"xlflow process cleanup", "all", "3", "2 terminated", "1 failed"} {
		if !strings.Contains(got, want) {
			t.Fatalf("process cleanup --all output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersProcessCleanupPidModePartialFailure(t *testing.T) {
	env := New("process cleanup")
	env.Process = map[string]any{
		"action": "cleanup",
		"mode":   "pid",
		"total":  2,
		"results": []map[string]any{
			{"pid": 1234, "terminated": true, "method": "graceful"},
			{"pid": 5678, "terminated": false, "method": "none"},
		},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"1 terminated", "1 failed", "PID: 1234", "PID: 5678", "none"} {
		if !strings.Contains(got, want) {
			t.Fatalf("process cleanup pid partial failure output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersProcessCleanupFailure(t *testing.T) {
	env := Failure("process cleanup", Error{Code: "process_enumeration_failed", Message: "failed to enumerate Excel processes"})
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"process_enumeration_failed", "failed to enumerate Excel processes"} {
		if !strings.Contains(got, want) {
			t.Fatalf("process cleanup failure output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsDistinguishesProcessEmptyFromUnavailable(t *testing.T) {
	env := Failure("process list", Error{Code: "process_enumeration_failed", Message: "could not enumerate"})
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	if strings.Contains(got, "unavailable") || strings.Contains(got, "0 process") {
		t.Fatalf("failed process list should not render count or unavailable:\n%s", got)
	}
}

func TestWriteWithOptionsRendersProcessListNilProcessOK(t *testing.T) {
	env := New("process list")
	env.Process = nil
	env.Status = StatusOK
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	if !strings.Contains(got, "unavailable") {
		t.Fatalf("process list nil-Process + StatusOK should render unavailable:\n%s", got)
	}
	if !strings.Contains(got, "xlflow process list") {
		t.Fatalf("process list nil-Process + StatusOK should include command label:\n%s", got)
	}
}

func TestWriteWithOptionsRendersProcessDefaultsToLogs(t *testing.T) {
	env := New("process")
	env.Process = []map[string]any{
		{"pid": 1234, "has_workbook": true},
	}
	env.Logs = []string{"process log entry"}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	if strings.Contains(got, "has workbook") {
		t.Fatalf("process command should not render process-specific table:\n%s", got)
	}
	if !strings.Contains(got, "process log entry") {
		t.Fatalf("process command should render logs:\n%s", got)
	}
}

func TestWriteWithOptionsRendersProcessCleanupUnknownMode(t *testing.T) {
	env := New("process cleanup")
	env.Process = map[string]any{
		"action": "cleanup",
		"total":  1,
		"results": []map[string]any{
			{"pid": 1234, "terminated": true, "method": "graceful"},
		},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	if !strings.Contains(got, "unknown") {
		t.Fatalf("process cleanup with no mode should render as unknown:\n%s", got)
	}
}

func TestWriteWithOptionsRendersProcessCleanupNilProcessStatusFailed(t *testing.T) {
	env := New("process cleanup")
	env.Status = StatusFailed
	env.Process = nil
	env.Error = &Error{Code: "process_no_fallback", Message: "process error fallback"}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	if !strings.Contains(got, "process error fallback") {
		t.Fatalf("process cleanup nil-Process + StatusFailed should render logs:\n%s", got)
	}
	if strings.Contains(got, "Process cleanup result unavailable") {
		t.Fatalf("process cleanup nil-Process + StatusFailed should not render unavailable fallback:\n%s", got)
	}
	if strings.Contains(got, "Mode") {
		t.Fatalf("process cleanup nil-Process + StatusFailed should not render process table:\n%s", got)
	}
}

func TestWriteWithOptionsRendersProcessCleanupAutoZeroTargets(t *testing.T) {
	env := New("process cleanup")
	env.Process = map[string]any{
		"action":  "cleanup",
		"mode":    "auto",
		"total":   0,
		"results": []map[string]any{},
	}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	if !strings.Contains(got, "auto") {
		t.Fatalf("process cleanup auto zero targets should render mode:\n%s", got)
	}
	if !strings.Contains(got, "0") {
		t.Fatalf("process cleanup auto zero targets should render total:\n%s", got)
	}
}
