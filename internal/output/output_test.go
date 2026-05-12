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

func TestWriteJSONEnvelopeIncludesTrace(t *testing.T) {
	env := New("run")
	env.Trace = map[string]any{
		"enabled": true,
		"events":  []map[string]string{{"message": "start"}},
	}
	var buf bytes.Buffer
	if err := Write(&buf, env, true); err != nil {
		t.Fatal(err)
	}
	var decoded map[string]any
	if err := json.Unmarshal(buf.Bytes(), &decoded); err != nil {
		t.Fatal(err)
	}
	if _, ok := decoded["trace"].(map[string]any); !ok {
		t.Fatalf("expected trace result in JSON envelope: %s", buf.String())
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

func TestWriteWithOptionsRendersRunSummaryAndTrace(t *testing.T) {
	env := New("run")
	env.Macro = map[string]any{"name": "Main.Run", "duration_ms": 42}
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "saved": false}
	env.Target = map[string]any{"kind": "file", "path": "build/Book.xlsm", "description": "Saved workbook file on disk"}
	env.Session = map[string]any{"active": false, "workbook_path": "build/Book.xlsm", "dirty": false, "save_required": false}
	env.Trace = map[string]any{"events": []map[string]any{{"timestamp": "2026-04-30 10:00:00", "message": "start"}}}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Target:", "Saved workbook file on disk", "Session state:", "inactive", "Main.Run", "42ms", "left unchanged", "Trace", "start"} {
		if !strings.Contains(got, want) {
			t.Fatalf("run output missing %q:\n%s", want, got)
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

func TestWriteWithOptionsRendersRunTraceHelperLifecycle(t *testing.T) {
	tests := []struct {
		name  string
		trace map[string]any
		want  string
	}{
		{
			name: "temporary reverted",
			trace: map[string]any{
				"enabled":            true,
				"lifecycle":          "temporary",
				"temporary_reverted": true,
			},
			want: "temporary helper injected for this run and reverted afterward",
		},
		{
			name: "existing helper",
			trace: map[string]any{
				"enabled":   true,
				"lifecycle": "existing",
			},
			want: "using an existing workbook trace helper",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			env := New("run")
			env.Macro = map[string]any{"name": "Main.Run", "duration_ms": 42}
			env.Workbook = map[string]any{"path": "build/Book.xlsm", "saved": false}
			env.Trace = tt.trace
			var buf bytes.Buffer
			if err := WriteWithOptions(&buf, env, Options{}); err != nil {
				t.Fatal(err)
			}
			if got := buf.String(); !strings.Contains(got, tt.want) {
				t.Fatalf("run trace lifecycle output missing %q:\n%s", tt.want, got)
			}
		})
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
		"suggestion": "If you want source-controlled tracing, run `xlflow trace enable`.",
	}}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"VBA105", "src/modules/Main.bas:3", "Suggestion:", "xlflow trace enable"} {
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
	env.Hints = []map[string]any{{"code": "userform_planned_commands", "message": "Planned/future commands for deeper UserForm inspection include `xlflow form snapshot <name> --designer`."}}
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
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "saved": false, "session": true, "session_mode": "explicit", "needs_save": true}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"SAVE REQUIRED", "xlflow save before session stop", "explicit xlflow session workbook"} {
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
	env.Session = map[string]any{"running": true, "workbook_open": true, "needs_save": true}
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "needs_save": true}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Running:", "Workbook open:", "SAVE REQUIRED", "xlflow save before session stop"} {
		if !strings.Contains(got, want) {
			t.Fatalf("session status output missing %q:\n%s", want, got)
		}
	}
}

func TestWriteWithOptionsRendersTraceCommandSummary(t *testing.T) {
	env := New("trace")
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "saved": true, "session": true}
	env.Source = map[string]any{"path": "src/modules/XlflowTrace.bas", "updated": true}
	env.Trace = map[string]any{"lifecycle": "enabled", "log_dir": ".xlflow/traces"}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"saved live session workbook with trace helper", "persisted in workbook and source", "Trace dir"} {
		if !strings.Contains(got, want) {
			t.Fatalf("trace output missing %q:\n%s", want, got)
		}
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
	for _, want := range []string{"Snapshot target:", "live session workbook", "Form:", "UserForm1", "Basis:", "designer", "Coordinates:", "parent-relative", "Controls:", "3", "Output:", "artifacts/UserForm1.form.yaml", "Format:", "YAML", "SAVE REQUIRED", "save_required"} {
		if !strings.Contains(got, want) {
			t.Fatalf("form snapshot output missing %q:\n%s", want, got)
		}
	}
	if strings.Count(got, "Target:") != 1 {
		t.Fatalf("form snapshot output should render one Target label:\n%s", got)
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

func TestWriteWithOptionsDoesNotDuplicateTraceEventsFromLogs(t *testing.T) {
	env := New("run")
	env.Macro = map[string]any{"name": "Main.Run", "duration_ms": 42}
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "saved": false}
	env.Logs = []string{"ran Main.Run in 42ms", "[2026-04-30 10:00:00] start"}
	env.Trace = map[string]any{"events": []map[string]any{{"timestamp": "2026-04-30 10:00:00", "message": "start"}}}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	if strings.Count(got, "start") != 1 {
		t.Fatalf("expected trace event once:\n%s", got)
	}
	if !strings.Contains(got, "ran Main.Run in 42ms") {
		t.Fatalf("expected non-trace log to remain:\n%s", got)
	}
}

func TestWriteWithOptionsRendersTestFailures(t *testing.T) {
	env := Failure("test", Error{Code: "test_failed", Message: "1 of 2 test(s) failed"})
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
