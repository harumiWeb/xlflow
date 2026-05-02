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
	env.Trace = map[string]any{"events": []map[string]any{{"timestamp": "2026-04-30 10:00:00", "message": "start"}}}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"Main.Run", "42ms", "left unchanged", "Trace", "start"} {
		if !strings.Contains(got, want) {
			t.Fatalf("run output missing %q:\n%s", want, got)
		}
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
	env.Workbook = map[string]any{"path": "build/Book.xlsm", "saved": false, "session": true}
	env.Source = map[string]any{"changed_only": true, "changed": true}
	var buf bytes.Buffer
	if err := WriteWithOptions(&buf, env, Options{}); err != nil {
		t.Fatal(err)
	}
	got := buf.String()
	for _, want := range []string{"updated live session workbook only", "xlflow save --session"} {
		if !strings.Contains(got, want) {
			t.Fatalf("push output missing %q:\n%s", want, got)
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
