package output

import (
	"bytes"
	"encoding/json"
	"errors"
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
