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
