package cli

import (
	"encoding/json"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/filepull"
	"github.com/harumiWeb/xlflow/internal/output"
)

func TestPullFileBackendPublishesSavedWorkbookSource(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))

	stdout, err := runBuildCommandForTest(dir, "--json", "pull", "--backend", "file")
	if err != nil {
		t.Fatalf("pull --backend file: %v\n%s", err, stdout)
	}
	var env output.Envelope
	if err := json.Unmarshal([]byte(stdout), &env); err != nil {
		t.Fatal(err)
	}
	pull := cliObjectMap(env.Pull)
	if pull["backend"] != "file" || pull["source"] != "saved_workbook" {
		t.Fatalf("pull authority = %#v", pull)
	}
	if target := cliObjectMap(env.Target); target["kind"] != "file" {
		t.Fatalf("target = %#v, want kind=file", target)
	}
	for _, path := range []string{
		filepath.Join(dir, "src", "modules", "Module1.bas"),
		filepath.Join(dir, "src", "classes", "Class1.cls"),
		filepath.Join(dir, "src", "workbook", "Sheet1.bas"),
		filepath.Join(dir, "src", "workbook", "ThisWorkbook.bas"),
	} {
		body, readErr := os.ReadFile(path)
		if readErr != nil || strings.HasPrefix(string(body), "\xef\xbb\xbf") || !strings.HasSuffix(string(body), "\n") {
			t.Fatalf("source %s was not published as UTF-8 without BOM and final newline: %v", path, readErr)
		}
	}
}

func TestPullFileBackendWarnsWhenMatchingSessionIsIgnored(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p1_compiled.bin"))
	if err := os.MkdirAll(filepath.Join(dir, ".xlflow"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, ".xlflow", "session.json"), []byte(`{"pid":123,"workbook_path":"build/Book.xlsm"}`), 0o644); err != nil {
		t.Fatal(err)
	}

	stdout, err := runBuildCommandForTest(dir, "--json", "pull", "--backend", "file")
	if err != nil {
		t.Fatal(err)
	}
	if !jsonHasWarningCode(t, stdout, "file_pull_live_session_ignored") {
		t.Fatalf("missing live-session warning: %s", stdout)
	}
}

func TestPullFileBackendRejectsUserFormBeforeMutation(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	writePackTemplate(t, dir, readPackFixture(t, "testdata", "corpus", "p4_form.bin"))
	existing := filepath.Join(dir, "src", "modules", "keep.bas")
	if err := os.WriteFile(existing, []byte("keep"), 0o644); err != nil {
		t.Fatal(err)
	}

	stdout, err := runBuildCommandForTest(dir, "--json", "pull", "--backend", "file")
	if err == nil || output.ExitCode(err) != output.ExitValidation {
		t.Fatalf("error = %v, exit = %d; want validation failure", err, output.ExitCode(err))
	}
	if jsonErrorCode(t, stdout) != "pull_userform_unsupported" {
		t.Fatalf("unexpected failure: %s", stdout)
	}
	body, readErr := os.ReadFile(existing)
	if readErr != nil || string(body) != "keep" {
		t.Fatalf("source mutated after rejection: %q, %v", body, readErr)
	}
}

func TestPullFileBackendRejectsSessionFlag(t *testing.T) {
	dir := t.TempDir()
	writePackConfig(t, dir)
	stdout, err := runBuildCommandForTest(dir, "--json", "pull", "--backend", "file", "--session")
	if err == nil || output.ExitCode(err) != output.ExitConfig || jsonErrorCode(t, stdout) != "pull_args_invalid" {
		t.Fatalf("stdout=%s err=%v", stdout, err)
	}
}

func TestFilePullPublicationFailureKeepsPublicationErrorCode(t *testing.T) {
	err := errors.Join(filepull.ErrPublish, &os.PathError{Op: "write", Path: "src/modules/Main.bas", Err: os.ErrPermission})
	if got := filePullErrorCode(err); got != "pull_source_publish_failed" {
		t.Fatalf("filePullErrorCode() = %q", got)
	}
}

func jsonErrorCode(t *testing.T, body string) string {
	t.Helper()
	var envelope struct {
		Error *output.Error `json:"error"`
	}
	if err := json.Unmarshal([]byte(body), &envelope); err != nil {
		t.Fatal(err)
	}
	if envelope.Error == nil {
		return ""
	}
	return envelope.Error.Code
}

func jsonHasWarningCode(t *testing.T, body, code string) bool {
	t.Helper()
	var envelope struct {
		Warnings []struct {
			Code string `json:"code"`
		} `json:"warnings"`
	}
	if err := json.Unmarshal([]byte(body), &envelope); err != nil {
		t.Fatal(err)
	}
	for _, warning := range envelope.Warnings {
		if warning.Code == code {
			return true
		}
	}
	return false
}
