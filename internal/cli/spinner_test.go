package cli

import (
	"bytes"
	"errors"
	"strings"
	"testing"

	tea "github.com/charmbracelet/bubbletea"
	"github.com/harumiWeb/xlflow/internal/excel"
)

func TestSpinnerModelQuitsOnDone(t *testing.T) {
	done := make(chan error, 1)
	wantErr := errors.New("boom")
	done <- wantErr
	model := newSpinnerModel("Running macro", done)
	updated, cmd := model.Update(spinnerDoneMsg{err: wantErr})
	if cmd == nil {
		t.Fatal("expected quit command")
	}
	if _, ok := cmd().(tea.QuitMsg); !ok {
		t.Fatal("expected quit message")
	}
	spinner, ok := updated.(spinnerModel)
	if !ok {
		t.Fatalf("updated model = %T", updated)
	}
	if !errors.Is(spinner.err, wantErr) {
		t.Fatalf("spinner err = %v, want %v", spinner.err, wantErr)
	}
}

func TestRunSpinnerReturnsWorkError(t *testing.T) {
	wantErr := errors.New("work failed")
	var buf bytes.Buffer
	err := runSpinner(&buf, "Running macro", func() error {
		return wantErr
	})
	if !errors.Is(err, wantErr) {
		t.Fatalf("runSpinner err = %v, want %v", err, wantErr)
	}
	if !strings.Contains(buf.String(), "Running macro") {
		t.Fatalf("spinner output = %q", buf.String())
	}
}

func TestWithSpinnerWritesSingleLineProgressForNonInteractive(t *testing.T) {
	var stderr bytes.Buffer
	a := &app{
		json:           true,
		stderr:         &stderr,
		stdoutTerminal: func() bool { return false },
		stderrTerminal: func() bool { return false },
	}

	err := a.withSpinner("Running macro", func() error {
		return nil
	})
	if err != nil {
		t.Fatalf("withSpinner err = %v", err)
	}
	if got := stderr.String(); got != "xlflow: Running macro...\n" {
		t.Fatalf("progress output = %q", got)
	}
}

func TestWithExcelProgressSkipsWhenProgressDisabled(t *testing.T) {
	var stderr bytes.Buffer
	a := &app{stderr: &stderr}

	err := a.withExcelProgress("Running macro", excel.CommandOptions{Stderr: &stderr, Progress: false}, func() error {
		return nil
	})
	if err != nil {
		t.Fatalf("withExcelProgress err = %v", err)
	}
	if got := stderr.String(); got != "" {
		t.Fatalf("expected no progress output, got %q", got)
	}
}
