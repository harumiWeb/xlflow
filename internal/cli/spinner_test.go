package cli

import (
	"bytes"
	"errors"
	"strings"
	"testing"

	tea "github.com/charmbracelet/bubbletea"
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
