//go:build windows

package excel

import (
	"bytes"
	"strings"
	"testing"
	"time"

	winio "github.com/Microsoft/go-winio"
)

const (
	uiStreamTestPipeTimeout  = 10 * time.Second
	uiStreamTestEventTimeout = 10 * time.Second
	uiStreamTestCloseTimeout = 10 * time.Second
)

func TestFormatUIStreamEventRedactsInputValue(t *testing.T) {
	got := formatUIStreamEvent(map[string]any{"kind": "inputbox", "dialog_id": "customer-name", "response_source": "default", "resolved_value": "Alice", "redacted": true})
	for _, want := range []string{"xlflow: ui", "kind=inputbox", "id=customer-name", "source=default", "value=[redacted]"} {
		if !strings.Contains(got, want) {
			t.Fatalf("formatUIStreamEvent() = %q, want substring %q", got, want)
		}
	}
}

func TestFormatUIStreamEventEscapesControlCharacters(t *testing.T) {
	got := formatUIStreamEvent(map[string]any{"kind": "inputbox", "dialog_id": "customer\nname", "response_source": "default", "resolved_value": "Alice\tBob"})
	for _, want := range []string{"id=customer\\nname", "value=Alice\\tBob"} {
		if !strings.Contains(got, want) {
			t.Fatalf("formatUIStreamEvent() = %q, want substring %q", got, want)
		}
	}
}

func TestUIStreamSessionCollectsNamedPipeEvents(t *testing.T) {
	var stderr bytes.Buffer
	session, err := newUIStreamSession(&stderr)
	if err != nil {
		t.Fatal(err)
	}
	timeout := uiStreamTestPipeTimeout
	conn, err := winio.DialPipe(session.PipePath(), &timeout)
	if err != nil {
		_ = session.Close()
		t.Fatal(err)
	}
	if _, err := conn.Write([]byte("{\"kind\":\"msgbox\",\"dialog_id\":\"confirm-save\",\"response_source\":\"scripted\",\"resolved_result\":\"yes\"}\n")); err != nil {
		_ = conn.Close()
		_ = session.Close()
		t.Fatal(err)
	}
	_ = conn.Close()
	events := waitForUIStreamEvents(session, 1, uiStreamTestEventTimeout)
	if len(events) != 1 {
		closeErr := session.Close()
		events = session.Events()
		if len(events) == 1 && closeErr == nil {
			t.Log("UI stream event was collected only after session close")
		} else {
			t.Fatalf("events = %#v after close error %v, want 1 event", events, closeErr)
		}
	} else if err := session.Close(); err != nil {
		t.Fatal(err)
	}
	if len(events) != 1 {
		t.Fatalf("events = %#v, want 1 event", events)
	}
	if events[0]["dialog_id"] != "confirm-save" {
		t.Fatalf("dialog_id = %#v, want confirm-save", events[0]["dialog_id"])
	}
	if !strings.Contains(stderr.String(), "id=confirm-save") {
		t.Fatalf("stderr = %q, want rendered event", stderr.String())
	}
}

func TestUIStreamSessionCloseDoesNotHangOnOpenConnection(t *testing.T) {
	var stderr bytes.Buffer
	session, err := newUIStreamSession(&stderr)
	if err != nil {
		t.Fatal(err)
	}
	timeout := uiStreamTestPipeTimeout
	conn, err := winio.DialPipe(session.PipePath(), &timeout)
	if err != nil {
		_ = session.Close()
		t.Fatal(err)
	}
	if _, err := conn.Write([]byte("{\"kind\":\"msgbox\",\"dialog_id\":\"hang-check\",\"response_source\":\"scripted\",\"resolved_result\":\"yes\"}\n")); err != nil {
		_ = conn.Close()
		_ = session.Close()
		t.Fatal(err)
	}
	events := waitForUIStreamEvents(session, 1, uiStreamTestEventTimeout)
	if len(events) != 1 {
		_ = conn.Close()
		_ = session.Close()
		t.Fatalf("events = %#v, want 1 event", events)
	}
	closed := make(chan error, 1)
	go func() {
		closed <- session.Close()
	}()
	select {
	case err := <-closed:
		if err != nil {
			t.Fatal(err)
		}
	case <-time.After(uiStreamTestCloseTimeout):
		_ = conn.Close()
		t.Fatal("session.Close() timed out while the pipe client remained open")
	}
	if err := conn.Close(); err != nil {
		t.Fatal(err)
	}
	if events[0]["dialog_id"] != "hang-check" {
		t.Fatalf("dialog_id = %#v, want hang-check", events[0]["dialog_id"])
	}
}

func TestUIStreamSessionRejectsOversizedPendingLine(t *testing.T) {
	var stderr bytes.Buffer
	session, err := newUIStreamSession(&stderr)
	if err != nil {
		t.Fatal(err)
	}
	timeout := uiStreamTestPipeTimeout
	conn, err := winio.DialPipe(session.PipePath(), &timeout)
	if err != nil {
		_ = session.Close()
		t.Fatal(err)
	}
	if _, err := conn.Write([]byte(strings.Repeat("x", uiStreamMaxPendingBytes+1))); err != nil {
		_ = conn.Close()
		_ = session.Close()
		t.Fatal(err)
	}
	if err := conn.Close(); err != nil {
		t.Fatal(err)
	}
	if err := session.Close(); err == nil || !strings.Contains(err.Error(), "ui stream message exceeds") {
		t.Fatalf("session.Close() error = %v, want oversized line error", err)
	}
}

func waitForUIStreamEvents(session *uiStreamSession, want int, timeout time.Duration) []map[string]any {
	deadline := time.Now().Add(timeout)
	for time.Now().Before(deadline) {
		events := session.Events()
		if len(events) >= want {
			return events
		}
		time.Sleep(10 * time.Millisecond)
	}
	return session.Events()
}
