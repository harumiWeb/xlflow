//go:build windows

package excel

import (
	"bytes"
	"strings"
	"testing"
	"time"

	winio "github.com/Microsoft/go-winio"
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
	timeout := 2 * time.Second
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
	if err := session.Close(); err != nil {
		t.Fatal(err)
	}
	events := session.Events()
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
