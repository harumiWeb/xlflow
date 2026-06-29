//go:build windows

package excel

import (
	"bytes"
	"strings"
	"testing"
	"time"

	winio "github.com/Microsoft/go-winio"
)

func TestFormatDebugStreamEventEscapesControlCharacters(t *testing.T) {
	got := formatDebugStreamEvent(map[string]any{"message": "hello\nworld", "runtime_mode": "headless", "source": "XlflowDebug.Log"})
	for _, want := range []string{"xlflow: debug", "source=XlflowDebug.Log", "mode=headless", `message=hello\nworld`} {
		if !strings.Contains(got, want) {
			t.Fatalf("formatDebugStreamEvent() = %q, want substring %q", got, want)
		}
	}
}

func TestDebugStreamSessionCollectsNamedPipeEvents(t *testing.T) {
	var stderr bytes.Buffer
	session, err := newDebugStreamSession(&stderr)
	if err != nil {
		t.Fatal(err)
	}
	timeout := uiStreamTestPipeTimeout
	conn, err := winio.DialPipe(session.PipePath(), &timeout)
	if err != nil {
		_ = session.Close()
		t.Fatal(err)
	}
	if _, err := conn.Write([]byte(`{"message":"first line","runtime_mode":"headless","source":"XlflowDebug.Log"}` + "\n")); err != nil {
		_ = conn.Close()
		_ = session.Close()
		t.Fatal(err)
	}
	_ = conn.Close()
	result := waitForDebugStreamResult(session, 1, uiStreamTestEventTimeout)
	resultMap, ok := result.(map[string]any)
	if !ok {
		closeErr := session.Close()
		result = session.Result()
		resultMap, ok = result.(map[string]any)
		if !ok {
			t.Fatalf("result = %#v after close error %v, want map", result, closeErr)
		}
		t.Log("debug stream event was collected only after session close")
	} else if err := session.Close(); err != nil {
		t.Fatal(err)
	}
	if count, ok := resultMap["count"].(int); ok {
		if count != 1 {
			t.Fatalf("count = %d, want 1", count)
		}
	} else if count, ok := resultMap["count"].(float64); !ok || int(count) != 1 {
		t.Fatalf("count = %#v, want 1", resultMap["count"])
	}
	if !strings.Contains(stderr.String(), "message=first line") {
		t.Fatalf("stderr = %q, want rendered debug event", stderr.String())
	}
}

func TestDebugStreamSessionTracksTruncation(t *testing.T) {
	var stderr bytes.Buffer
	session, err := newDebugStreamSession(&stderr)
	if err != nil {
		t.Fatal(err)
	}
	timeout := uiStreamTestPipeTimeout
	conn, err := winio.DialPipe(session.PipePath(), &timeout)
	if err != nil {
		_ = session.Close()
		t.Fatal(err)
	}
	for i := 0; i < debugStreamMaxEvents+5; i++ {
		line := []byte(`{"message":"line","runtime_mode":"headless","source":"XlflowDebug.Log"}` + "\n")
		if _, err := conn.Write(line); err != nil {
			_ = conn.Close()
			_ = session.Close()
			t.Fatal(err)
		}
	}
	if err := conn.Close(); err != nil {
		t.Fatal(err)
	}
	if err := session.Close(); err != nil {
		t.Fatal(err)
	}
	resultMap, ok := session.Result().(map[string]any)
	if !ok {
		t.Fatalf("result = %#v, want map", session.Result())
	}
	if truncated, ok := resultMap["truncated"].(bool); !ok || !truncated {
		t.Fatalf("truncated = %#v, want true", resultMap["truncated"])
	}
	events := resultMap["events"].([]map[string]any)
	if len(events) != debugStreamMaxEvents {
		t.Fatalf("events length = %d, want %d", len(events), debugStreamMaxEvents)
	}
}

func waitForDebugStreamResult(session *debugStreamSession, want int, timeout time.Duration) any {
	deadline := time.Now().Add(timeout)
	for time.Now().Before(deadline) {
		resultMap, ok := session.Result().(map[string]any)
		if ok {
			events, _ := resultMap["events"].([]map[string]any)
			if len(events) >= want {
				return resultMap
			}
		}
		time.Sleep(10 * time.Millisecond)
	}
	return session.Result()
}
