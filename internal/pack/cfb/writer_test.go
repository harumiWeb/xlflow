package cfb

import (
	"bytes"
	"io"
	"strings"
	"testing"

	"github.com/richardlehane/mscfb"
)

// readBack parses the cfb output with mscfb into a path->contents map.
func readBack(t *testing.T, data []byte) map[string][]byte {
	t.Helper()
	doc, err := mscfb.New(bytes.NewReader(data))
	if err != nil {
		t.Fatalf("mscfb.New: %v", err)
	}
	out := map[string][]byte{}
	for entry, err := doc.Next(); err == nil; entry, err = doc.Next() {
		buf, rerr := io.ReadAll(entry)
		if rerr != nil {
			t.Fatalf("read %q: %v", entry.Name, rerr)
		}
		key := entry.Name
		if len(entry.Path) > 0 {
			key = strings.Join(entry.Path, "/") + "/" + entry.Name
		}
		out[key] = buf
	}
	return out
}

func TestSingleStreamRoundTrip(t *testing.T) {
	w := NewWriter()
	w.AddStream([]string{"PROJECT"}, []byte("hello"))
	data, err := w.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	got := readBack(t, data)
	if !bytes.Equal(got["PROJECT"], []byte("hello")) {
		t.Errorf("PROJECT = %q, want %q", got["PROJECT"], "hello")
	}
}

func TestNestedStreams(t *testing.T) {
	w := NewWriter()
	w.AddStream([]string{"PROJECT"}, bytes.Repeat([]byte("p"), 340))
	w.AddStream([]string{"VBA", "dir"}, bytes.Repeat([]byte("d"), 513))
	w.AddStream([]string{"VBA", "_VBA_PROJECT"}, []byte("1234567"))
	w.AddStream([]string{"VBA", "Spike"}, bytes.Repeat([]byte("s"), 141))
	data, err := w.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	got := readBack(t, data)
	want := map[string]int{
		"PROJECT": 340, "VBA/dir": 513, "VBA/_VBA_PROJECT": 7, "VBA/Spike": 141,
	}
	for k, n := range want {
		if len(got[k]) != n {
			t.Errorf("%s len = %d, want %d", k, len(got[k]), n)
		}
	}
}

func TestLargeStreamUsesFAT(t *testing.T) {
	w := NewWriter()
	w.AddStream([]string{"big"}, bytes.Repeat([]byte("x"), 5000))
	data, err := w.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	got := readBack(t, data)
	if len(got["big"]) != 5000 {
		t.Errorf("big len = %d, want 5000", len(got["big"]))
	}
}

// A name must fit in the 64B field (UTF-16 31 code units + NUL).
// Silently writing a longer name into a directory entry would corrupt adjacent
// fields, so it is rejected with an error.
func TestNameTooLong(t *testing.T) {
	w := NewWriter()
	w.AddStream([]string{strings.Repeat("a", 32)}, []byte("x"))
	if _, err := w.Bytes(); err == nil {
		t.Error("expected an error for a 32-character stream name, got nil")
	}

	// Exactly 31 characters is allowed.
	w2 := NewWriter()
	w2.AddStream([]string{strings.Repeat("a", 31)}, []byte("x"))
	if _, err := w2.Bytes(); err != nil {
		t.Errorf("31 characters should be allowed but errored: %v", err)
	}
}
