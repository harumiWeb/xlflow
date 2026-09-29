package cfb

import (
	"bytes"
	"encoding/binary"
	"io"
	"strings"
	"testing"

	"github.com/richardlehane/mscfb"
)

func TestVersion4RoundTrip(t *testing.T) {
	w, err := NewWriterForFormat(FormatV4)
	if err != nil {
		t.Fatal(err)
	}
	w.AddStream([]string{"PROJECT"}, []byte("v4 project"))
	w.AddStream([]string{"VBA", "large"}, bytes.Repeat([]byte("v4"), 2000))
	data, err := w.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	if got := binary.LittleEndian.Uint16(data[26:28]); got != uint16(FormatV4) {
		t.Fatalf("major version = %d, want 4", got)
	}
	if got := binary.LittleEndian.Uint16(data[30:32]); got != 12 {
		t.Fatalf("sector shift = %d, want 12", got)
	}
	if got := binary.LittleEndian.Uint32(data[40:44]); got == 0 {
		t.Fatal("v4 directory sector count is zero")
	}
	if len(data)%4096 != 0 {
		t.Fatalf("v4 file size %d is not 4096-byte aligned", len(data))
	}
	c, err := Open(data)
	if err != nil {
		t.Fatalf("Open(v4): %v", err)
	}
	if c.Format() != FormatV4 {
		t.Fatalf("format = %d, want v4", c.Format())
	}
	got := readBack(t, data)
	if !bytes.Equal(got["PROJECT"], []byte("v4 project")) || len(got["VBA/large"]) != 4000 {
		t.Fatalf("external v4 read-back mismatch: PROJECT=%q large=%d", got["PROJECT"], len(got["VBA/large"]))
	}
}

func TestWriterRejectsForbiddenNameCharacters(t *testing.T) {
	for _, forbidden := range []rune{0, '/', '\\', ':', '!'} {
		w := NewWriter()
		w.AddStream([]string{"bad" + string(forbidden) + "name"}, []byte("x"))
		if _, err := w.Bytes(); err == nil || !strings.Contains(err.Error(), "forbidden character") {
			t.Fatalf("Bytes(%q) error = %v, want forbidden-character rejection", forbidden, err)
		}
	}
	w := NewWriter()
	w.AddStream([]string{""}, []byte("x"))
	if _, err := w.Bytes(); err == nil || !strings.Contains(err.Error(), "must not be empty") {
		t.Fatalf("Bytes(empty name) error = %v, want empty-name rejection", err)
	}
}

func TestWriterEmitsDIFATBeyondHeaderEntries(t *testing.T) {
	w := NewWriter()
	payload := bytes.Repeat([]byte{0xA5}, 8<<20)
	w.AddStream([]string{"large"}, payload)
	data, err := w.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	if got := binary.LittleEndian.Uint32(data[44:48]); got <= difatHeaderLen {
		t.Fatalf("FAT sector count = %d, want more than %d", got, difatHeaderLen)
	}
	if got := binary.LittleEndian.Uint32(data[72:76]); got == 0 {
		t.Fatal("DIFAT sector count is zero")
	}
	if got := binary.LittleEndian.Uint32(data[68:72]); got == endOfChain {
		t.Fatal("first DIFAT sector is ENDOFCHAIN")
	}
	c, err := Open(data)
	if err != nil {
		t.Fatalf("Open(DIFAT output): %v", err)
	}
	stream, ok := c.Stream("large")
	if !ok || !bytes.Equal(stream, payload) {
		t.Fatalf("internal DIFAT read-back mismatch: ok=%v len=%d", ok, len(stream))
	}
	got := readBack(t, data)
	if !bytes.Equal(got["large"], payload) {
		t.Fatalf("external DIFAT read-back mismatch: len=%d", len(got["large"]))
	}
}

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
