package cfb

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"io"
	"slices"
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

func TestStorageMetadataRoundTrip(t *testing.T) {
	root := StorageMeta{StateBits: 1, Modified: 0x01DCFD81AB636A60}
	form := StorageMeta{
		CLSID:     [16]byte{0x20, 0x20, 0x18, 0x6e, 0x60, 0xf4, 0xce, 0x11, 0x9b, 0xcd, 0x00, 0xaa, 0x00, 0x60, 0x8e, 0x01},
		StateBits: 2,
		Created:   0x01DCFD81AB631C40,
		Modified:  0x01DCFD81AB636A60,
	}
	empty := StorageMeta{CLSID: [16]byte{1, 2, 3}, Modified: 42}

	for _, format := range []Format{FormatV3, FormatV4} {
		t.Run(fmt.Sprintf("v%d", format), func(t *testing.T) {
			w, err := NewWriterForFormat(format)
			if err != nil {
				t.Fatal(err)
			}
			w.AddStorage(nil, root)
			w.AddStorage([]string{"UserForm1", "i02"}, form)
			w.AddStorage([]string{"UserForm1", "empty"}, empty)
			w.AddStream([]string{"UserForm1", "i02", "f"}, []byte("form"))

			data, err := w.Bytes()
			if err != nil {
				t.Fatal(err)
			}
			got, err := Open(data)
			if err != nil {
				t.Fatal(err)
			}
			want := map[string]StorageMeta{
				"":                root,
				"UserForm1":       {},
				"UserForm1/empty": empty,
				"UserForm1/i02":   form,
			}
			if !slices.Equal(got.StoragePaths(), []string{"", "UserForm1", "UserForm1/i02", "UserForm1/empty"}) {
				t.Fatalf("storage paths = %q", got.StoragePaths())
			}
			for path, wantMeta := range want {
				gotMeta, ok := got.Storage(path)
				if !ok || gotMeta != wantMeta {
					t.Errorf("Storage(%q) = (%+v, %v), want (%+v, true)", path, gotMeta, ok, wantMeta)
				}
			}
		})
	}
}

func TestWriterStorageDefinitionConflicts(t *testing.T) {
	meta := StorageMeta{Modified: 1}
	w := NewWriter()
	w.AddStorage([]string{"form"}, meta)
	w.AddStorage([]string{"form"}, meta)
	w.AddStream([]string{"form", "f"}, []byte("x"))
	if _, err := w.Bytes(); err != nil {
		t.Fatalf("identical storage definitions: %v", err)
	}

	w = NewWriter()
	w.AddStorage([]string{"form"}, meta)
	w.AddStorage([]string{"form"}, StorageMeta{Modified: 2})
	if _, err := w.Bytes(); err == nil || !strings.Contains(err.Error(), "conflicting metadata") {
		t.Fatalf("conflicting storage metadata error = %v", err)
	}

	w = NewWriter()
	w.AddStream([]string{"form"}, []byte("x"))
	w.AddStorage([]string{"form"}, meta)
	if _, err := w.Bytes(); err == nil || !strings.Contains(err.Error(), "storage") {
		t.Fatalf("stream/storage collision error = %v", err)
	}
}

func TestWriterRejectsInvalidRootMetadata(t *testing.T) {
	w := NewWriter()
	w.AddStorage(nil, StorageMeta{Created: 1})
	if _, err := w.Bytes(); err == nil || !strings.Contains(err.Error(), "root storage") {
		t.Fatalf("root creation FILETIME error = %v", err)
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
