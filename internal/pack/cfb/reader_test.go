package cfb

import (
	"bytes"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"testing"

	"github.com/richardlehane/mscfb"
)

func loadBin(t *testing.T, book string) []byte {
	t.Helper()
	b, err := os.ReadFile(filepath.Join("testdata", "corpus", book+".bin"))
	if err != nil {
		t.Fatal(err)
	}
	return b
}

func TestOpenRejectsBadSignature(t *testing.T) {
	// Inputs shorter than 512B are rejected by the length check.
	if _, err := Open([]byte("too short")); err == nil {
		t.Error("Open should return an error for inputs shorter than 512B")
	}
	// At least 512B but with a bad signature -> rejected by the signature branch (passes the length check to exercise it).
	bad := make([]byte, headerSize)
	copy(bad, []byte("XXXXXXXX"))
	if _, err := Open(bad); err == nil {
		t.Error("Open should return an error for a bad signature")
	}
}

func TestOpenAcceptsRealBin(t *testing.T) {
	if _, err := Open(loadBin(t, "p1_compiled")); err != nil {
		t.Fatalf("Open(p1_compiled) failed: %v", err)
	}
}

func TestContainerHasAllStreams(t *testing.T) {
	c, err := Open(loadBin(t, "p1_compiled"))
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{"PROJECT", "VBA/dir", "VBA/_VBA_PROJECT", "VBA/Module1", "VBA/Class1"} {
		if _, ok := c.Stream(want); !ok {
			t.Errorf("stream %q is missing", want)
		}
	}
	// dir begins with 0x01 (the CompressedContainer SignatureByte).
	if d, _ := c.Stream("VBA/dir"); len(d) == 0 || d[0] != 0x01 {
		t.Errorf("VBA/dir does not begin with 0x01")
	}
}

func TestNestedStorageStream(t *testing.T) {
	c, err := Open(loadBin(t, "p4_form"))
	if err != nil {
		t.Fatal(err)
	}
	// A form has a nested storage.
	if _, ok := c.Stream("UserForm1/\x03VBFrame"); !ok {
		t.Errorf("nested storage UserForm1/\\x03VBFrame was not read")
	}
}

func TestPathsMatchGolden(t *testing.T) {
	for _, book := range []string{"p1_compiled", "p2_refs", "p3_protected", "p4_form", "p5_mbcs"} {
		c, err := Open(loadBin(t, book))
		if err != nil {
			t.Fatalf("%s: %v", book, err)
		}
		want, rerr := os.ReadFile(filepath.Join("testdata", "corpus", book+".streams"))
		if rerr != nil {
			t.Fatalf("%s: failed to read golden streams: %v", book, rerr)
		}
		// Golden files are committed with LF, but Windows checkouts with
		// core.autocrlf=true rewrite them to CRLF. TrimSpace only trims the
		// whole blob, so normalize CRLF before splitting to keep the per-line
		// comparison platform-independent.
		normalized := strings.ReplaceAll(string(want), "\r\n", "\n")
		wantPaths := strings.Split(strings.TrimSpace(normalized), "\n")
		got := append([]string{}, c.Paths()...)
		sort.Strings(got)
		sort.Strings(wantPaths)
		if strings.Join(got, "|") != strings.Join(wantPaths, "|") {
			t.Errorf("%s: paths\n got=%v\nwant=%v", book, got, wantPaths)
		}
	}
}

// lookupStream looks up the Container by a key originating from mscfb.
// Because mscfb strips leading control characters (0x01-0x1F) from entry names,
// if a direct hit fails it tries alternate keys with a control character prepended.
func lookupStream(c *Container, key string) ([]byte, bool) {
	if d, ok := c.Stream(key); ok {
		return d, ok
	}
	// Find the last path separator and try prepending a control character to the final component.
	slash := strings.LastIndexByte(key, '/')
	name := key
	prefix := ""
	if slash >= 0 {
		prefix = key[:slash+1]
		name = key[slash+1:]
	}
	for cp := byte(1); cp < 0x20; cp++ {
		if d, ok := c.Stream(prefix + string(cp) + name); ok {
			return d, ok
		}
	}
	return nil, false
}

func TestContentMatchesMscfb(t *testing.T) {
	for _, book := range []string{"p1_compiled", "p4_form"} {
		data := loadBin(t, book)
		c, err := Open(data)
		if err != nil {
			t.Fatalf("%s: Open failed: %v", book, err)
		}
		doc, err := mscfb.New(bytes.NewReader(data))
		if err != nil {
			t.Fatal(err)
		}
		for entry, e := doc.Next(); e == nil; entry, e = doc.Next() {
			if entry.Size == 0 {
				continue
			}
			buf := make([]byte, entry.Size)
			n, _ := entry.Read(buf)
			key := entry.Name
			if len(entry.Path) > 0 {
				key = strings.Join(entry.Path, "/") + "/" + entry.Name
			}
			got, ok := lookupStream(c, key)
			if !ok || !bytes.Equal(got, buf[:n]) {
				t.Errorf("%s: stream %q does not match mscfb (ok=%v)", book, key, ok)
			}
		}
	}
}
