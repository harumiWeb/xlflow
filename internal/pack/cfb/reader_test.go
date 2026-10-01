package cfb

import (
	"bytes"
	"encoding/binary"
	"io"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"testing"

	"github.com/richardlehane/mscfb"
)

func TestOpenRejectsDeclaredCountsBeforeAllocation(t *testing.T) {
	w := NewWriter()
	w.AddStream([]string{"PROJECT"}, []byte("x"))
	data, err := w.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	binary.LittleEndian.PutUint32(data[44:48], ^uint32(0))
	if _, err := Open(data); err == nil || !strings.Contains(err.Error(), "exceeds physical sector count") {
		t.Fatalf("Open error = %v, want physical-sector count rejection", err)
	}
}

func TestOpenRejectsFATThatDoesNotCoverPhysicalSectors(t *testing.T) {
	w := NewWriter()
	w.AddStream([]string{"PROJECT"}, []byte("x"))
	data, err := w.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	g, err := geometryFor(FormatV3)
	if err != nil {
		t.Fatal(err)
	}
	physical := (len(data) - g.headerSpan) / g.sectorSize
	data = append(data, make([]byte, (g.entriesPerFatSector+1-physical)*g.sectorSize)...)
	if _, err := Open(data); err == nil || !strings.Contains(err.Error(), "FAT covers") {
		t.Fatalf("Open error = %v, want insufficient FAT coverage rejection", err)
	}
}

func TestOpenRejectsOversizedDeclaredRegularChainBeforeAllocation(t *testing.T) {
	w := NewWriter()
	w.AddStream([]string{"large"}, bytes.Repeat([]byte("x"), 5000))
	data, err := w.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	firstDir := binary.LittleEndian.Uint32(data[48:52])
	dirOffset := headerSize + int(firstDir)*512
	binary.LittleEndian.PutUint64(data[dirOffset+dirEntrySize+120:dirOffset+dirEntrySize+128], uint64(^uint32(0)))
	if _, err := Open(data); err == nil || !strings.Contains(err.Error(), "physical sectors") {
		t.Fatalf("Open error = %v, want oversized regular-chain rejection", err)
	}
}

func TestOpenRejectsOversizedV4DirectoryCountBeforeAllocation(t *testing.T) {
	w, err := NewWriterForFormat(FormatV4)
	if err != nil {
		t.Fatal(err)
	}
	w.AddStream([]string{"PROJECT"}, []byte("x"))
	data, err := w.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	physical := (len(data) - 4096) / 4096
	binary.LittleEndian.PutUint32(data[40:44], uint32(physical+1))
	if _, err := Open(data); err == nil || !strings.Contains(err.Error(), "physical sectors") {
		t.Fatalf("Open error = %v, want oversized directory-chain rejection", err)
	}
}

func TestOpenRejectsCyclicDirectoryChain(t *testing.T) {
	w := NewWriter()
	w.AddStream([]string{"PROJECT"}, bytes.Repeat([]byte("x"), 5000))
	data, err := w.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	firstDir := binary.LittleEndian.Uint32(data[48:52])
	firstFat := binary.LittleEndian.Uint32(data[76:80])
	fatOffset := headerSize + int(firstFat)*512 + int(firstDir)*4
	binary.LittleEndian.PutUint32(data[fatOffset:fatOffset+4], firstDir)
	if _, err := Open(data); err == nil || !strings.Contains(err.Error(), "shared by directory") {
		t.Fatalf("Open error = %v, want cyclic directory-chain rejection", err)
	}
}

func TestOpenRejectsCyclicMiniFATChain(t *testing.T) {
	w := NewWriter()
	w.AddStream([]string{"small"}, bytes.Repeat([]byte("x"), 128))
	data, err := w.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	firstMiniFAT := binary.LittleEndian.Uint32(data[60:64])
	miniFATOffset := headerSize + int(firstMiniFAT)*512
	binary.LittleEndian.PutUint32(data[miniFATOffset:miniFATOffset+4], 0)
	if _, err := Open(data); err == nil || !strings.Contains(err.Error(), "mini-sector 0 is shared") {
		t.Fatalf("Open error = %v, want cyclic mini-FAT rejection", err)
	}
}

func TestOpenRejectsCyclicDIFATChain(t *testing.T) {
	w := NewWriter()
	w.AddStream([]string{"large"}, bytes.Repeat([]byte{0xA5}, 8<<20))
	data, err := w.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	firstDIFAT := binary.LittleEndian.Uint32(data[68:72])
	if firstDIFAT == endOfChain {
		t.Fatal("fixture did not produce a DIFAT sector")
	}
	g, err := geometryFor(FormatV3)
	if err != nil {
		t.Fatal(err)
	}
	difatOffset := headerSize + int(firstDIFAT)*512
	binary.LittleEndian.PutUint32(data[difatOffset+g.entriesPerDifatSector*4:], firstDIFAT)
	if _, err := Open(data); err == nil || !strings.Contains(err.Error(), "DIFAT chain ends") {
		t.Fatalf("Open error = %v, want cyclic DIFAT rejection", err)
	}
}

func TestOpenRejectsVersionGeometryMismatch(t *testing.T) {
	w := NewWriter()
	w.AddStream([]string{"PROJECT"}, []byte("x"))
	data, err := w.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	binary.LittleEndian.PutUint16(data[26:28], uint16(FormatV4))
	if _, err := Open(data); err == nil || !strings.Contains(err.Error(), "4096-byte sector geometry") {
		t.Fatalf("Open error = %v, want v4 geometry rejection", err)
	}
}

func TestOpenRejectsOutOfRangeDirectoryReference(t *testing.T) {
	w := NewWriter()
	w.AddStream([]string{"PROJECT"}, []byte("x"))
	data, err := w.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	firstDir := binary.LittleEndian.Uint32(data[48:52])
	dirOffset := headerSize + int(firstDir)*512
	binary.LittleEndian.PutUint32(data[dirOffset+76:dirOffset+80], ^uint32(0)-1)
	if _, err := Open(data); err == nil || !strings.Contains(err.Error(), "directory reference") {
		t.Fatalf("Open error = %v, want out-of-range directory-reference rejection", err)
	}
}

func TestParseDirEntryRejectsForbiddenNameCharacters(t *testing.T) {
	for _, forbidden := range []uint16{0, '/', '\\', ':', '!'} {
		raw := make([]byte, dirEntrySize)
		binary.LittleEndian.PutUint16(raw[0:2], forbidden)
		binary.LittleEndian.PutUint16(raw[2:4], 0)
		binary.LittleEndian.PutUint16(raw[64:66], 4)
		raw[66] = objStream
		if _, err := parseDirEntry(raw, FormatV3); err == nil || !strings.Contains(err.Error(), "forbidden character") {
			t.Fatalf("parseDirEntry(%q) error = %v, want forbidden-character rejection", rune(forbidden), err)
		}
	}
}

func TestOpenRejectsMetadataForbiddenForObjectType(t *testing.T) {
	w := NewWriter()
	w.AddStream([]string{"PROJECT"}, []byte("x"))
	data, err := w.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	firstDir := binary.LittleEndian.Uint32(data[48:52])
	dirOffset := headerSize + int(firstDir)*512

	nonzeroStreamCLSID := bytes.Clone(data)
	nonzeroStreamCLSID[dirOffset+dirEntrySize+80] = 1
	if _, err := Open(nonzeroStreamCLSID); err == nil || !strings.Contains(err.Error(), "nonzero CLSID") {
		t.Fatalf("stream CLSID error = %v", err)
	}

	nonzeroStreamTime := bytes.Clone(data)
	binary.LittleEndian.PutUint64(nonzeroStreamTime[dirOffset+dirEntrySize+100:], 1)
	if _, err := Open(nonzeroStreamTime); err == nil || !strings.Contains(err.Error(), "nonzero FILETIME") {
		t.Fatalf("stream FILETIME error = %v", err)
	}

	nonzeroRootCreation := bytes.Clone(data)
	binary.LittleEndian.PutUint64(nonzeroRootCreation[dirOffset+100:], 1)
	if _, err := Open(nonzeroRootCreation); err == nil || !strings.Contains(err.Error(), "root directory entry") {
		t.Fatalf("root creation FILETIME error = %v", err)
	}
}

func loadBin(t *testing.T, book string) []byte {
	t.Helper()
	b, err := os.ReadFile(filepath.Join("testdata", "corpus", book+".bin"))
	if err != nil {
		t.Fatal(err)
	}
	return b
}

func TestPathsReturnsCopy(t *testing.T) {
	// Paths() must not expose the Container's internal order slice; a caller
	// mutating the result must not corrupt subsequent reads.
	c := &Container{order: []string{"PROJECT", "VBA/dir"}}
	got := c.Paths()
	got[0] = "MUTATED"
	if c.Paths()[0] != "PROJECT" {
		t.Errorf("Paths() exposed internal slice: got %q after caller mutation", c.Paths()[0])
	}
}

func TestStoragePathsReturnsCopy(t *testing.T) {
	c := &Container{storageOrder: []string{"", "VBA"}}
	got := c.StoragePaths()
	got[0] = "MUTATED"
	if c.StoragePaths()[0] != "" {
		t.Errorf("StoragePaths() exposed internal slice: got %q after caller mutation", c.StoragePaths()[0])
	}
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
			buf, rerr := io.ReadAll(entry)
			if rerr != nil {
				t.Fatalf("%s: read %q: %v", book, entry.Name, rerr)
			}
			key := entry.Name
			if len(entry.Path) > 0 {
				key = strings.Join(entry.Path, "/") + "/" + entry.Name
			}
			got, ok := lookupStream(c, key)
			if !ok || !bytes.Equal(got, buf) {
				t.Errorf("%s: stream %q does not match mscfb (ok=%v)", book, key, ok)
			}
		}
	}
}
