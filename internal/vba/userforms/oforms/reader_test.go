package oforms

import (
	"bytes"
	"encoding/binary"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
)

func TestReadFormSimpleFixture(t *testing.T) {
	container := openFixture(t, "p4_form.bin")
	forms := DiscoverForms(container)
	if len(forms) != 1 || forms[0] != "UserForm1" {
		t.Fatalf("DiscoverForms = %q, want [UserForm1]", forms)
	}
	form, err := ReadForm(container, forms[0], 932)
	if err != nil {
		t.Fatal(err)
	}
	if form.Name != "UserForm1" || len(form.Levels) != 1 {
		t.Fatalf("form = name %q, %d levels", form.Name, len(form.Levels))
	}
	if form.DesignerSource.Text == "" || form.CompObj.UserType.Text != "Microsoft Forms 2.0 Form" {
		t.Fatalf("designer/CompObj not decoded: source=%q userType=%q", form.DesignerSource.Text, form.CompObj.UserType.Text)
	}
	if len(form.CompObj.Raw) == 0 || len(form.Levels[0].CompObjRaw) == 0 {
		t.Fatal("CompObj raw bytes are empty")
	}
	if &form.CompObj.Raw[0] != &form.Levels[0].CompObjRaw[0] {
		t.Fatal("level and parsed CompObj do not share their retained raw bytes")
	}
	if len(form.Controls) == 0 {
		t.Fatal("simple fixture has no controls")
	}
	button := form.Controls[0]
	if button.Name != "CommandButton1" || button.Kind != "MSForms.CommandButton" || button.Record == nil {
		t.Fatalf("control = name %q kind %q record=%v", button.Name, button.Kind, button.Record != nil)
	}
	if got := button.Record.Strings["Caption"].Text; got != "CommandButton1" {
		t.Fatalf("button caption = %q", got)
	}
	if got := button.Record.Sizes["Size"]; got != (Size{Width: 4445, Height: 2328}) {
		t.Fatalf("button size = %+v", got)
	}
	for _, control := range form.Controls {
		if control.Site == nil || control.Site.Raw == nil {
			t.Fatalf("control %q lost its site bytes", control.Name)
		}
		if control.Record == nil && control.OpaqueRaw == nil {
			t.Fatalf("control %q has neither parsed nor opaque persistence", control.Name)
		}
	}
	assertRawIsolation(t, form)
}

func TestKnownControlRecordTablesAcceptMinimalRecords(t *testing.T) {
	for index, spec := range specsByCacheIndex {
		t.Run(controlKinds[index], func(t *testing.T) {
			maskBytes := 4
			if spec.mask64 {
				maskBytes = 8
			}
			body := []byte{0, spec.major, byte(maskBytes), 0}
			body = append(body, make([]byte, maskBytes)...)
			if spec.textProps {
				body = append(body, 0, 2, 4, 0, 0, 0, 0, 0)
			}
			record, consumed, err := parseRecord(body, spec, 932, len(body))
			if err != nil {
				t.Fatal(err)
			}
			if consumed != len(body) || record.Type != spec.typeName {
				t.Fatalf("record = type %q consumed %d/%d", record.Type, consumed, len(body))
			}
		})
	}
}

func TestReadFormPreservesStructurallyBoundedUnknownControl(t *testing.T) {
	fixture := openFixture(t, "p4_form.bin")
	compObj, _ := fixture.Stream("UserForm1/\x01CompObj")
	vbFrame, _ := fixture.Stream("UserForm1/\x03VBFrame")

	site := make([]byte, 20)
	binary.LittleEndian.PutUint16(site[2:4], 16) // mask + DataBlock
	binary.LittleEndian.PutUint32(site[4:8], (1<<2)|(1<<5)|(1<<7))
	binary.LittleEndian.PutUint32(site[8:12], 1)
	binary.LittleEndian.PutUint32(site[12:16], 4)
	binary.LittleEndian.PutUint16(site[16:18], 99)
	fStream := []byte{0, 4, 4, 0, 0, 0, 0, 0} // empty Form record
	fStream = append(fStream, 0, 0)           // empty class table
	counts := make([]byte, 8)
	binary.LittleEndian.PutUint32(counts[0:4], 1)
	binary.LittleEndian.PutUint32(counts[4:8], 24)
	fStream = append(fStream, counts...)
	fStream = append(fStream, 0, 1, 0, 0) // one depth/type entry + padding
	fStream = append(fStream, site...)

	w := cfb.NewWriter()
	w.AddStorage([]string{"UserForm1"}, cfb.StorageMeta{})
	w.AddStream([]string{"UserForm1", "f"}, fStream)
	w.AddStream([]string{"UserForm1", "o"}, []byte{1, 2, 3, 4})
	w.AddStream([]string{"UserForm1", "\x01CompObj"}, compObj)
	w.AddStream([]string{"UserForm1", "\x03VBFrame"}, vbFrame)
	w.AddStream([]string{"UserForm1", "vendor"}, []byte{9, 8, 7})
	data, err := w.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	container, err := cfb.Open(data)
	if err != nil {
		t.Fatal(err)
	}
	form, err := ReadForm(container, "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	control := form.Controls[0]
	if control.Kind != "MSForms.Control" || !bytes.Equal(control.OpaqueRaw, []byte{1, 2, 3, 4}) {
		t.Fatalf("opaque control = kind %q raw %x", control.Kind, control.OpaqueRaw)
	}
	if !bytes.Equal(form.Levels[0].ExtraStreams["vendor"], []byte{9, 8, 7}) {
		t.Fatalf("extra stream = %x", form.Levels[0].ExtraStreams["vendor"])
	}
}

func TestReadFormNestedFixture(t *testing.T) {
	form, err := ReadForm(openFixture(t, "p6_nested_form.bin"), "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	if len(form.Levels) < 5 {
		t.Fatalf("levels = %d, want nested fixture levels", len(form.Levels))
	}
	containers := 0
	var walk func([]*Control)
	walk = func(controls []*Control) {
		for _, control := range controls {
			if control.Level != nil {
				containers++
				if control.Level.Path == "" || control.Record != control.Level.Record {
					t.Fatalf("container %q is not bound to its level", control.Name)
				}
			}
			walk(control.Children)
		}
	}
	walk(form.Controls)
	if containers < 4 {
		t.Fatalf("container controls = %d, want at least 4", containers)
	}
	for _, level := range form.Levels {
		var total uint64
		for _, site := range level.Sites {
			total += uint64(uint32(site.Values["ObjectStreamSize"]))
		}
		if total != uint64(len(level.ORaw)) {
			t.Fatalf("%s ObjectStreamSize total = %d, o bytes = %d", level.Path, total, len(level.ORaw))
		}
	}
}

func TestReadFormRejectsMissingObjectStream(t *testing.T) {
	original := openFixture(t, "p4_form.bin")
	container := rewriteContainer(t, original, map[string][]byte{"UserForm1/o": nil}, nil)
	var err error
	_, err = ReadForm(container, "UserForm1", 932)
	if !errors.Is(err, ErrMalformed) {
		t.Fatalf("ReadForm error = %v, want ErrMalformed", err)
	}
	var parseErr *ParseError
	if !errors.As(err, &parseErr) || parseErr.Stream != "o" {
		t.Fatalf("ReadForm error = %#v, want o-stream ParseError", err)
	}
}

func TestReadFormRejectsObjectSizeMismatch(t *testing.T) {
	original := openFixture(t, "p4_form.bin")
	oRaw, _ := original.Stream("UserForm1/o")
	container := rewriteContainer(t, original, map[string][]byte{
		"UserForm1/o": append(bytes.Clone(oRaw), 0xff),
	}, nil)
	_, err := ReadForm(container, "UserForm1", 932)
	if !errors.Is(err, ErrMalformed) {
		t.Fatalf("ReadForm error = %v, want ErrMalformed", err)
	}
	var parseErr *ParseError
	if !errors.As(err, &parseErr) || parseErr.Stream != "o" {
		t.Fatalf("ReadForm error = %#v, want o-stream ParseError", err)
	}
}

func TestReadFormRejectsOrphanContainerStorage(t *testing.T) {
	original := openFixture(t, "p4_form.bin")
	container := rewriteContainer(t, original, nil, []string{"UserForm1/i99"})
	_, err := ReadForm(container, "UserForm1", 932)
	if !errors.Is(err, ErrMalformed) {
		t.Fatalf("ReadForm error = %v, want ErrMalformed", err)
	}
}

func TestParseSiteDataRejectsDepthCountMismatch(t *testing.T) {
	data := make([]byte, 12)
	binary.LittleEndian.PutUint32(data[0:4], 2)
	binary.LittleEndian.PutUint32(data[4:8], 4)
	copy(data[8:], []byte{0, 0x83, 1, 0})
	if _, err := parseSiteData(data, 0, 932, false); err == nil {
		t.Fatal("parseSiteData accepted a run accounting for too many sites")
	}
}

func TestReadFormRejectsUnsupportedCodePage(t *testing.T) {
	_, err := ReadForm(openFixture(t, "p4_form.bin"), "UserForm1", 42)
	if !errors.Is(err, ErrMalformed) {
		t.Fatalf("ReadForm error = %v, want ErrMalformed", err)
	}
}

func TestParseCompObjRejectsOversizedStream(t *testing.T) {
	data := make([]byte, maxDesignerStreamSize+1)
	_, err := parseCompObj(data, "UserForm1", 932)
	if !errors.Is(err, ErrMalformed) {
		t.Fatalf("parseCompObj error = %v, want ErrMalformed", err)
	}
	var parseErr *ParseError
	if !errors.As(err, &parseErr) || parseErr.Stream != "\x01CompObj" {
		t.Fatalf("parseCompObj error = %#v, want CompObj ParseError", err)
	}
}

func openFixture(t testing.TB, name string) *cfb.Container {
	t.Helper()
	body, err := os.ReadFile(filepath.Join("..", "..", "..", "pack", "vbaproject", "testdata", "corpus", name))
	if err != nil {
		t.Fatal(err)
	}
	container, err := cfb.Open(body)
	if err != nil {
		t.Fatal(err)
	}
	return container
}

func splitPath(path string) []string {
	if path == "" {
		return nil
	}
	var parts []string
	for path != "" {
		before, after, found := strings.Cut(path, "/")
		parts = append(parts, before)
		if !found {
			break
		}
		path = after
	}
	return parts
}

func rewriteContainer(t *testing.T, original *cfb.Container, replacements map[string][]byte, extraStorages []string) *cfb.Container {
	t.Helper()
	w, err := cfb.NewWriterForFormat(original.Format())
	if err != nil {
		t.Fatal(err)
	}
	for _, path := range original.StoragePaths() {
		meta, _ := original.Storage(path)
		w.AddStorage(splitPath(path), meta)
	}
	for _, path := range extraStorages {
		w.AddStorage(splitPath(path), cfb.StorageMeta{})
	}
	for _, path := range original.Paths() {
		if replacement, found := replacements[path]; found {
			if replacement != nil {
				w.AddStream(splitPath(path), replacement)
			}
			continue
		}
		body, _ := original.Stream(path)
		w.AddStream(splitPath(path), body)
	}
	data, err := w.Bytes()
	if err != nil {
		t.Fatal(err)
	}
	container, err := cfb.Open(data)
	if err != nil {
		t.Fatal(err)
	}
	return container
}

func assertRawIsolation(t *testing.T, form *Form) {
	t.Helper()
	before := form.Levels[0].FRaw[0]
	form.DesignerSource.Raw[0] ^= 0xff
	if form.Levels[0].FRaw[0] != before {
		t.Fatal("independent raw streams unexpectedly alias")
	}
}
