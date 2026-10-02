package ovba

import (
	"bytes"
	"encoding/binary"
	"os"
	"testing"
)

func TestReferenceGroupsAndTruncations(t *testing.T) {
	body, err := os.ReadFile("testdata/p4_form/dir.plain")
	if err != nil {
		t.Fatal(err)
	}
	info, err := ParseDir(body)
	if err != nil {
		t.Fatal(err)
	}
	refs, err := ParseProjectReferences(info.RefsRaw, info.CodePage)
	if err != nil {
		t.Fatal(err)
	}
	if len(refs) != 3 || refs[2].Kind != 0x002F {
		t.Fatalf("groups = %+v", refs)
	}
	var joined []byte
	for _, ref := range refs {
		joined = append(joined, ref.Raw...)
	}
	if !bytes.Equal(joined, info.RefsRaw) {
		t.Fatal("lossless reference span mismatch")
	}
	last := refs[2].Raw
	for n := 1; n < len(last); n++ {
		if _, err := ParseProjectReferences(last[:n], info.CodePage); err == nil {
			t.Fatalf("accepted truncated CONTROL at %d", n)
		}
	}
	// Aggregate CONTROL lengths are explicitly ignored on read.
	control := bytes.Clone(last)
	for i := 0; i+6 <= len(control); {
		id := binary.LittleEndian.Uint16(control[i:])
		n := int(binary.LittleEndian.Uint32(control[i+2:]))
		if id == 0x002F || id == 0x0030 {
			binary.LittleEndian.PutUint32(control[i+2:], 0)
		}
		i += 6 + n
	}
	if _, err := ParseProjectReferences(control, info.CodePage); err != nil {
		t.Fatal(err)
	}
}

func TestReferenceWithoutName(t *testing.T) {
	raw, err := BuildProjectReferences([]RegisteredReferenceSpec{{Name: "Test", LibID: `*\G{00020430-0000-0000-C000-000000000046}#2.0#0#stdole2.tlb#OLE Automation`}}, 1252)
	if err != nil {
		t.Fatal(err)
	}
	for range 2 {
		n := int(binary.LittleEndian.Uint32(raw[2:]))
		raw = raw[6+n:]
	}
	refs, err := ParseProjectReferences(raw, 1252)
	if err != nil || len(refs) != 1 || refs[0].Name != "" {
		t.Fatalf("unnamed ref: %v %v", refs, err)
	}
	info, err := BuildProjectInformation(ProjectInformationSpec{SysKind: 1, CodePage: 1252, Name: "Test"})
	if err != nil {
		t.Fatal(err)
	}
	modules, err := BuildProjectModules(nil, 1252)
	if err != nil {
		t.Fatal(err)
	}
	plain := append(append(bytes.Clone(info), raw...), modules...)
	parsed, err := ParseDir(plain)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(parsed.RefsRaw, raw) || !bytes.Equal(parsed.ProjectInfoRaw, info) {
		t.Fatal("unnamed reference was absorbed into project information")
	}
}

func FuzzProjectReferences(f *testing.F) {
	raw, _ := BuildProjectReferences([]RegisteredReferenceSpec{{Name: "Test", LibID: `*\G{00020430-0000-0000-C000-000000000046}#2.0#0##`}}, 1252)
	f.Add(raw)
	f.Fuzz(func(t *testing.T, raw []byte) {
		refs, err := ParseProjectReferences(raw, 1252)
		if err != nil {
			return
		}
		var joined []byte
		for _, ref := range refs {
			joined = append(joined, ref.Raw...)
		}
		if !bytes.Equal(joined, raw) {
			t.Fatal("reference bytes changed")
		}
	})
}
