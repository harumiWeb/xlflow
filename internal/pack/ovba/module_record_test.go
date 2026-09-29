package ovba

import (
	"encoding/binary"
	"fmt"
	"strings"
	"testing"
)

// dirPlainFixture builds a minimal dir.plain: project information records
// (SYSKIND/LCID/CODEPAGE) followed by a PROJECTMODULES section containing the
// supplied module record payloads.
func dirPlainFixture(codepage uint16, moduleRecords ...[]byte) []byte {
	p4 := func(v uint32) []byte {
		b := make([]byte, 4)
		binary.LittleEndian.PutUint32(b, v)
		return b
	}
	p2 := func(v uint16) []byte {
		b := make([]byte, 2)
		binary.LittleEndian.PutUint16(b, v)
		return b
	}
	var plain []byte
	plain = append(plain, sizedRecord(0x0001, p4(3))...)
	plain = append(plain, sizedRecord(0x0014, p4(1033))...)
	plain = append(plain, sizedRecord(0x0003, p2(codepage))...)
	plain = append(plain, sizedRecord(0x000F, p2(uint16(len(moduleRecords))))...)
	plain = append(plain, sizedRecord(0x0013, p2(0xFFFF))...)
	for _, m := range moduleRecords {
		plain = append(plain, m...)
	}
	plain = append(plain, sizedRecord(0x0010, nil)...)
	return plain
}

func TestParseDirNonASCIINamesAndMetadata(t *testing.T) {
	spec := ModuleSpec{
		Name:        "\u6a19\u6e96\u30e2\u30b8\u30e5\u30fc\u30eb",
		StreamName:  "\u6a19\u6e96\u30e2\u30b8\u30e5\u30fc\u30eb",
		TypeID:      0x0021,
		DocString:   "\u8aac\u660e",
		HelpContext: 7,
		ReadOnly:    true,
	}
	rec, err := modRecord(spec, 932)
	if err != nil {
		t.Fatalf("modRecord: %v", err)
	}
	di, err := ParseDir(dirPlainFixture(932, rec))
	if err != nil {
		t.Fatalf("ParseDir: %v", err)
	}
	if len(di.Modules) != 1 {
		t.Fatalf("modules = %d, want 1", len(di.Modules))
	}
	m := di.Modules[0]
	if m.Name != spec.Name || m.StreamName != spec.StreamName {
		t.Errorf("names drifted: Name=%q StreamName=%q", m.Name, m.StreamName)
	}
	if m.DocString != spec.DocString || m.HelpContext != 7 || !m.ReadOnly {
		t.Errorf("metadata drifted: doc=%q help=%d readOnly=%v", m.DocString, m.HelpContext, m.ReadOnly)
	}
}

func TestParseDirPreservesUnknownModuleRecords(t *testing.T) {
	spec := ModuleSpec{Name: "Module1", StreamName: "Module1", TypeID: 0x0021}
	rec, err := modRecord(spec, 1252)
	if err != nil {
		t.Fatal(err)
	}
	// Inject an unknown record (0x7777) inside the module, before its terminator.
	extra := sizedRecord(0x7777, []byte{0xDE, 0xAD})
	rec = append(rec[:len(rec)-6], append(extra, sizedRecord(0x002B, nil)...)...)
	di, err := ParseDir(dirPlainFixture(1252, rec))
	if err != nil {
		t.Fatalf("ParseDir: %v", err)
	}
	if len(di.Modules) != 1 || len(di.Modules[0].Extra) != 1 {
		t.Fatalf("extras = %+v", di.Modules)
	}
	if di.Modules[0].Extra[0].ID != 0x7777 {
		t.Errorf("extra id = 0x%04X", di.Modules[0].Extra[0].ID)
	}
}

func TestParseDirRejectsMismatchedUnicodeName(t *testing.T) {
	nameMBCS, err := EncodeMBCS("\u30e2\u30b8\u30e5", 932)
	if err != nil {
		t.Fatal(err)
	}
	var rec []byte
	rec = append(rec, sizedRecord(0x0019, nameMBCS)...)
	rec = append(rec, sizedRecord(0x0047, utf16le("different"))...) // inconsistent unicode
	rec = append(rec, sizedRecord(0x001A, nameMBCS)...)
	rec = append(rec, sizedRecord(0x0032, utf16le("\u30e2\u30b8\u30e5"))...)
	rec = append(rec, sizedRecord(0x001C, nil)...)
	rec = append(rec, sizedRecord(0x0048, nil)...)
	rec = append(rec, recU32(0x0031, 4, 0)...)
	rec = append(rec, recU32(0x001E, 4, 0)...)
	rec = append(rec, recU16(0x002C, 2, 0xFFFF)...)
	rec = append(rec, sizedRecord(0x0021, nil)...)
	rec = append(rec, sizedRecord(0x002B, nil)...)
	if _, err := ParseDir(dirPlainFixture(932, rec)); err == nil || !strings.Contains(err.Error(), "mismatch") {
		t.Fatalf("err = %v, want MBCS/Unicode mismatch", err)
	}
}

func TestParseDirRejectsMisorderedOrUnterminated(t *testing.T) {
	spec := ModuleSpec{Name: "M", StreamName: "M", TypeID: 0x0021}
	rec, err := modRecord(spec, 1252)
	if err != nil {
		t.Fatal(err)
	}
	// Unterminated: drop the module terminator entirely.
	if _, err := ParseDir(dirPlainFixture(1252, rec[:len(rec)-6])); err == nil {
		t.Error("module without terminator should fail")
	}
	// Misordered: put MODULEOFFSET (0x0031) before MODULESTREAMNAME.
	nameMBCS := []byte("M")
	var bad []byte
	bad = append(bad, sizedRecord(0x0019, nameMBCS)...)
	bad = append(bad, sizedRecord(0x0047, utf16le("M"))...)
	bad = append(bad, recU32(0x0031, 4, 0)...)
	bad = append(bad, sizedRecord(0x001A, nameMBCS)...)
	bad = append(bad, sizedRecord(0x0032, utf16le("M"))...)
	bad = append(bad, sizedRecord(0x001C, nil)...)
	bad = append(bad, sizedRecord(0x0048, nil)...)
	bad = append(bad, recU32(0x001E, 4, 0)...)
	bad = append(bad, recU16(0x002C, 2, 0xFFFF)...)
	bad = append(bad, sizedRecord(0x0021, nil)...)
	bad = append(bad, sizedRecord(0x002B, nil)...)
	if _, err := ParseDir(dirPlainFixture(1252, bad)); err == nil || !strings.Contains(err.Error(), "out of order") {
		t.Fatalf("err = %v, want out-of-order record error", err)
	}
}

func TestBuildProjectModulesRejectsUnrepresentableName(t *testing.T) {
	_, err := BuildProjectModules([]ModuleSpec{
		{Name: "mod\U0001F600", StreamName: "mod\U0001F600", TypeID: 0x0021},
	}, 932)
	if err == nil {
		t.Fatal("unrepresentable name should fail")
	}
}

func TestParseDirSupplementaryPlaneRoundTrip(t *testing.T) {
	// Code page 65001 can encode supplementary-plane characters, which take a
	// surrogate pair in the paired UTF-16 records.
	spec := ModuleSpec{
		Name:       "mod\U00010400ule",
		StreamName: "mod\U00010400ule",
		TypeID:     0x0021,
		DocString:  "doc\U00010401",
	}
	rec, err := modRecord(spec, 65001)
	if err != nil {
		t.Fatalf("modRecord: %v", err)
	}
	di, err := ParseDir(dirPlainFixture(65001, rec))
	if err != nil {
		t.Fatalf("ParseDir: %v", err)
	}
	if len(di.Modules) != 1 {
		t.Fatalf("modules = %d, want 1", len(di.Modules))
	}
	m := di.Modules[0]
	if m.Name != spec.Name || m.StreamName != spec.StreamName || m.DocString != spec.DocString {
		t.Errorf("supplementary-plane fields drifted: Name=%q StreamName=%q DocString=%q", m.Name, m.StreamName, m.DocString)
	}
}

func TestParseDirRejectsUnpairedSurrogates(t *testing.T) {
	build := func(nameUnits []uint16) []byte {
		nameMBCS := []byte("M")
		var uni []byte
		for _, u := range nameUnits {
			var p [2]byte
			binary.LittleEndian.PutUint16(p[:], u)
			uni = append(uni, p[:]...)
		}
		var rec []byte
		rec = append(rec, sizedRecord(0x0019, nameMBCS)...)
		rec = append(rec, sizedRecord(0x0047, uni)...)
		rec = append(rec, sizedRecord(0x001A, nameMBCS)...)
		rec = append(rec, sizedRecord(0x0032, utf16le("M"))...)
		rec = append(rec, sizedRecord(0x001C, nil)...)
		rec = append(rec, sizedRecord(0x0048, nil)...)
		rec = append(rec, recU32(0x0031, 4, 0)...)
		rec = append(rec, recU32(0x001E, 4, 0)...)
		rec = append(rec, recU16(0x002C, 2, 0xFFFF)...)
		rec = append(rec, sizedRecord(0x0021, nil)...)
		rec = append(rec, sizedRecord(0x002B, nil)...)
		return rec
	}
	for _, tc := range []struct {
		name  string
		units []uint16
	}{
		{"unpaired high surrogate at end", []uint16{0x004D, 0xD800}},
		{"high surrogate followed by BMP", []uint16{0xD800, 0x004D}},
		{"lone low surrogate", []uint16{0x004D, 0xDC00}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			if _, err := ParseDir(dirPlainFixture(1252, build(tc.units))); err == nil {
				t.Fatal("unpaired surrogate should fail")
			}
		})
	}
}

func TestBuildProjectModulesRejectsReservedExtraIDs(t *testing.T) {
	for _, id := range []uint16{0x000F, 0x0010, 0x0013, 0x0019, 0x0047, 0x001A,
		0x0032, 0x001C, 0x0048, 0x0031, 0x001E, 0x002C, 0x0021, 0x0022, 0x0025,
		0x0028, 0x002B} {
		t.Run(fmt.Sprintf("0x%04X", id), func(t *testing.T) {
			_, err := BuildProjectModules([]ModuleSpec{
				{Name: "M", StreamName: "M", TypeID: 0x0021,
					Extra: []ModuleExtraRecord{{ID: id}}},
			}, 1252)
			if err == nil {
				t.Fatal("reserved record id supplied as an extra should fail")
			}
		})
	}
	// An unknown id must still pass through.
	rec, err := BuildProjectModules([]ModuleSpec{
		{Name: "M", StreamName: "M", TypeID: 0x0021,
			Extra: []ModuleExtraRecord{{ID: 0x7777, Payload: []byte{0x01}}}},
	}, 1252)
	if err != nil {
		t.Fatalf("unknown extra record: %v", err)
	}
	if len(rec) == 0 {
		t.Fatal("empty output")
	}
}
