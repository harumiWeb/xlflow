package ovba

import (
	"encoding/binary"
	"os"
	"path/filepath"
	"testing"
)

func TestWalkRecordsReachesModules(t *testing.T) {
	plain, err := os.ReadFile(filepath.Join("testdata", "p1_compiled", "dir.plain"))
	if err != nil {
		t.Fatal(err)
	}
	var sawSysKind, sawModuleOffset bool
	for _, r := range walkRecords(plain) {
		switch r.id {
		case 0x0001:
			sawSysKind = true
		case 0x0031:
			sawModuleOffset = true // unreachable unless 0x0009 is skipped correctly
		}
	}
	if !sawSysKind || !sawModuleOffset {
		t.Errorf("the walk did not reach the modules (sysKind=%v moduleOffset=%v)", sawSysKind, sawModuleOffset)
	}
}

func TestParseDirP1(t *testing.T) {
	plain, err := os.ReadFile(filepath.Join("testdata", "p1_compiled", "dir.plain"))
	if err != nil {
		t.Fatal(err)
	}
	di := ParseDir(plain)
	if di.CodePage != 932 {
		t.Errorf("CodePage = %d, want 932", di.CodePage)
	}
	if di.SysKind != 3 {
		t.Errorf("SysKind = %d, want 3", di.SysKind)
	}
	want := map[string]uint32{"ThisWorkbook": 1380, "Sheet1": 1545, "Module1": 1130, "Class1": 2644}
	got := map[string]uint32{}
	for _, m := range di.Modules {
		got[m.Name] = m.Offset
	}
	for k, v := range want {
		if got[k] != v {
			t.Errorf("module %s offset = %d, want %d", k, got[k], v)
		}
	}
	if len(di.RefNames) != 2 { // stdole/Office
		t.Errorf("RefNames = %v, want [stdole Office]", di.RefNames)
	}
}

func TestParseDirReferences(t *testing.T) {
	read := func(book string) DirInfo {
		plain, err := os.ReadFile(filepath.Join("testdata", book, "dir.plain"))
		if err != nil {
			t.Fatal(err)
		}
		return ParseDir(plain)
	}
	// p4_form: duplicate names from nested REFERENCECONTROL must not inflate the count (I1 regression)
	if got := read("p4_form").RefNames; !eqStrings(got, []string{"stdole", "Office", "MSForms"}) {
		t.Errorf("p4_form RefNames = %v, want [stdole Office MSForms]", got)
	}
	// p2_refs: the additional Scripting reference is picked up
	if got := read("p2_refs").RefNames; !eqStrings(got, []string{"stdole", "Office", "Scripting"}) {
		t.Errorf("p2_refs RefNames = %v, want [stdole Office Scripting]", got)
	}
	// RefsRaw is the verbatim span that begins with the reference record (0x0016)
	if raw := read("p4_form").RefsRaw; len(raw) == 0 || raw[0] != 0x16 {
		t.Errorf("RefsRaw does not begin with 0x0016 (len=%d)", len(raw))
	}
}

func eqStrings(a, b []string) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
}

func TestParseProjectText(t *testing.T) {
	sample := "ID=\"{ABC}\"\r\n" +
		"Document=ThisWorkbook/&H00000000\r\n" +
		"Document=Sheet1/&H00000000\r\n" +
		"Module=Module1\r\n" +
		"Class=Class1\r\n" +
		"BaseClass=UserForm1\r\n" +
		"Name=\"VBAProject\"\r\n" +
		"CMG=\"0C0E\"\r\nDPB=\"5D5F\"\r\nGC=\"AEAC\"\r\n"
	pt := ParseProjectText([]byte(sample))
	if pt.ID != "{ABC}" || pt.Name != "VBAProject" {
		t.Errorf("ID/Name = %q/%q", pt.ID, pt.Name)
	}
	if pt.CMG != "0C0E" || pt.DPB != "5D5F" || pt.GC != "AEAC" {
		t.Errorf("CMG/DPB/GC = %q/%q/%q", pt.CMG, pt.DPB, pt.GC)
	}
	want := map[string]string{"ThisWorkbook": "Document", "Sheet1": "Document",
		"Module1": "Module", "Class1": "Class", "UserForm1": "BaseClass"}
	for name, kind := range want {
		if pt.Kinds[name] != kind {
			t.Errorf("%s kind = %q, want %q", name, pt.Kinds[name], kind)
		}
	}
}

func TestParseDirProjectInfoRaw(t *testing.T) {
	plain, err := os.ReadFile(filepath.Join("testdata", "golden", "dir.plain"))
	if err != nil {
		t.Fatal(err)
	}
	di := ParseDir(plain)
	if len(di.ProjectInfoRaw) == 0 {
		t.Fatal("ProjectInfoRaw is empty")
	}
	// PROJECTINFORMATION starts at the beginning of dir.plain.
	if di.ProjectInfoRaw[0] != plain[0] {
		t.Error("ProjectInfoRaw does not start at the beginning")
	}
	// The first record is SYSKIND (id=0x0001).
	if id := binary.LittleEndian.Uint16(di.ProjectInfoRaw); id != 0x0001 {
		t.Errorf("first record id = 0x%04X, want 0x0001 (SYSKIND)", id)
	}
	// Invariant: ProjectInfoRaw ++ RefsRaw covers exactly up to just before PROJECTMODULES (0x000F).
	// The byte immediately after is the MODULES count record id=0x000F.
	joined := len(di.ProjectInfoRaw) + len(di.RefsRaw)
	if id := binary.LittleEndian.Uint16(plain[joined:]); id != 0x000F {
		t.Errorf("id immediately after ProjectInfoRaw+RefsRaw = 0x%04X, want 0x000F (PROJECTMODULES)", id)
	}
}
