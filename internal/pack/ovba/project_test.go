package ovba

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"
)

func TestVBAProjectStub(t *testing.T) {
	want, _ := os.ReadFile(filepath.Join("testdata", "golden", "_VBA_PROJECT"))
	if got := VBAProjectStub(); !bytesEqual(got, want) {
		t.Errorf("VBAProjectStub = %x, want %x", got, want)
	}
}

func TestBuildProjectModulesMatchesGolden(t *testing.T) {
	plain, err := os.ReadFile(filepath.Join("testdata", "golden", "dir.plain"))
	if err != nil {
		t.Fatal(err)
	}
	di := ParseDir(plain)
	// The PROJECTMODULES portion of the golden = from just after ProjectInfoRaw + RefsRaw to the end.
	got := BuildProjectModules([]ModuleSpec{
		{Name: "Sheet1", StreamName: "Sheet1", TypeID: 0x0022},
		{Name: "ThisWorkbook", StreamName: "ThisWorkbook", TypeID: 0x0022},
		{Name: "Spike", StreamName: "Spike", TypeID: 0x0021},
	})
	want := plain[len(di.ProjectInfoRaw)+len(di.RefsRaw):]
	if !bytes.Equal(got, want) {
		t.Errorf("PROJECTMODULES mismatch\n got  %d bytes\n want %d bytes", len(got), len(want))
	}
}

func TestModRecordDistinctStreamName(t *testing.T) {
	// When Name != StreamName, both must appear verbatim.
	r := modRecord("Long Module Name", "LongMod", 0x0021)
	if !bytes.Contains(r, []byte("Long Module Name")) {
		t.Error("MODULENAME is missing")
	}
	if !bytes.Contains(r, []byte("LongMod")) {
		t.Error("MODULESTREAMNAME is missing")
	}
}
