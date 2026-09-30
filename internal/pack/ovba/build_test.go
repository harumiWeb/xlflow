package ovba

import (
	"bytes"
	"testing"
)

func TestBuildProjectWMUsesExcelTerminator(t *testing.T) {
	got, err := BuildProjectWM([]ProjectComponentSpec{{Kind: "Module", Name: "Main"}}, 1252)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.HasSuffix(got, []byte{'n', 0, 0, 0, 0, 0}) {
		t.Fatalf("PROJECTwm suffix = %x", got)
	}
}

func TestBuildProjectInformationUsesCanonicalLCID(t *testing.T) {
	got, err := BuildProjectInformation(ProjectInformationSpec{SysKind: 1, CodePage: 932, Name: "VBAProject"})
	if err != nil {
		t.Fatal(err)
	}
	records, err := walkRecords(got)
	if err != nil {
		t.Fatal(err)
	}
	found := map[uint16]bool{}
	for _, record := range records {
		if record.id != 0x0002 && record.id != 0x0014 {
			continue
		}
		if len(record.payload) != 4 || le32(record.payload) != CanonicalProjectLCID {
			t.Fatalf("record %#04x payload = %x", record.id, record.payload)
		}
		found[record.id] = true
	}
	if !found[0x0002] || !found[0x0014] {
		t.Fatalf("canonical LCID records found = %v", found)
	}
}

func TestBuildProjectTextIsDeterministicAndComplete(t *testing.T) {
	specs := []ProjectComponentSpec{{Kind: "Document", Name: "ThisWorkbook"}, {Kind: "Module", Name: "Main"}}
	first, err := BuildProjectText("VBAProject", specs, 1252)
	if err != nil {
		t.Fatal(err)
	}
	second, err := BuildProjectText("VBAProject", specs, 1252)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(first, second) {
		t.Fatal("PROJECT stream is not deterministic")
	}
	for _, want := range [][]byte{[]byte(`ID="{61CB3C72-4521-44A9-8538-FAE59C5B1642}"`), []byte("Document=ThisWorkbook"), []byte("Module=Main"), []byte("ThisWorkbook=0, 0, 0, 0, C"), []byte("Main=32, 32, 1872, 983, Z")} {
		if !bytes.Contains(first, want) {
			t.Fatalf("PROJECT stream missing %q", want)
		}
	}
}
