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
	for _, want := range [][]byte{[]byte("Document=ThisWorkbook"), []byte("Module=Main"), []byte("ThisWorkbook=0, 0, 0, 0, C"), []byte("Main=32, 32, 1872, 983, Z")} {
		if !bytes.Contains(first, want) {
			t.Fatalf("PROJECT stream missing %q", want)
		}
	}
}
