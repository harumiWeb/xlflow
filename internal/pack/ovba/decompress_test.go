package ovba

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"
)

func TestDecompressRejectsTruncatedRawChunk(t *testing.T) {
	// Crafted input with the raw-chunk flag (bit15=0) but fewer than 4096B -> error, not a panic.
	if _, err := Decompress([]byte{0x01, 0x00, 0x00, 0x42}); err == nil {
		t.Error("a truncated raw chunk should return an error (not panic)")
	}
}

func TestRoundTripSmall(t *testing.T) {
	for _, in := range [][]byte{
		[]byte("Attribute VB_Name = \"Module1\"\r\nSub X()\r\nEnd Sub\r\n"),
		bytes.Repeat([]byte("a"), 50), // RLE
		{},                            // empty
	} {
		got, err := Decompress(Compress(in))
		if err != nil {
			t.Fatalf("Decompress error: %v", err)
		}
		if !bytes.Equal(got, in) {
			t.Errorf("round-trip mismatch\n in=%q\nout=%q", in, got)
		}
	}
}

func TestDecompressGoldenModules(t *testing.T) {
	books := []string{"p1_compiled", "p5_mbcs"}
	mods := []string{"Module1", "Class1", "Sheet1", "ThisWorkbook"}
	for _, book := range books {
		for _, mod := range mods {
			dir := filepath.Join("testdata", book, "modules")
			comp, err := os.ReadFile(filepath.Join(dir, mod+".comp"))
			if err != nil {
				continue // skip e.g. when p5 has no module of the same name
			}
			want, _ := os.ReadFile(filepath.Join(dir, mod+".plain"))
			got, err := Decompress(comp)
			if err != nil {
				t.Fatalf("%s/%s: %v", book, mod, err)
			}
			if !bytes.Equal(got, want) {
				t.Errorf("%s/%s: Decompress != golden plain (got %d, want %d)",
					book, mod, len(got), len(want))
			}
		}
	}
}
