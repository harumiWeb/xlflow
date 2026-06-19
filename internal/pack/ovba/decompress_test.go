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
		comp, err := Compress(in)
		if err != nil {
			t.Fatalf("Compress error: %v", err)
		}
		got, err := Decompress(comp)
		if err != nil {
			t.Fatalf("Decompress error: %v", err)
		}
		if !bytes.Equal(got, in) {
			t.Errorf("round-trip mismatch\n in=%q\nout=%q", in, got)
		}
	}
}

func TestDecompressChunkRejectsOversizedOutput(t *testing.T) {
	// flag 0x02: bit0 literal, bit1 copy-token. One seed byte then a copy token
	// (offset 1, max length) expands the output to ~4099 bytes, past the 4096
	// per-chunk ceiling ([MS-OVBA] §2.4.1.1.4). A malformed chunk must be rejected
	// rather than allocating unbounded output.
	body := []byte{0x02, 0x41, 0xFF, 0x0F}
	if _, err := decompressChunk(body); err == nil {
		t.Fatal("expected error for chunk exceeding 4096 decompressed bytes, got nil")
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
