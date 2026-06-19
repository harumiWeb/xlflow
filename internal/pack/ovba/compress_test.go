package ovba

import (
	"os"
	"path/filepath"
	"testing"
)

func TestCopyTokenHelp(t *testing.T) {
	// bitCount and maxLength for each pos (byte position from the start of the chunk).
	// Per the [MS-OVBA] §2.4.1.3.19.3 definition: bitCount = max(ceil(log2(pos)), 4),
	// maxLength = (0xFFFF >> bitCount) + 3.
	tests := []struct {
		pos          int
		wantBitCount int
		wantMaxLen   int
	}{
		{pos: 1, wantBitCount: 4, wantMaxLen: 4098},
		{pos: 16, wantBitCount: 4, wantMaxLen: 4098},
		{pos: 17, wantBitCount: 5, wantMaxLen: 2050},
		{pos: 4096, wantBitCount: 12, wantMaxLen: 18},
	}
	for _, tt := range tests {
		bc, maxLen := copyTokenHelp(tt.pos)
		if bc != tt.wantBitCount || maxLen != tt.wantMaxLen {
			t.Errorf("copyTokenHelp(%d) = (bitCount %d, maxLen %d), want (%d, %d)",
				tt.pos, bc, maxLen, tt.wantBitCount, tt.wantMaxLen)
		}
	}
}

func TestLongestMatch(t *testing.T) {
	// In "abcabc", from pos=3 the match "abc" (offset 3, length 3) starts at pos=0.
	chunk := []byte("abcabc")
	off, length := longestMatch(chunk, 3)
	if off != 3 || length != 3 {
		t.Errorf("longestMatch = (off %d, len %d), want (3, 3)", off, length)
	}

	// No match (at the start) yields length 0.
	if _, l := longestMatch(chunk, 0); l != 0 {
		t.Errorf("longestMatch at 0 len = %d, want 0", l)
	}

	// Overlapping copy (RLE): in "aaaa", from pos=1 offset 1, length 3 (reads what was already emitted).
	run := []byte("aaaa")
	o, l := longestMatch(run, 1)
	if o != 1 || l != 3 {
		t.Errorf("longestMatch RLE = (off %d, len %d), want (1, 3)", o, l)
	}
}

func TestCompressMatchesGolden(t *testing.T) {
	// Mapping of input (normalized source) -> golden (the compressed stream from an accepted bin).
	cases := []struct{ in, golden string }{
		{"src/Spike.norm", "golden/Spike"},
		{"src/Sheet1.norm", "golden/Sheet1"},
		{"src/ThisWorkbook.norm", "golden/ThisWorkbook"},
	}
	for _, c := range cases {
		in, err := os.ReadFile(filepath.Join("testdata", c.in))
		if err != nil {
			t.Fatal(err)
		}
		want, err := os.ReadFile(filepath.Join("testdata", c.golden))
		if err != nil {
			t.Fatal(err)
		}
		got := Compress(in)
		if !bytesEqual(got, want) {
			t.Errorf("%s: Compress len %d != golden len %d (first diff at %d)",
				c.in, len(got), len(want), firstDiff(got, want))
		}
	}
}

func bytesEqual(a, b []byte) bool {
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

func firstDiff(a, b []byte) int {
	n := len(a)
	if len(b) < n {
		n = len(b)
	}
	for i := 0; i < n; i++ {
		if a[i] != b[i] {
			return i
		}
	}
	return n
}
