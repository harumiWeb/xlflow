package ovba

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"
)

func FuzzDecompress(f *testing.F) {
	for _, path := range []string{
		filepath.Join("testdata", "p1_compiled", "modules", "Module1.comp"),
		filepath.Join("testdata", "p5_mbcs", "modules", "ThisWorkbook.comp"),
	} {
		data, err := os.ReadFile(path)
		if err != nil {
			f.Fatal(err)
		}
		f.Add(data)
	}
	f.Add([]byte{0x01})

	f.Fuzz(func(t *testing.T, data []byte) {
		plain, err := Decompress(data)
		if err != nil {
			return
		}
		compressed, err := Compress(plain)
		if err != nil {
			return
		}
		roundTrip, err := Decompress(compressed)
		if err != nil {
			t.Fatalf("decompress recompressed data: %v", err)
		}
		if !bytes.Equal(roundTrip, plain) {
			t.Fatal("decompress-compress-decompress changed the payload")
		}
	})
}

func FuzzParseDir(f *testing.F) {
	for _, name := range []string{"p1_compiled", "p4_form", "p5_mbcs"} {
		data, err := os.ReadFile(filepath.Join("testdata", name, "dir.plain"))
		if err != nil {
			f.Fatal(err)
		}
		f.Add(data)
	}
	f.Fuzz(func(t *testing.T, data []byte) {
		info, err := ParseDir(data)
		if err == nil && info.CodePage == 0 {
			t.Fatal("ParseDir succeeded without a code page")
		}
	})
}

func FuzzDecryptData(f *testing.F) {
	f.Add(encryptData(0x00, 0x7B, []byte{1, 0, 0, 0}))
	f.Add([]byte{0x00, 0x02})
	f.Fuzz(func(_ *testing.T, data []byte) {
		_, _ = DecryptData(data)
	})
}
