package vbaproject

import (
	"os"
	"path/filepath"
	"testing"
)

func FuzzRead(f *testing.F) {
	for _, name := range []string{"p1_compiled.bin", "p2_refs.bin", "p4_form.bin", "p5_mbcs.bin"} {
		data, err := os.ReadFile(filepath.Join("testdata", "corpus", name))
		if err != nil {
			f.Fatal(err)
		}
		f.Add(data)
	}
	f.Fuzz(func(t *testing.T, data []byte) {
		project, err := Read(data)
		if err != nil || project.Protection.IsProtected {
			return
		}
		rebuilt, err := Write(project)
		if err != nil {
			return
		}
		if _, err := Read(rebuilt); err != nil {
			t.Fatalf("read rebuilt project: %v", err)
		}
	})
}
