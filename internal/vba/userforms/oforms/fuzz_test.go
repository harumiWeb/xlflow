package oforms

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
)

func FuzzReadForm(f *testing.F) {
	for _, name := range []string{"p4_form.bin", "p6_nested_form.bin"} {
		body, err := os.ReadFile(filepath.Join("..", "..", "..", "pack", "vbaproject", "testdata", "corpus", name))
		if err != nil {
			f.Fatal(err)
		}
		f.Add(body)
	}
	f.Fuzz(func(t *testing.T, data []byte) {
		container, err := cfb.Open(data)
		if err != nil {
			return
		}
		for _, name := range DiscoverForms(container) {
			_, _ = ReadForm(container, name, 932)
		}
	})
}

func FuzzParseFormStream(f *testing.F) {
	for _, fixture := range []string{"p4_form.bin", "p6_nested_form.bin"} {
		body, err := os.ReadFile(filepath.Join("..", "..", "..", "pack", "vbaproject", "testdata", "corpus", fixture))
		if err != nil {
			f.Fatal(err)
		}
		container, err := cfb.Open(body)
		if err != nil {
			f.Fatal(err)
		}
		for _, path := range container.Paths() {
			if strings.HasPrefix(path, "UserForm1/") && strings.HasSuffix(path, "/f") {
				stream, _ := container.Stream(path)
				f.Add(stream)
			}
		}
	}
	f.Fuzz(func(t *testing.T, data []byte) {
		_, _ = parseFormStream(data, "fuzz", 932)
	})
}
