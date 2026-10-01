package cfb

import (
	"bytes"
	"os"
	"path/filepath"
	"slices"
	"strings"
	"testing"
)

func FuzzOpen(f *testing.F) {
	for _, name := range []string{"p1_compiled.bin", "p4_form.bin", "p5_mbcs.bin"} {
		data, err := os.ReadFile(filepath.Join("testdata", "corpus", name))
		if err != nil {
			f.Fatal(err)
		}
		f.Add(data)
	}
	for _, format := range []Format{FormatV3, FormatV4} {
		w, err := NewWriterForFormat(format)
		if err != nil {
			f.Fatal(err)
		}
		w.AddStream([]string{"PROJECT"}, []byte("seed"))
		w.AddStream([]string{"VBA", "Module1"}, bytes.Repeat([]byte("x"), 5000))
		w.AddStorage(nil, StorageMeta{Modified: 1})
		w.AddStorage([]string{"Form", "empty"}, StorageMeta{CLSID: [16]byte{1, 2, 3}, Modified: 2})
		data, err := w.Bytes()
		if err != nil {
			f.Fatal(err)
		}
		f.Add(data)
	}

	f.Fuzz(func(t *testing.T, data []byte) {
		container, err := Open(data)
		if err != nil {
			return
		}
		w, err := NewWriterForFormat(container.Format())
		if err != nil {
			t.Fatal(err)
		}
		for _, path := range container.StoragePaths() {
			meta, ok := container.Storage(path)
			if !ok {
				t.Fatalf("StoragePaths returned missing storage %q", path)
			}
			var parts []string
			if path != "" {
				parts = strings.Split(path, "/")
			}
			w.AddStorage(parts, meta)
		}
		for _, path := range container.Paths() {
			stream, ok := container.Stream(path)
			if !ok {
				t.Fatalf("Paths returned missing stream %q", path)
			}
			w.AddStream(strings.Split(path, "/"), stream)
		}
		rebuilt, err := w.Bytes()
		if err != nil {
			t.Fatalf("rebuild accepted container: %v", err)
		}
		roundTrip, err := Open(rebuilt)
		if err != nil {
			t.Fatalf("open rebuilt container: %v", err)
		}
		if roundTrip.Format() != container.Format() {
			t.Fatalf("format changed from %d to %d", container.Format(), roundTrip.Format())
		}
		if !slices.Equal(roundTrip.StoragePaths(), container.StoragePaths()) {
			t.Fatalf("storage paths changed from %q to %q", container.StoragePaths(), roundTrip.StoragePaths())
		}
		for _, path := range container.StoragePaths() {
			want, _ := container.Storage(path)
			got, ok := roundTrip.Storage(path)
			if !ok || got != want {
				t.Fatalf("storage %q metadata changed during round-trip", path)
			}
		}
		for _, path := range container.Paths() {
			want, _ := container.Stream(path)
			got, ok := roundTrip.Stream(path)
			if !ok || !bytes.Equal(got, want) {
				t.Fatalf("stream %q changed during round-trip", path)
			}
		}
	})
}
