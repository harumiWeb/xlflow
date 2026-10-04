package oforms

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
)

func FuzzReadForm(f *testing.F) {
	for _, fixture := range corpusProjectFixtures() {
		body, err := os.ReadFile(fixture)
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
			form, err := ReadForm(container, name, 932)
			if err != nil {
				continue
			}
			serialized, err := SerializeForm(form, 932)
			if err != nil {
				t.Fatalf("SerializeForm after successful ReadForm: %v", err)
			}
			assertSerializedSubtree(t, container, name, serialized)
		}
	})
}

func FuzzParseFormStream(f *testing.F) {
	for _, fixture := range corpusProjectFixtures() {
		body, err := os.ReadFile(fixture)
		if err != nil {
			f.Fatal(err)
		}
		container, err := cfb.Open(body)
		if err != nil {
			f.Fatal(err)
		}
		for _, formName := range DiscoverForms(container) {
			prefix := formName + "/"
			for _, path := range container.Paths() {
				if strings.HasPrefix(path, prefix) && strings.HasSuffix(path, "/f") || path == prefix+"f" {
					stream, _ := container.Stream(path)
					f.Add(stream)
				}
			}
		}
	}
	f.Fuzz(func(t *testing.T, data []byte) {
		_, _ = parseFormStream(data, "fuzz", 932)
	})
}

func corpusProjectFixtures() []string {
	return []string{
		filepath.Join("..", "..", "..", "pack", "vbaproject", "testdata", "corpus", "p4_form.bin"),
		filepath.Join("..", "..", "..", "pack", "vbaproject", "testdata", "corpus", "p6_nested_form.bin"),
		filepath.Join("..", "compiler", "testdata", "excel-authored", "00-baseline.bin"),
		filepath.Join("..", "compiler", "testdata", "frame-excel-authored", "baseline.bin"),
		filepath.Join("..", "compiler", "testdata", "generation-excel-authored", "baseline.bin"),
		filepath.Join("..", "compiler", "testdata", "multipage-excel-authored", "baseline.bin"),
		filepath.Join("..", "compiler", "testdata", "multipage-excel-authored", "empty.bin"),
	}
}

func FuzzParseSiteRecord(f *testing.F) {
	for _, fixture := range corpusProjectFixtures() {
		container := openCorpusFixture(f, fixture)
		for _, name := range DiscoverForms(container) {
			form, err := ReadForm(container, name, 932)
			if err != nil {
				f.Fatal(err)
			}
			for _, level := range form.Levels {
				for _, site := range level.Sites {
					f.Add(bytes.Clone(site.Raw))
				}
			}
		}
	}
	f.Fuzz(func(t *testing.T, data []byte) {
		reader := newByteReader(data)
		_, _ = parseSite(reader, 932)
	})
}

func FuzzParseControlPropertyRecord(f *testing.F) {
	for _, fixture := range corpusProjectFixtures() {
		container := openCorpusFixture(f, fixture)
		for _, name := range DiscoverForms(container) {
			form, err := ReadForm(container, name, 932)
			if err != nil {
				f.Fatal(err)
			}
			for _, control := range form.Controls {
				addRecordFuzzSeed(f, control)
			}
		}
	}
	f.Fuzz(func(t *testing.T, cacheIndex uint16, data []byte) {
		spec := specsByCacheIndex[cacheIndex]
		if spec == nil {
			return
		}
		_, _, _ = parseRecord(data, spec, 932, len(data))
	})
}

func openCorpusFixture(t testing.TB, fixture string) *cfb.Container {
	t.Helper()
	body, err := os.ReadFile(fixture)
	if err != nil {
		t.Fatal(err)
	}
	container, err := cfb.Open(body)
	if err != nil {
		t.Fatal(err)
	}
	return container
}

func addRecordFuzzSeed(f *testing.F, control *Control) {
	if control.Record != nil && len(control.Record.Raw) > 0 {
		f.Add(control.CLSIDCacheIndex, bytes.Clone(control.Record.Raw))
	}
	for _, child := range control.Children {
		addRecordFuzzSeed(f, child)
	}
}
