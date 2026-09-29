package vbaproject

import (
	"bytes"
	"os"
	"testing"
)

func loadBin(t *testing.T, name string) []byte {
	t.Helper()
	b, err := os.ReadFile("testdata/corpus/" + name)
	if err != nil {
		t.Fatal(err)
	}
	return b
}

func TestWriteRoundTrip(t *testing.T) {
	for _, name := range []string{"p1_compiled.bin", "p2_refs.bin", "p4_form.bin", "p5_mbcs.bin", "p6_nested_form.bin"} {
		t.Run(name, func(t *testing.T) {
			p, err := Read(loadBin(t, name))
			if err != nil {
				t.Fatal(err)
			}
			out, err := Write(p)
			if err != nil {
				t.Fatalf("write: %v", err)
			}
			p2, err := Read(out)
			if err != nil {
				t.Fatalf("re-read: %v", err)
			}
			if !bytes.Equal(p.ProjectInfoRaw, p2.ProjectInfoRaw) {
				t.Error("ProjectInfoRaw drift")
			}
			if !bytes.Equal(p.ReferencesRaw, p2.ReferencesRaw) {
				t.Error("ReferencesRaw drift")
			}
			if !bytes.Equal(p.ProjectStreamRaw, p2.ProjectStreamRaw) {
				t.Error("ProjectStreamRaw drift")
			}
			if len(p.Modules) != len(p2.Modules) {
				t.Fatalf("module count %d -> %d", len(p.Modules), len(p2.Modules))
			}
			for i := range p.Modules {
				a, b := p.Modules[i], p2.Modules[i]
				if a.Name != b.Name || a.Type != b.Type || a.Source != b.Source {
					t.Errorf("module %d (%s) source/type drift", i, a.Name)
				}
			}
		})
	}
}

func TestWriteRejectsProtected(t *testing.T) {
	p, err := Read(loadBin(t, "p3_protected.bin"))
	if err != nil {
		t.Fatal(err)
	}
	if !p.Protection.IsProtected {
		t.Skip("p3 was not detected as protected (precondition broken)")
	}
	if _, err := Write(p); err == nil {
		t.Error("a protected project should return an error")
	}
}

func TestWriteRejectsNonASCIIModuleName(t *testing.T) {
	p, err := Read(loadBin(t, "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	// Changing a module name to non-ASCII cannot be written losslessly to both the dir MBCS field and the
	// UNICODE field (a breeding ground for utf16le non-BMP truncation and CODEPAGE misencoding).
	p.Modules[len(p.Modules)-1].Name = "\u30e2\u30b8\u30e5\u30fc\u30eb"
	if _, err := Write(p); err == nil {
		t.Error("a non-ASCII module name should return an error")
	}
}

func TestWriteSupportsStandardAndClassTopologyChanges(t *testing.T) {
	p, err := Read(loadBin(t, "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	p.Modules = append(p.Modules[:2],
		Module{Name: "RenamedModule", StreamName: "RenamedModule", Type: ModuleStd, Source: "Attribute VB_Name = \"RenamedModule\"\r\n"},
		Module{Name: "AddedClass", StreamName: "AddedClass", Type: ModuleClass, Source: "Attribute VB_Name = \"AddedClass\"\r\n"},
	)
	out, err := Write(p)
	if err != nil {
		t.Fatal(err)
	}
	got, err := Read(out)
	if err != nil {
		t.Fatal(err)
	}
	want := map[string]ModuleType{
		"ThisWorkbook":  ModuleDocument,
		"Sheet1":        ModuleDocument,
		"RenamedModule": ModuleStd,
		"AddedClass":    ModuleClass,
	}
	if len(got.Modules) != len(want) {
		t.Fatalf("module count = %d, want %d", len(got.Modules), len(want))
	}
	for _, module := range got.Modules {
		wantType, ok := want[module.Name]
		if !ok || wantType != module.Type {
			t.Errorf("module %q type = %v, want %v (present=%v)", module.Name, module.Type, wantType, ok)
		}
	}
}

func TestWriteRejectsTemplateOwnedTopologyChanges(t *testing.T) {
	p, err := Read(loadBin(t, "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	p.Modules = p.Modules[1:]
	if _, err := Write(p); err == nil {
		t.Fatal("removing a document module should fail")
	}
}

func TestWriteRejectsDuplicateNamesAndStreams(t *testing.T) {
	for _, tc := range []struct {
		name   string
		module Module
	}{
		{name: "name", module: Module{Name: "module1", StreamName: "Other", Type: ModuleStd}},
		{name: "stream", module: Module{Name: "Other", StreamName: "module1", Type: ModuleStd}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			p, err := Read(loadBin(t, "p1_compiled.bin"))
			if err != nil {
				t.Fatal(err)
			}
			p.Modules = append(p.Modules, tc.module)
			if _, err := Write(p); err == nil {
				t.Fatal("duplicate module identity should fail")
			}
		})
	}
}
