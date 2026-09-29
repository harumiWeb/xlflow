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

func TestWriteAcceptsCodepageRepresentableModuleName(t *testing.T) {
	p, err := Read(loadBin(t, "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	// The project code page is 932, so a Japanese module name is representable
	// in both the MBCS record and its paired UTF-16 record.
	p.Modules[len(p.Modules)-1].Name = "\u30e2\u30b8\u30e5\u30fc\u30eb"
	p.Modules[len(p.Modules)-1].StreamName = "\u30e2\u30b8\u30e5\u30fc\u30eb"
	out, err := Write(p)
	if err != nil {
		t.Fatalf("write: %v", err)
	}
	p2, err := Read(out)
	if err != nil {
		t.Fatalf("re-read: %v", err)
	}
	got := p2.Modules[len(p2.Modules)-1]
	if got.Name != "\u30e2\u30b8\u30e5\u30fc\u30eb" || got.StreamName != "\u30e2\u30b8\u30e5\u30fc\u30eb" {
		t.Errorf("non-ASCII names drifted: Name=%q StreamName=%q", got.Name, got.StreamName)
	}
}

func TestWriteRejectsUnrepresentableModuleName(t *testing.T) {
	p, err := Read(loadBin(t, "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	// CP932 cannot represent an emoji; the writer must fail loudly rather than
	// emit replacement bytes into the MBCS records.
	p.Modules[len(p.Modules)-1].Name = "mod\U0001F600"
	if _, err := Write(p); err == nil {
		t.Error("an unrepresentable module name should return an error")
	}
	p.Modules[len(p.Modules)-1].Name = "Module1"
	p.Modules[len(p.Modules)-1].StreamName = "stream\U0001F600"
	if _, err := Write(p); err == nil {
		t.Error("an unrepresentable stream name should return an error")
	}
}

func TestWritePreservesModuleMetadata(t *testing.T) {
	p, err := Read(loadBin(t, "p4_form.bin"))
	if err != nil {
		t.Fatal(err)
	}
	// Stamp metadata onto one module and require it to survive read/write/read.
	idx := 0
	for i := range p.Modules {
		if p.Modules[i].Type == ModuleStd {
			idx = i
		}
	}
	p.Modules[idx].DocString = "説明 \u30c9\u30ad\u30e5\u30e1\u30f3\u30c8"
	p.Modules[idx].HelpContext = 42
	p.Modules[idx].ReadOnly = true
	out, err := Write(p)
	if err != nil {
		t.Fatalf("write: %v", err)
	}
	p2, err := Read(out)
	if err != nil {
		t.Fatalf("re-read: %v", err)
	}
	got := p2.Modules[idx]
	if got.DocString != p.Modules[idx].DocString {
		t.Errorf("DocString drifted: %q -> %q", p.Modules[idx].DocString, got.DocString)
	}
	if got.HelpContext != 42 || !got.ReadOnly {
		t.Errorf("metadata drifted: HelpContext=%d ReadOnly=%v", got.HelpContext, got.ReadOnly)
	}
	// The template's private flags (class + form) must also be preserved.
	for i := range p.Modules {
		if p.Modules[i].Private != p2.Modules[i].Private {
			t.Errorf("module %s Private drifted: %v -> %v", p.Modules[i].Name, p.Modules[i].Private, p2.Modules[i].Private)
		}
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

func TestWriteRejectsCFBEquivalentStreamNames(t *testing.T) {
	// CFB directory entry names compare case-insensitively on upcased UTF-16
	// code units, which is not the same equivalence as Unicode lowercase: the
	// Turkish dotless i ("\u0131") upcases to ASCII "I". Two modules whose
	// stream names differ only that way would collide in the container.
	p := &Project{
		Props: ProjectProps{CodePage: 1254},
		Modules: []Module{
			{Name: "I", StreamName: "I", Type: ModuleStd, Source: "Attribute VB_Name = \"I\"\r\n"},
			{Name: "\u0131", StreamName: "\u0131", Type: ModuleStd, Source: "Attribute VB_Name = \"\u0131\"\r\n"},
		},
	}
	if _, err := Write(p); err == nil {
		t.Fatal("stream names that collide in the CFB directory should fail")
	}
}

func TestWriteRejectsReservedStreamNames(t *testing.T) {
	for _, name := range []string{"dir", "DIR", "_VBA_PROJECT"} {
		t.Run(name, func(t *testing.T) {
			p := &Project{
				Props: ProjectProps{CodePage: 1252},
				Modules: []Module{
					{Name: name, StreamName: name, Type: ModuleStd, Source: "Attribute VB_Name = \"" + name + "\"\r\n"},
				},
			}
			if _, err := Write(p); err == nil {
				t.Fatal("a module stream name reserved by the writer should fail")
			}
		})
	}
}
