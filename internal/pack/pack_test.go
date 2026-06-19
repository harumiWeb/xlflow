package pack

import (
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
	"github.com/harumiWeb/xlflow/internal/pack/vbaproject"
)

func readTestFile(t *testing.T, parts ...string) []byte {
	t.Helper()
	path := filepath.Join(append([]string{"testdata"}, parts...)...)
	b, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	return b
}

func readTestText(t *testing.T, parts ...string) string {
	t.Helper()
	return string(readTestFile(t, parts...))
}

func p1SourceModules(t *testing.T) []SourceModule {
	t.Helper()
	return []SourceModule{
		{Name: "Module1", Type: ModuleTypeStandard, Source: readTestText(t, "disk", "p1", "modules", "Module1.bas")},
		{Name: "Class1", Type: ModuleTypeClass, Source: readTestText(t, "disk", "p1", "classes", "Class1.cls")},
	}
}

func TestGenerateVBAProjectGoldenAndSources(t *testing.T) {
	template := readTestFile(t, "corpus", "p1_compiled.bin")
	out, err := GenerateVBAProject(template, p1SourceModules(t))
	if err != nil {
		t.Fatal(err)
	}
	want := readTestFile(t, "golden", "p1_pack.bin")
	if !bytes.Equal(out, want) {
		t.Fatalf("generated vbaProject.bin drifted from golden: got %d bytes, want %d", len(out), len(want))
	}

	project, err := vbaproject.Read(out)
	if err != nil {
		t.Fatalf("read generated: %v", err)
	}
	got := map[string]string{}
	existing := map[string]*vbaproject.Module{}
	for _, m := range project.Modules {
		got[m.Name] = m.Source
		module := m
		existing[m.Name] = &module
	}
	wantModule, err := vbaproject.NormalizeModuleSource(vbaproject.ModuleStd, readTestText(t, "disk", "p1", "modules", "Module1.bas"), existing["Module1"])
	if err != nil {
		t.Fatal(err)
	}
	wantClass, err := vbaproject.NormalizeModuleSource(vbaproject.ModuleClass, readTestText(t, "disk", "p1", "classes", "Class1.cls"), existing["Class1"])
	if err != nil {
		t.Fatal(err)
	}
	if got["Module1"] != wantModule {
		t.Fatalf("Module1 source mismatch after decompress")
	}
	if got["Class1"] != wantClass {
		t.Fatalf("Class1 source mismatch after decompress")
	}
}

func TestGenerateVBAProjectReplacesUnambiguousDocuments(t *testing.T) {
	template := readTestFile(t, "corpus", "p1_compiled.bin")
	sources := append(p1SourceModules(t),
		SourceModule{Name: "Sheet1", Type: ModuleTypeDocument, Source: readTestText(t, "disk", "p1", "workbook", "Sheet1.bas")},
		SourceModule{Name: "ThisWorkbook", Type: ModuleTypeDocument, Source: readTestText(t, "disk", "p1", "workbook", "ThisWorkbook.bas")},
	)
	out, err := GenerateVBAProject(template, sources)
	if err != nil {
		t.Fatal(err)
	}
	project, err := vbaproject.Read(out)
	if err != nil {
		t.Fatalf("read generated: %v", err)
	}
	for _, m := range project.Modules {
		if m.Type != vbaproject.ModuleDocument {
			continue
		}
		if !strings.HasPrefix(m.Source, `Attribute VB_Name = "`) {
			t.Fatalf("%s document header was not preserved:\n%s", m.Name, m.Source)
		}
		if !strings.Contains(m.Source, "Debug.Print") {
			t.Fatalf("%s document body was not replaced from disk source:\n%s", m.Name, m.Source)
		}
	}
}

func TestGenerateVBAProjectPreservesOpaqueStreams(t *testing.T) {
	template := readTestFile(t, "corpus", "p6_nested_form.bin")
	orig, err := cfb.Open(template)
	if err != nil {
		t.Fatal(err)
	}
	out, err := GenerateVBAProject(template, nil)
	if err != nil {
		t.Fatal(err)
	}
	got, err := cfb.Open(out)
	if err != nil {
		t.Fatal(err)
	}
	for _, path := range orig.Paths() {
		if !strings.HasPrefix(path, "UserForm1/") && path != "PROJECTwm" {
			continue
		}
		wantStream, _ := orig.Stream(path)
		gotStream, ok := got.Stream(path)
		if !ok {
			t.Fatalf("%s was dropped", path)
		}
		if !bytes.Equal(gotStream, wantStream) {
			t.Fatalf("%s was not preserved byte-for-byte", path)
		}
	}
}

func TestGenerateVBAProjectRoundTripStable(t *testing.T) {
	template := readTestFile(t, "corpus", "p2_refs.bin")
	out, err := GenerateVBAProject(template, nil)
	if err != nil {
		t.Fatal(err)
	}
	project, err := vbaproject.Read(out)
	if err != nil {
		t.Fatalf("read generated: %v", err)
	}
	out2, err := vbaproject.Write(project)
	if err != nil {
		t.Fatalf("write generated: %v", err)
	}
	if !bytes.Equal(out, out2) {
		t.Fatal("read -> write output is not stable")
	}
}

func TestGenerateVBAProjectTypedErrors(t *testing.T) {
	template := readTestFile(t, "corpus", "p1_compiled.bin")
	cases := []struct {
		name    string
		bin     []byte
		sources []SourceModule
		want    error
	}{
		{
			name: "protected project",
			bin:  readTestFile(t, "corpus", "p3_protected.bin"),
			want: ErrProtectedProject,
		},
		{
			name: "signed project",
			bin:  signedFixture(t, template),
			want: ErrSignedProject,
		},
		{
			name: "form generation",
			bin:  template,
			sources: []SourceModule{{
				Name: "UserForm1", Type: ModuleTypeForm, Source: "VERSION 5.00\r\nBegin VB.UserForm UserForm1\r\nEnd\r\n",
			}},
			want: ErrUserFormGenerationUnsupported,
		},
		{
			name:    "unknown layout",
			bin:     template,
			sources: []SourceModule{{Name: "Missing", Type: ModuleTypeStandard, Source: `Attribute VB_Name = "Missing"` + "\r\n"}},
			want:    ErrAmbiguousLayout,
		},
		{
			name:    "ambiguous duplicate source",
			bin:     template,
			sources: []SourceModule{{Name: "Module1", Type: ModuleTypeStandard}, {Name: "Module1", Type: ModuleTypeStandard}},
			want:    ErrAmbiguousLayout,
		},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			_, err := GenerateVBAProject(tc.bin, tc.sources)
			if !errors.Is(err, tc.want) {
				t.Fatalf("error = %v, want errors.Is(..., %v)", err, tc.want)
			}
		})
	}
}

func TestGenerateVBAProjectMissingTemplateModuleExplainsMVPLimitation(t *testing.T) {
	template := readTestFile(t, "corpus", "p1_compiled.bin")
	_, err := GenerateVBAProject(template, []SourceModule{{
		Name: "Missing", Type: ModuleTypeStandard, Source: `Attribute VB_Name = "Missing"` + "\r\n",
	}})
	if !errors.Is(err, ErrAmbiguousLayout) {
		t.Fatalf("error = %v, want errors.Is(..., ErrAmbiguousLayout)", err)
	}
	want := `source module "Missing" is not in the template; pack updates existing modules only and cannot add new modules in the experimental MVP`
	if !strings.Contains(err.Error(), want) {
		t.Fatalf("error message = %q, want it to contain %q", err.Error(), want)
	}
}

func signedFixture(t *testing.T, template []byte) []byte {
	t.Helper()
	project, err := vbaproject.Read(template)
	if err != nil {
		t.Fatal(err)
	}
	project.RawStreams["_VBA_PROJECT_CUR/VBAProjectSignature"] = []byte("signature sentinel")
	out, err := vbaproject.Write(project)
	if err != nil {
		t.Fatal(err)
	}
	return out
}
