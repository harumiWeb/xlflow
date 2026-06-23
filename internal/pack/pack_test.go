package pack

import (
	"archive/zip"
	"bytes"
	"errors"
	"io"
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
			// UserForm1 is absent from p1_compiled.bin, so this is a new form: pack cannot
			// author a form's designer storage, so it is rejected as UserForm generation.
			name: "new form on form-less template",
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

func TestBuildWorkbookReplacesVBAProjectAndPreservesOtherEntries(t *testing.T) {
	templateBin := readTestFile(t, "corpus", "p1_compiled.bin")
	templateXlsm := buildWorkbookFixture(t, templateBin)
	out, meta, err := BuildWorkbook(templateXlsm, p1SourceModules(t))
	if err != nil {
		t.Fatal(err)
	}
	if meta.Standard != 1 || meta.Class != 1 || meta.Document != 0 {
		t.Fatalf("module counts = %+v, want standard=1 class=1 document=0", meta)
	}
	if meta.CarriedStreams == 0 {
		t.Fatalf("expected carried stream count from template, got %+v", meta)
	}

	wantEntries := zipEntries(t, templateXlsm)
	gotEntries := zipEntries(t, out)
	if len(gotEntries) != len(wantEntries) {
		t.Fatalf("entry count = %d, want %d", len(gotEntries), len(wantEntries))
	}
	for i := range wantEntries {
		want, got := wantEntries[i], gotEntries[i]
		if got.name != want.name {
			t.Fatalf("entry[%d] name = %q, want %q", i, got.name, want.name)
		}
		if got.method != want.method {
			t.Fatalf("%s method = %d, want %d", got.name, got.method, want.method)
		}
		if got.name == "xl/vbaProject.bin" {
			if bytes.Equal(got.body, want.body) {
				t.Fatal("xl/vbaProject.bin was not replaced")
			}
			continue
		}
		if !bytes.Equal(got.body, want.body) {
			t.Fatalf("%s body was not preserved byte-for-byte", got.name)
		}
	}
	if _, err := vbaproject.Read(mustEntry(t, out, "xl/vbaProject.bin")); err != nil {
		t.Fatalf("regenerated vbaProject.bin is not re-readable: %v", err)
	}
}

func TestBuildWorkbookRejectsMissingVBAProject(t *testing.T) {
	templateXlsm := buildWorkbookFixture(t, nil)
	_, _, err := BuildWorkbook(templateXlsm, nil)
	if !errors.Is(err, ErrAmbiguousLayout) {
		t.Fatalf("error = %v, want errors.Is(..., ErrAmbiguousLayout)", err)
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

type zipEntryForTest struct {
	name   string
	method uint16
	body   []byte
}

func buildWorkbookFixture(t *testing.T, vbaProject []byte) []byte {
	t.Helper()
	var buf bytes.Buffer
	writer := zip.NewWriter(&buf)
	add := func(name string, method uint16, body []byte) {
		t.Helper()
		header := &zip.FileHeader{Name: name, Method: method}
		w, err := writer.CreateHeader(header)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := w.Write(body); err != nil {
			t.Fatal(err)
		}
	}
	add("[Content_Types].xml", zip.Deflate, []byte(`<Types></Types>`))
	add("xl/workbook.xml", zip.Store, []byte(`<workbook/>`))
	if vbaProject != nil {
		add("xl/vbaProject.bin", zip.Deflate, vbaProject)
	}
	add("docProps/core.xml", zip.Deflate, []byte(`<core/>`))
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

func zipEntries(t *testing.T, data []byte) []zipEntryForTest {
	t.Helper()
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	entries := make([]zipEntryForTest, 0, len(reader.File))
	for _, f := range reader.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatal(err)
		}
		body, err := io.ReadAll(rc)
		closeErr := rc.Close()
		if err != nil {
			t.Fatal(err)
		}
		if closeErr != nil {
			t.Fatal(closeErr)
		}
		entries = append(entries, zipEntryForTest{name: f.Name, method: f.Method, body: body})
	}
	return entries
}

func mustEntry(t *testing.T, data []byte, name string) []byte {
	t.Helper()
	for _, entry := range zipEntries(t, data) {
		if entry.name == name {
			return entry.body
		}
	}
	t.Fatalf("missing zip entry %s", name)
	return nil
}

func TestGenerateVBAProjectUpdatesExistingFormCodeBehind(t *testing.T) {
	template := readTestFile(t, "corpus", "p4_form.bin")
	src := SourceModule{
		Name: "UserForm1",
		Type: ModuleTypeForm,
		Source: "VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1\r\n" +
			"   Caption = \"x\"\r\nEnd\r\n" +
			"Attribute VB_Name = \"UserForm1\"\r\n" +
			"Attribute VB_GlobalNameSpace = False\r\n" +
			"Attribute VB_Creatable = False\r\n" +
			"Attribute VB_PredeclaredId = True\r\n" +
			"Attribute VB_Exposed = False\r\n" +
			"Private Sub CommandButton1_Click()\r\n    Debug.Print \"PACKED_EDIT\"\r\nEnd Sub\r\n",
	}
	out, err := GenerateVBAProject(template, []SourceModule{src})
	if err != nil {
		t.Fatalf("GenerateVBAProject: %v", err)
	}
	in, err := vbaproject.Read(template)
	if err != nil {
		t.Fatal(err)
	}
	got, err := vbaproject.Read(out)
	if err != nil {
		t.Fatal(err)
	}
	var sawEdit bool
	for _, m := range got.Modules {
		if m.Type == vbaproject.ModuleForm && strings.Contains(m.Source, "PACKED_EDIT") {
			sawEdit = true
		}
	}
	if !sawEdit {
		t.Error("edited code-behind not found in output form module")
	}
	if len(in.RawStreams) != len(got.RawStreams) {
		t.Fatalf("rawstream count changed: in=%d out=%d", len(in.RawStreams), len(got.RawStreams))
	}
	for key, a := range in.RawStreams {
		b, ok := got.RawStreams[key]
		if !ok || string(a) != string(b) {
			t.Errorf("designer stream %q not carried byte-for-byte (present=%v)", key, ok)
		}
	}
}

func TestGenerateVBAProjectRejectsNewFormWithFormError(t *testing.T) {
	template := readTestFile(t, "corpus", "p4_form.bin")
	src := SourceModule{
		Name: "BrandNewForm",
		Type: ModuleTypeForm,
		Source: "VERSION 5.00\r\nBegin {GUID} BrandNewForm\r\nEnd\r\n" +
			"Attribute VB_Name = \"BrandNewForm\"\r\nPrivate Sub a()\r\nEnd Sub\r\n",
	}
	_, err := GenerateVBAProject(template, []SourceModule{src})
	if !errors.Is(err, ErrUserFormGenerationUnsupported) {
		t.Fatalf("want ErrUserFormGenerationUnsupported, got %v", err)
	}
	if errors.Is(err, ErrAmbiguousLayout) {
		t.Fatal("new form must not surface as ErrAmbiguousLayout")
	}
}
