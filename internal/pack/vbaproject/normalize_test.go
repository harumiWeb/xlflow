package vbaproject

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// loadDisk reads a disk-form golden (testdata/disk/<book>/<sub>/<file>).
// The disk form is UTF-8, so it is read directly with string().
func loadDisk(t *testing.T, book, sub, file string) string {
	t.Helper()
	b, err := os.ReadFile(filepath.Join("testdata", "disk", book, sub, file))
	if err != nil {
		t.Fatal(err)
	}
	return string(b)
}

// loadExpected reads an in-bin-form golden (testdata/expected/<book>/<name>.txt).
func loadExpected(t *testing.T, book, name string) string {
	t.Helper()
	b, err := os.ReadFile(filepath.Join("testdata", "expected", book, name+".txt"))
	if err != nil {
		t.Fatal(err)
	}
	return string(b)
}

func TestToCRLF(t *testing.T) {
	cases := []struct{ in, want string }{
		{"a\nb\n", "a\r\nb\r\n"},     // LF -> CRLF
		{"a\r\nb\r\n", "a\r\nb\r\n"}, // already CRLF stays unchanged (idempotent)
		{"a\r\nb\nc", "a\r\nb\r\nc"}, // mixed -> CRLF
	}
	for _, c := range cases {
		if got := toCRLF(c.in); got != c.want {
			t.Errorf("toCRLF(%q) = %q, want %q", c.in, got, c.want)
		}
	}
}

func TestNormalizeStd(t *testing.T) {
	// The disk dir uses short names (p1/p5); the expected dir uses in-bin naming (p1_compiled/p5_mbcs).
	cases := []struct{ diskBook, expBook string }{
		{"p1", "p1_compiled"},
		{"p5", "p5_mbcs"}, // confirms that Japanese comments pass through as UTF-8
	}
	for _, c := range cases {
		t.Run(c.diskBook, func(t *testing.T) {
			disk := loadDisk(t, c.diskBook, "modules", "Module1.bas")
			got, err := NormalizeModuleSource(ModuleStd, disk, nil)
			if err != nil {
				t.Fatal(err)
			}
			want := loadExpected(t, c.expBook, "Module1")
			if got != want {
				t.Errorf("std normalize mismatch\n got: %q\nwant: %q", got, want)
			}
		})
	}
}

func TestNormalizeStdMissingVBName(t *testing.T) {
	_, err := NormalizeModuleSource(ModuleStd, "Public Sub Foo()\r\nEnd Sub\r\n", nil)
	if err == nil {
		t.Fatal("a std disk form without Attribute VB_Name should return an error")
	}
}

func TestNormalizeClass(t *testing.T) {
	disk := loadDisk(t, "p1", "classes", "Class1.cls")
	got, err := NormalizeModuleSource(ModuleClass, disk, nil)
	if err != nil {
		t.Fatal(err)
	}
	want := loadExpected(t, "p1_compiled", "Class1")
	if got != want {
		t.Errorf("class normalize mismatch\n--- got ---\n%s\n--- want ---\n%s", got, want)
	}
}

func TestNormalizeDocument(t *testing.T) {
	p, err := Read(loadBin(t, "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	existing := map[string]*Module{}
	for i := range p.Modules {
		existing[p.Modules[i].Name] = &p.Modules[i]
	}
	for _, name := range []string{"ThisWorkbook", "Sheet1"} {
		t.Run(name, func(t *testing.T) {
			disk := loadDisk(t, "p1", "workbook", name+".bas")
			got, err := NormalizeModuleSource(ModuleDocument, disk, existing[name])
			if err != nil {
				t.Fatal(err)
			}
			want := loadExpected(t, "p1_compiled", name)
			if got != want {
				t.Errorf("document normalize mismatch (%s)\n--- got ---\n%s\n--- want ---\n%s", name, got, want)
			}
		})
	}
}

func TestNormalizeErrors(t *testing.T) {
	// The missing-VB_Name case for std is already covered by TestNormalizeStdMissingVBName, so it is not handled here.
	cases := []struct {
		name     string
		mt       ModuleType
		disk     string
		existing *Module
	}{
		{"form is unsupported", ModuleForm, "anything", nil},
		{"unknown ModuleType", ModuleType(99), "anything", nil},
		{"document requires existing", ModuleDocument, "Private Sub X()\r\nEnd Sub\r\n", nil},
		{"class requires a VERSION header", ModuleClass, "Attribute VB_Name = \"C\"\r\n", nil},
		{"document existing has no header", ModuleDocument, "code\r\n", &Module{Source: "code only\r\n"}},
		{"class requires the anchors (VB_Name/VB_Exposed)", ModuleClass, "VERSION 1.0 CLASS\r\nBEGIN\r\nEND\r\nPublic Sub X()\r\nEnd Sub\r\n", nil},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			if _, err := NormalizeModuleSource(c.mt, c.disk, c.existing); err == nil {
				t.Errorf("%s: expected an error, got nil", c.name)
			}
		})
	}
}

func TestNormalizeEndToEnd(t *testing.T) {
	p, err := Read(loadBin(t, "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	layout := map[ModuleType]struct{ sub, ext string }{
		ModuleStd:      {"modules", ".bas"},
		ModuleClass:    {"classes", ".cls"},
		ModuleDocument: {"workbook", ".bas"},
	}
	for i := range p.Modules {
		m := &p.Modules[i]
		l, ok := layout[m.Type]
		if !ok {
			t.Fatalf("%s: unsupported type %d", m.Name, m.Type)
		}
		disk := loadDisk(t, "p1", l.sub, m.Name+l.ext)
		norm, err := NormalizeModuleSource(m.Type, disk, m)
		if err != nil {
			t.Fatalf("%s: normalize: %v", m.Name, err)
		}
		m.Source = norm
	}
	out, err := Write(p)
	if err != nil {
		t.Fatalf("write: %v", err)
	}
	p2, err := Read(out)
	if err != nil {
		t.Fatalf("re-read: %v", err)
	}
	for _, m := range p2.Modules {
		want := loadExpected(t, "p1_compiled", m.Name)
		if m.Source != want {
			t.Errorf("%s: e2e round-trip does not match the in-bin golden (got %d chars)", m.Name, len(m.Source))
		}
	}
}

// Edit case: shows that replacing the code body preserves the 8 attribute-header lines from existing.
func TestNormalizeDocumentEdited(t *testing.T) {
	p, err := Read(loadBin(t, "p1_compiled.bin"))
	if err != nil {
		t.Fatal(err)
	}
	var tw *Module
	for i := range p.Modules {
		if p.Modules[i].Name == "ThisWorkbook" {
			tw = &p.Modules[i]
		}
	}
	if tw == nil {
		t.Fatal("ThisWorkbook is not in the corpus")
	}
	editedDisk := "Private Sub Workbook_Open()\r\n    Debug.Print \"EDITED_P2\"\r\nEnd Sub\r\n"
	got, err := NormalizeModuleSource(ModuleDocument, editedDisk, tw)
	if err != nil {
		t.Fatal(err)
	}
	want := strings.Join([]string{
		`Attribute VB_Name = "ThisWorkbook"`,
		`Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"`,
		`Attribute VB_GlobalNameSpace = False`,
		`Attribute VB_Creatable = False`,
		`Attribute VB_PredeclaredId = True`,
		`Attribute VB_Exposed = True`,
		`Attribute VB_TemplateDerived = False`,
		`Attribute VB_Customizable = True`,
	}, "\r\n") + "\r\n" + editedDisk
	if got != want {
		t.Errorf("edited document mismatch\n--- got ---\n%s\n--- want ---\n%s", got, want)
	}
}

func TestNormalizeFormKeepsExistingHeaderSwapsCode(t *testing.T) {
	existing := &Module{
		Name: "UserForm1",
		Type: ModuleForm,
		Source: "Attribute VB_Name = \"UserForm1\"\r\n" +
			"Attribute VB_Base = \"0{5A352896-124D-420C-80E8-2C4BC088CC1C}{95A506B3-11F2-4259-860C-CB0FC1817653}\"\r\n" +
			"Attribute VB_GlobalNameSpace = False\r\n" +
			"Attribute VB_Creatable = False\r\n" +
			"Attribute VB_PredeclaredId = True\r\n" +
			"Attribute VB_Exposed = False\r\n" +
			"Attribute VB_TemplateDerived = False\r\n" +
			"Attribute VB_Customizable = False\r\n" +
			"Private Sub CommandButton1_Click()\r\n" +
			"    Debug.Print \"OLD\"\r\n" +
			"End Sub\r\n",
	}
	disk := "VERSION 5.00\r\n" +
		"Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1\r\n" +
		"   Caption = \"Form\"\r\n" +
		"   ClientHeight = 3000\r\n" +
		"End\r\n" +
		"Attribute VB_Name = \"UserForm1\"\r\n" +
		"Attribute VB_GlobalNameSpace = False\r\n" +
		"Attribute VB_Creatable = False\r\n" +
		"Attribute VB_PredeclaredId = True\r\n" +
		"Attribute VB_Exposed = False\r\n" +
		"Private Sub CommandButton1_Click()\r\n" +
		"    Debug.Print \"NEW\"\r\n" +
		"End Sub\r\n"
	want := "Attribute VB_Name = \"UserForm1\"\r\n" +
		"Attribute VB_Base = \"0{5A352896-124D-420C-80E8-2C4BC088CC1C}{95A506B3-11F2-4259-860C-CB0FC1817653}\"\r\n" +
		"Attribute VB_GlobalNameSpace = False\r\n" +
		"Attribute VB_Creatable = False\r\n" +
		"Attribute VB_PredeclaredId = True\r\n" +
		"Attribute VB_Exposed = False\r\n" +
		"Attribute VB_TemplateDerived = False\r\n" +
		"Attribute VB_Customizable = False\r\n" +
		"Private Sub CommandButton1_Click()\r\n" +
		"    Debug.Print \"NEW\"\r\n" +
		"End Sub\r\n"
	got, err := NormalizeModuleSource(ModuleForm, disk, existing)
	if err != nil {
		t.Fatalf("NormalizeModuleSource: %v", err)
	}
	if got != want {
		t.Errorf("normalizeForm mismatch:\n got=%q\nwant=%q", got, want)
	}
}

func TestNormalizeFormErrorsWithoutExisting(t *testing.T) {
	_, err := NormalizeModuleSource(ModuleForm, "Attribute VB_Name = \"X\"\r\ncode\r\n", nil)
	if err == nil {
		t.Fatal("expected error when existing is nil")
	}
}

func TestNormalizeFormErrorsWhenFRMHasNoAttribute(t *testing.T) {
	existing := &Module{Source: "Attribute VB_Name = \"X\"\r\nold\r\n"}
	_, err := NormalizeModuleSource(ModuleForm, "VERSION 5.00\r\nBegin x\r\nEnd\r\n", existing)
	if err == nil {
		t.Fatal("expected error when .frm has no Attribute VB_ line")
	}
}
