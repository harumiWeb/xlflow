package vbaproject

import (
	"bytes"
	"encoding/binary"
	"os"
	"path/filepath"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
)

func loadCorpus(t *testing.T, book string) []byte {
	t.Helper()
	b, err := os.ReadFile(filepath.Join("testdata", "corpus", book+".bin"))
	if err != nil {
		t.Fatal(err)
	}
	return b
}

func TestReadP1Types(t *testing.T) {
	p, err := Read(loadCorpus(t, "p1_compiled"))
	if err != nil {
		t.Fatal(err)
	}
	want := map[string]ModuleType{
		"Module1": ModuleStd, "Class1": ModuleClass,
		"ThisWorkbook": ModuleDocument, "Sheet1": ModuleDocument,
	}
	got := map[string]ModuleType{}
	for _, m := range p.Modules {
		got[m.Name] = m.Type
	}
	for name, typ := range want {
		if got[name] != typ {
			t.Errorf("%s type = %d, want %d", name, got[name], typ)
		}
	}
	if p.Props.CodePage != 932 || p.Props.SysKind != 3 {
		t.Errorf("props codepage/syskind = %d/%d", p.Props.CodePage, p.Props.SysKind)
	}
	if p.Protection.CMG == "" || p.Protection.DPB == "" {
		t.Errorf("CMG/DPB is empty")
	}
}

func TestReadSourceMatchesOlevba(t *testing.T) {
	books := []string{"p1_compiled", "p2_refs", "p3_protected", "p4_form", "p5_mbcs"}
	for _, book := range books {
		p, err := Read(loadCorpus(t, book))
		if err != nil {
			t.Fatalf("%s: %v", book, err)
		}
		for _, m := range p.Modules {
			want, err := os.ReadFile(filepath.Join("testdata", "expected", book, m.Name+".txt"))
			if err != nil {
				t.Errorf("%s/%s: no expected golden", book, m.Name)
				continue
			}
			if m.Source != string(want) {
				t.Errorf("%s/%s: source does not match the olevba golden (got %d, want %d chars)",
					book, m.Name, len(m.Source), len(want))
			}
		}
	}
}

func TestReadReferences(t *testing.T) {
	p, err := Read(loadCorpus(t, "p4_form"))
	if err != nil {
		t.Fatal(err)
	}
	var names []string
	for _, r := range p.References {
		names = append(names, r.Name)
	}
	want := []string{"stdole", "Office", "MSForms"}
	if len(names) != len(want) {
		t.Fatalf("References = %v, want %v", names, want)
	}
	for i := range want {
		if names[i] != want[i] {
			t.Errorf("References[%d] = %q, want %q", i, names[i], want[i])
		}
	}
	if len(p.ReferencesRaw) == 0 {
		t.Error("ReferencesRaw is empty (verbatim span not captured)")
	}
}

func TestProtectedAndForm(t *testing.T) {
	prot, err := Read(loadCorpus(t, "p3_protected"))
	if err != nil {
		t.Fatal(err)
	}
	if !prot.Protection.IsProtected {
		t.Errorf("p3 should be IsProtected=true")
	}
	form, err := Read(loadCorpus(t, "p4_form"))
	if err != nil {
		t.Fatal(err)
	}
	var hasForm bool
	for _, m := range form.Modules {
		if m.Name == "UserForm1" && m.Type == ModuleForm {
			hasForm = true
		}
	}
	if !hasForm {
		t.Errorf("p4 should have UserForm1=ModuleForm")
	}
	// Unprotected -> IsProtected=false
	p1, err := Read(loadCorpus(t, "p1_compiled"))
	if err != nil {
		t.Fatal(err)
	}
	if p1.Protection.IsProtected {
		t.Errorf("p1 should be IsProtected=false")
	}
}

func TestIsProtectedClassifiesCorpus(t *testing.T) {
	// p3 is the protected sample; the rest are unprotected. CMG comes straight
	// from the PROJECT stream of each real fixture, so this is the external
	// oracle for the decryption + bit interpretation.
	cases := []struct {
		book string
		want bool
	}{
		{"p1_compiled", false},
		{"p2_refs", false},
		{"p3_protected", true},
		{"p5_mbcs", false},
		{"p6_nested_form", false},
	}
	for _, c := range cases {
		t.Run(c.book, func(t *testing.T) {
			p, err := Read(loadCorpus(t, c.book))
			if err != nil {
				t.Fatalf("Read(%s): %v", c.book, err)
			}
			if p.Protection.IsProtected != c.want {
				t.Errorf("%s: IsProtected=%v want %v (CMG=%q)",
					c.book, p.Protection.IsProtected, c.want, p.Protection.CMG)
			}
		})
	}
}

func TestIsProtectedEdgeCases(t *testing.T) {
	// Empty CMG => no protection record => unprotected, no error.
	got, err := isProtected("")
	if err != nil || got {
		t.Fatalf("empty CMG: got (%v, %v), want (false, nil)", got, err)
	}
	// Non-hex CMG => fail-loud.
	if _, err := isProtected("zz"); err == nil {
		t.Error("non-hex CMG: expected error, got nil")
	}
}

func TestReadPreservationRawSpans(t *testing.T) {
	bin, err := os.ReadFile("testdata/corpus/p1_compiled.bin")
	if err != nil {
		t.Fatal(err)
	}
	p, err := Read(bin)
	if err != nil {
		t.Fatal(err)
	}
	// ProjectStreamRaw is byte-identical to the CFB PROJECT stream.
	c, err := cfb.Open(bin)
	if err != nil {
		t.Fatal(err)
	}
	projRaw, ok := c.Stream("PROJECT")
	if !ok {
		t.Fatal("missing PROJECT stream")
	}
	if !bytes.Equal(p.ProjectStreamRaw, projRaw) {
		t.Error("ProjectStreamRaw does not match the PROJECT stream")
	}
	// ProjectInfoRaw is a non-empty span that begins with SYSKIND(0x0001).
	if len(p.ProjectInfoRaw) < 2 {
		t.Fatalf("ProjectInfoRaw too short: %d", len(p.ProjectInfoRaw))
	}
	if id := binary.LittleEndian.Uint16(p.ProjectInfoRaw); id != 0x0001 {
		t.Errorf("ProjectInfoRaw leading id = 0x%04X, want 0x0001", id)
	}
}
