package vbaproject

import (
	"bytes"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
)

// PROJECTwm is a root-level stream that lives outside the VBA storage and the
// PROJECT stream. The minimal-profile writer used to drop every stream it did
// not regenerate; the structural-boundary pass-through must instead carry any
// stream whose first path segment is neither "VBA" nor "PROJECT" through a
// Read -> Write round-trip byte-for-byte. PROJECTwm is the witness on a
// non-form fixture (the same mechanism later carries UserForm designer storages).
func TestWritePreservesRootStreamsOutsideVBA(t *testing.T) {
	in := loadBin(t, "p2_refs.bin")

	orig, err := cfb.Open(in)
	if err != nil {
		t.Fatal(err)
	}
	want, ok := orig.Stream("PROJECTwm")
	if !ok {
		t.Skip("fixture has no PROJECTwm; nothing to assert")
	}

	p, err := Read(in)
	if err != nil {
		t.Fatal(err)
	}
	out, err := Write(p)
	if err != nil {
		t.Fatalf("write: %v", err)
	}

	got, err := cfb.Open(out)
	if err != nil {
		t.Fatal(err)
	}
	have, ok := got.Stream("PROJECTwm")
	if !ok {
		t.Fatal("PROJECTwm was dropped: root streams outside VBA/ must pass through")
	}
	if !bytes.Equal(have, want) {
		t.Errorf("PROJECTwm not preserved verbatim: %d bytes -> %d bytes", len(want), len(have))
	}
}

// A UserForm contributes a code-behind module stream under VBA/ and a separate
// root-level designer storage (UserForm1/{f, o, \x01CompObj, \x03VBFrame}) that
// holds the form definition. Writing a form-bearing project must carry that
// designer storage through byte-for-byte; the writer never models or edits it.
func TestWriteFormPreservesDesignerStorage(t *testing.T) {
	in := loadBin(t, "p4_form.bin")

	orig, err := cfb.Open(in)
	if err != nil {
		t.Fatal(err)
	}
	p, err := Read(in)
	if err != nil {
		t.Fatal(err)
	}
	out, err := Write(p)
	if err != nil {
		t.Fatalf("write: %v", err)
	}

	got, err := cfb.Open(out)
	if err != nil {
		t.Fatal(err)
	}
	for _, child := range []string{"f", "o", "\x01CompObj", "\x03VBFrame"} {
		path := "UserForm1/" + child
		want, ok := orig.Stream(path)
		if !ok {
			t.Fatalf("precondition: fixture lacks %q", path)
		}
		have, ok := got.Stream(path)
		if !ok {
			t.Errorf("designer stream %q dropped", path)
			continue
		}
		if !bytes.Equal(have, want) {
			t.Errorf("designer stream %q not preserved verbatim (%d -> %d bytes)", path, len(want), len(have))
		}
	}
}

// Editing a UserForm's code-behind must persist while its sibling designer
// storage stays byte-for-byte identical. The two live in different CFB entries
// (the VBA/UserForm1 module stream vs the root-level UserForm1/ storage), so a
// source edit must not couple to the form layout.
func TestWriteFormEditCodeBehindKeepsDesigner(t *testing.T) {
	in := loadBin(t, "p4_form.bin")
	orig, err := cfb.Open(in)
	if err != nil {
		t.Fatal(err)
	}

	p, err := Read(in)
	if err != nil {
		t.Fatal(err)
	}
	formIndex := -1
	for i := range p.Modules {
		if p.Modules[i].Type == ModuleForm {
			formIndex = i
		}
	}
	if formIndex < 0 {
		t.Fatal("precondition: fixture has no form module")
	}
	form := &p.Modules[formIndex]
	edited := form.Source + "\r\n' edited\r\n"
	form.Source = edited
	wantName := form.Name

	out, err := Write(p)
	if err != nil {
		t.Fatalf("write: %v", err)
	}

	p2, err := Read(out)
	if err != nil {
		t.Fatalf("re-read: %v", err)
	}
	var got string
	for i := range p2.Modules {
		if p2.Modules[i].Name == wantName {
			got = p2.Modules[i].Source
		}
	}
	if got != edited {
		t.Errorf("edited code-behind did not round-trip")
	}

	gotCFB, err := cfb.Open(out)
	if err != nil {
		t.Fatal(err)
	}
	for _, child := range []string{"f", "o", "\x01CompObj", "\x03VBFrame"} {
		path := "UserForm1/" + child
		want, _ := orig.Stream(path)
		have, ok := gotCFB.Stream(path)
		if !ok || !bytes.Equal(have, want) {
			t.Errorf("designer stream %q changed by a code-behind edit", path)
		}
	}
}

// A UserForm with container controls (Frame, MultiPage, Page) nests their data
// in sub-storages of the designer storage, several levels deep (e.g.
// UserForm1/i05/i07/o). Every stream in that subtree must round-trip
// byte-for-byte: the writer rebuilds the nested storages from their stream paths
// and never models the controls. The control class survives in each storage's
// \x01CompObj ProgID even though the writer zeroes the storage's directory CLSID.
func TestWriteNestedFormPreservesDesignerStorages(t *testing.T) {
	in := loadBin(t, "p6_nested_form.bin")
	orig, err := cfb.Open(in)
	if err != nil {
		t.Fatal(err)
	}
	p, err := Read(in)
	if err != nil {
		t.Fatal(err)
	}
	out, err := Write(p)
	if err != nil {
		t.Fatalf("write: %v", err)
	}
	got, err := cfb.Open(out)
	if err != nil {
		t.Fatal(err)
	}

	checked, deepSeen := 0, false
	for _, path := range orig.Paths() {
		if !strings.HasPrefix(path, "UserForm1/") {
			continue
		}
		want, _ := orig.Stream(path)
		have, ok := got.Stream(path)
		if !ok {
			t.Errorf("designer stream %q dropped", path)
			continue
		}
		if !bytes.Equal(have, want) {
			t.Errorf("designer stream %q not preserved verbatim", path)
		}
		checked++
		if strings.Count(path, "/") >= 3 { // e.g. UserForm1/i05/i07/o
			deepSeen = true
		}
	}
	if checked == 0 {
		t.Fatal("precondition: fixture has no UserForm1 designer streams")
	}
	if !deepSeen {
		t.Fatal("precondition: fixture has no >=3-level nested designer storage")
	}
}
