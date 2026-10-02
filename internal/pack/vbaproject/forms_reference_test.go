package vbaproject

import (
	"bytes"
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/ovba"
)

func TestEnsureMSFormsReferencePreservesAndAddsOnce(t *testing.T) {
	for _, book := range []string{"p2_refs", "p4_form", "p6_nested_form"} {
		t.Run(book, func(t *testing.T) {
			p, err := Read(loadCorpus(t, book))
			if err != nil {
				t.Fatal(err)
			}
			before := cloneProject(p)
			result, err := EnsureMSFormsReference(p)
			if err != nil {
				t.Fatal(err)
			}
			if !reflect.DeepEqual(p, before) {
				t.Fatal("input changed")
			}
			if !bytes.HasPrefix(result.ReferencesRaw, p.ReferencesRaw) {
				t.Fatal("unrelated references changed")
			}
			again, err := EnsureMSFormsReference(result)
			if err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(result.ReferencesRaw, again.ReferencesRaw) {
				t.Fatal("duplicate Forms reference")
			}
			if book != "p2_refs" && !bytes.Equal(p.ReferencesRaw, result.ReferencesRaw) {
				t.Fatal("existing CONTROL reference changed")
			}
			body, err := Write(result)
			if err != nil {
				t.Fatal(err)
			}
			readback, err := Read(body)
			if err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(result.ReferencesRaw, readback.ReferencesRaw) {
				t.Fatal("reference round trip drift")
			}
		})
	}
}

func TestExcelFirstFormsReferenceEvidence(t *testing.T) {
	for _, stage := range []struct {
		name  string
		forms int
		kind  uint16
	}{
		{"form-free", 0, 0}, {"first-form", 1, 0x002F}, {"two-forms", 2, 0x002F}, {"blank-generated", 2, 0x000D},
	} {
		t.Run(stage.name, func(t *testing.T) {
			body, err := os.ReadFile(filepath.Join("testdata", "forms-reference-excel", stage.name+".bin"))
			if err != nil {
				t.Fatal(err)
			}
			p, err := Read(body)
			if err != nil {
				t.Fatal(err)
			}
			if len(p.Forms) != stage.forms {
				t.Fatalf("forms=%d", len(p.Forms))
			}
			refs, err := ovba.ParseProjectReferences(p.ReferencesRaw, p.Props.CodePage)
			if err != nil {
				t.Fatal(err)
			}
			formsCount := 0
			for _, ref := range refs {
				if ref.Kind == 0x000D && hasMSFormsLIBID([]byte(ref.LibID)) || ref.Kind == 0x002F && ref.OriginalTypeLib == userFormMSFormsReferenceGUIDWire {
					formsCount++
					if ref.Kind != stage.kind {
						t.Fatalf("Forms kind=%x", ref.Kind)
					}
				}
			}
			want := 0
			if stage.forms > 0 {
				want = 1
			}
			if formsCount != want {
				t.Fatalf("Forms references=%d", formsCount)
			}
			with, err := EnsureMSFormsReference(p)
			if err != nil {
				t.Fatal(err)
			}
			if stage.forms > 0 && !bytes.Equal(p.ReferencesRaw, with.ReferencesRaw) {
				t.Fatal("Excel Forms reference changed")
			}
		})
	}
}

func TestEnsureMSFormsReferenceBlankCodePages(t *testing.T) {
	for _, cp := range []uint16{874, 932, 936, 949, 950, 1250, 1251, 1252, 1253, 1254, 1255, 1256, 1257, 1258, 65001} {
		p, err := NewProject(NewProjectSpec{Name: "Test", CodePage: cp})
		if err != nil {
			t.Fatal(err)
		}
		r, err := EnsureMSFormsReference(p)
		if err != nil {
			t.Fatalf("cp%d: %v", cp, err)
		}
		refs, err := ovba.ParseProjectReferences(r.ReferencesRaw, cp)
		if err != nil || len(refs) != 3 {
			t.Fatalf("cp%d refs %v: %v", cp, refs, err)
		}
		if refs[2].Name != "MSForms" || refs[2].LibID != formsLibID {
			t.Fatal("wrong Forms identity")
		}
	}
}
