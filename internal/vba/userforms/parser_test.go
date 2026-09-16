package userforms

import (
	"encoding/binary"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
)

func TestParseExtractsNestedMSFormsControls(t *testing.T) {
	source := `VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerForm
   Caption = "Customer"
   Begin MSForms.TextBox txtName
      Height = 18
   End
   Begin Forms.CommandButton.1 cmdOK
      Caption = "OK"
   End
End
Attribute VB_Name = "CustomerForm"
`
	form := Parse(source)
	if !form.Complete {
		t.Fatal("valid UserForm designer was marked incomplete")
	}
	if form.Name != "CustomerForm" {
		t.Fatalf("form name = %q, want CustomerForm", form.Name)
	}
	if len(form.Controls) != 2 {
		t.Fatalf("controls = %d, want 2: %+v", len(form.Controls), form.Controls)
	}
	if form.Controls[0].Name != "txtName" || form.Controls[0].Type != "MSForms.TextBox" {
		t.Fatalf("unexpected first control: %+v", form.Controls[0])
	}
	if form.Controls[1].Name != "cmdOK" || form.Controls[1].Type != "MSForms.CommandButton" {
		t.Fatalf("unexpected second control: %+v", form.Controls[1])
	}
}

func TestParseMarksMalformedDesignerIncomplete(t *testing.T) {
	form := Parse(`VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerForm
   Begin MSForms.TextBox
End
`)
	if form.Complete {
		t.Fatalf("malformed UserForm designer was marked complete: %+v", form)
	}
}

func TestParseMarksTabSeparatedMalformedBeginIncomplete(t *testing.T) {
	form := Parse("VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerForm\nBegin\tMSForms.TextBox\nEnd\n")
	if form.Complete {
		t.Fatalf("tab-separated malformed Begin was marked complete: %+v", form)
	}
}

func TestParseMarksMismatchedAttributeNameIncomplete(t *testing.T) {
	form := Parse(`VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerForm
End
Attribute VB_Name = "StaleForm"
`)
	if form.Complete {
		t.Fatalf("mismatched Attribute VB_Name was marked complete: %+v", form)
	}
}

func TestParseAllowsVBAEndAfterDesigner(t *testing.T) {
	form := Parse(`VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerForm
End
Attribute VB_Name = "CustomerForm"
End
`)
	if !form.Complete {
		t.Fatalf("a VBA End statement after the designer was marked incomplete: %+v", form)
	}
}

func TestParseRejectsExtraEndBeforeVBASection(t *testing.T) {
	form := Parse(`VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerForm
End
End
Attribute VB_Name = "CustomerForm"
`)
	if form.Complete {
		t.Fatalf("extra End before Attribute VB_Name was marked complete: %+v", form)
	}
}

func TestFindFRXControlNamesMatchesControlCandidates(t *testing.T) {
	writer := cfb.NewWriter()
	stream := appendFRXControlRecord(nil, "ButtonRun")
	stream = append(stream, []byte("mat\x00")...)
	stream = append(stream, appendFRXControlRecord(nil, "ComboBoxBitmapVector")...)
	writer.AddStream([]string{"f"}, stream)
	data, err := writer.Bytes()
	if err != nil {
		t.Fatal(err)
	}

	helperEventName := "Apply" + "_BitmapVector" + "_Change"
	got, ok := FindFRXControlNames(data, []string{"ButtonRun", "ComboBox", "ComboBoxBitmapVector", helperEventName})
	if !ok {
		t.Fatal("synthetic FRX was not parsed")
	}
	if len(got) != 2 {
		t.Fatalf("control names = %#v, want two matched controls", got)
	}
	for _, name := range []string{"buttonrun", "comboboxbitmapvector"} {
		if _, ok := got[name]; !ok {
			t.Fatalf("control name %q was not found in %#v", name, got)
		}
	}
	if _, ok := got[strings.ToLower(helperEventName)]; ok {
		t.Fatalf("helper procedure was mistaken for a control: %#v", got)
	}
}

func appendFRXControlRecord(dst []byte, name string) []byte {
	record := make([]byte, frxRecordNameOffset+len(name))
	record[0] = 0xe5
	binary.LittleEndian.PutUint32(record[frxRecordNameLengthOffset:], uint32(len(name))|0x80000000)
	copy(record[frxRecordNameOffset:], name)
	return append(dst, record...)
}

func TestFindFRXControlNamesReadsIguanaTexArtifacts(t *testing.T) {
	path := filepath.Join("..", "..", "..", "testdata", "static-analysis-corpus", "projects", "third_party", "iguana-tex", "BatchEditForm.frx")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	helperName := "Apply" + "_BitmapVector"
	got, ok := FindFRXControlNames(data, []string{"ButtonRun", "ComboBoxBitmapVector", helperName})
	if !ok {
		t.Fatal("real FRX was not parsed")
	}
	for _, name := range []string{"buttonrun", "comboboxbitmapvector"} {
		if _, ok := got[name]; !ok {
			t.Fatalf("real FRX control name %q was not found in %#v", name, got)
		}
	}
	if _, ok := got[strings.ToLower(helperName)]; ok {
		t.Fatalf("real FRX helper name was mistaken for a control: %#v", got)
	}
}

func TestFindFRXControlNamesReadsNestedIguanaTexArtifacts(t *testing.T) {
	path := filepath.Join("..", "..", "..", "testdata", "static-analysis-corpus", "projects", "third_party", "iguana-tex", "LatexForm.frx")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	got, ok := FindFRXControlNames(data, []string{"CmdButtonExternalEditor", "ToggleButtonWrap", "MultiPage1"})
	if !ok {
		t.Fatal("real nested FRX was not parsed")
	}
	for _, name := range []string{"cmdbuttonexternaleditor", "togglebuttonwrap", "multipage1"} {
		if _, ok := got[name]; !ok {
			t.Fatalf("nested FRX control name %q was not found in %#v", name, got)
		}
	}
}
