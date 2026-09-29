package ovba

import (
	"bytes"
	"strings"
	"testing"
)

func TestRebuildProjectTextChangesOnlySourceOwnedComponents(t *testing.T) {
	raw := []byte("ID=\"{ABC}\"\r\n" +
		"Document=ThisWorkbook/&H00000000\r\n" +
		"Module=OldModule\r\n" +
		"Class=KeepClass\r\n" +
		"BaseClass=UserForm1\r\n" +
		"Name=\"VBAProject\"\r\n" +
		"HelpFile=\"C:\\\\docs\\\\help.chm\"\r\n")
	specs := []ProjectComponentSpec{
		{Kind: "Document", Name: "ThisWorkbook"},
		{Kind: "Class", Name: "KeepClass"},
		{Kind: "BaseClass", Name: "UserForm1"},
		{Kind: "Module", Name: "NewModule"},
	}

	got, err := RebuildProjectText(raw, specs)
	if err != nil {
		t.Fatal(err)
	}
	want := "ID=\"{ABC}\"\r\n" +
		"Document=ThisWorkbook/&H00000000\r\n" +
		"Class=KeepClass\r\n" +
		"BaseClass=UserForm1\r\n" +
		"Module=NewModule\r\n" +
		"Name=\"VBAProject\"\r\n" +
		"HelpFile=\"C:\\\\docs\\\\help.chm\"\r\n"
	if string(got) != want {
		t.Fatalf("PROJECT stream mismatch\ngot:\n%q\nwant:\n%q", got, want)
	}
}

func TestRebuildProjectTextLeavesUnchangedProjectByteExact(t *testing.T) {
	raw := []byte("ID=\"{ABC}\"\r\nDocument=Sheet1/&H00000000\r\nModule=Module1\r\nName=\"VBAProject\"\r\n")
	got, err := RebuildProjectText(raw, []ProjectComponentSpec{
		{Kind: "Document", Name: "Sheet1"},
		{Kind: "Module", Name: "Module1"},
	})
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(got, raw) {
		t.Fatalf("unchanged PROJECT drifted: got %q want %q", got, raw)
	}
}

func TestRebuildProjectTextRejectsTemplateOwnedTopologyChanges(t *testing.T) {
	raw := []byte("Document=Sheet1/&H00000000\r\nBaseClass=UserForm1\r\n")
	cases := []struct {
		name  string
		specs []ProjectComponentSpec
	}{
		{name: "remove document", specs: []ProjectComponentSpec{{Kind: "BaseClass", Name: "UserForm1"}}},
		{name: "rename document", specs: []ProjectComponentSpec{{Kind: "Document", Name: "Sheet2"}, {Kind: "BaseClass", Name: "UserForm1"}}},
		{name: "add form", specs: []ProjectComponentSpec{{Kind: "Document", Name: "Sheet1"}, {Kind: "BaseClass", Name: "UserForm1"}, {Kind: "BaseClass", Name: "UserForm2"}}},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			_, err := RebuildProjectText(raw, tc.specs)
			if err == nil || !strings.Contains(err.Error(), "template-owned") {
				t.Fatalf("error = %v, want template-owned topology error", err)
			}
		})
	}
}

func TestRebuildProjectTextReportsTemplateOwnedRemovalInDeclarationOrder(t *testing.T) {
	raw := []byte("Document=Sheet2/&H00000000\r\nDocument=Sheet1/&H00000000\r\n")
	_, err := RebuildProjectText(raw, nil)
	want := "template-owned PROJECT component Document=Sheet2 cannot be added, removed, or renamed"
	if err == nil || err.Error() != want {
		t.Fatalf("error = %v, want %q", err, want)
	}
}

func TestRebuildProjectTextReportsTemplateOwnedAdditionInSpecOrder(t *testing.T) {
	raw := []byte("Document=Sheet1/&H00000000\r\n")
	_, err := RebuildProjectText(raw, []ProjectComponentSpec{
		{Kind: "Document", Name: "Sheet1"},
		{Kind: "BaseClass", Name: "UserForm2"},
		{Kind: "BaseClass", Name: "UserForm1"},
	})
	want := "template-owned PROJECT component BaseClass=UserForm2 is not present in the template"
	if err == nil || err.Error() != want {
		t.Fatalf("error = %v, want %q", err, want)
	}
}

func TestRebuildProjectTextRejectsCaseInsensitiveDuplicates(t *testing.T) {
	_, err := RebuildProjectText([]byte("Module=Module1\r\n"), []ProjectComponentSpec{
		{Kind: "Module", Name: "Module1"},
		{Kind: "Class", Name: "module1"},
	})
	if err == nil || !strings.Contains(err.Error(), "duplicate") {
		t.Fatalf("error = %v, want duplicate component error", err)
	}
}

func TestRebuildProjectTextReplacesFinalDeclarationWithoutLeadingBlankLine(t *testing.T) {
	got, err := RebuildProjectText([]byte("Module=Old"), []ProjectComponentSpec{{Kind: "Module", Name: "New"}})
	if err != nil {
		t.Fatal(err)
	}
	if string(got) != "Module=New\r\n" {
		t.Fatalf("PROJECT stream = %q, want replacement without a leading blank line", got)
	}
}
