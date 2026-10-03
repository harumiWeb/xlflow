package oforms

import (
	"bytes"
	"errors"
	"strings"
	"testing"
)

func TestRewriteVBFrameCaptionEscapesQuotesAndPreservesOtherLines(t *testing.T) {
	base := openFixture(t, "p4_form.bin")
	form, err := ReadForm(base, "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	original := form.Levels[0].VBFrameRaw
	updated, err := rewriteVBFrameCaption(original, `Edited "root"`, 932)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Contains(updated, []byte(`Caption = "Edited ""root"""`)) {
		t.Fatalf("escaped caption missing: %q", updated)
	}
	for _, line := range [][]byte{[]byte("VERSION 5.00\r\n"), []byte("ClientHeight"), []byte("ClientWidth"), []byte("StartUpPosition")} {
		if !bytes.Contains(updated, line) {
			t.Fatalf("unrelated VBFrame line %q was lost", line)
		}
	}
	if bytes.Equal(original, updated) {
		t.Fatal("caption rewrite did not change VBFrame")
	}
	existing := []byte("VERSION 5.00\r\nBegin Form\r\n   Caption = \"old\"\r\nEnd\r\n")
	rewritten, err := rewriteVBFrameCaption(existing, "new", 932)
	if err != nil {
		t.Fatal(err)
	}
	if got := bytes.Count(rewritten, []byte("Caption =")); got != 1 || !bytes.Contains(rewritten, []byte(`Caption = "new"`)) {
		t.Fatalf("existing root caption count=%d bytes=%q", got, rewritten)
	}
}

func TestApplyEditsSameRootCaptionPreservesVBFrameBytes(t *testing.T) {
	base, err := ReadForm(openFixture(t, "p4_form.bin"), "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	withCaption, err := ApplyEdits(base, []Edit{{Control: "", Property: "Caption", Value: "Edited Root"}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	before, err := SerializeForm(withCaption, 932)
	if err != nil {
		t.Fatal(err)
	}
	repeated, err := ApplyEdits(withCaption, []Edit{{Control: "", Property: "Caption", Value: "Edited Root"}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	after, err := SerializeForm(repeated, 932)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(before.Streams["UserForm1/\x03VBFrame"], after.Streams["UserForm1/\x03VBFrame"]) {
		t.Fatal("same root caption rewrote VBFrame")
	}
	if !bytes.Equal(before.Streams["UserForm1/f"], after.Streams["UserForm1/f"]) {
		t.Fatal("same root caption reencoded f stream")
	}
}

func TestRewriteVBFrameCaptionPreservesPropertyBlocks(t *testing.T) {
	block := "   BeginProperty Opaque\r\n      Caption = \"opaque\"\r\n   EndProperty\r\n"
	before := []byte("VERSION 5.00\r\nBegin Form\r\n" + block + "End\r\n")
	after, err := rewriteVBFrameCaption(before, "root", 932)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Contains(after, []byte(block)) || bytes.Count(after, []byte("Caption = \"root\"")) != 1 {
		t.Fatalf("property block was mistaken for the root: %q", after)
	}
}

func TestRewriteVBFrameCaptionRejectsUnsupportedLayoutsAndValues(t *testing.T) {
	for _, tc := range []struct {
		name string
		raw  []byte
		text string
	}{
		{"duplicate-caption", []byte("VERSION 5.00\r\nBegin Form\r\n Caption = \"one\"\r\n Caption = \"two\"\r\nEnd\r\n"), "new"},
		{"unknown-layout", []byte("VERSION 5.00\r\nCaption = \"one\"\r\n"), "new"},
		{"crlf-caption", []byte("VERSION 5.00\r\nBegin Form\r\nEnd\r\n"), "bad\r\nline"},
		{"unrepresentable-caption", []byte("VERSION 5.00\r\nBegin Form\r\nEnd\r\n"), "😀"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			_, err := rewriteVBFrameCaption(tc.raw, tc.text, 932)
			if !errors.Is(err, ErrInvalidEdit) {
				t.Fatalf("error = %v, want ErrInvalidEdit", err)
			}
		})
	}
}

func TestApplyEditsSynchronizesRootCaptionVBFrame(t *testing.T) {
	base, err := ReadForm(openFixture(t, "p4_form.bin"), "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	got, err := ApplyEdits(base, []Edit{{Control: "", Property: "Caption", Value: "Edited Root"}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	if got.Levels[0].Record.Strings["Caption"].Text != "Edited Root" || !strings.Contains(got.DesignerSource.Text, `Caption = "Edited Root"`) {
		t.Fatalf("root caption was not dual-persisted: record=%q VBFrame=%q", got.Levels[0].Record.Strings["Caption"].Text, got.DesignerSource.Text)
	}
	if _, err := ApplyEdits(base, []Edit{{Control: "", Property: "Caption", Value: "😀"}}, 932); !errors.Is(err, ErrInvalidEdit) {
		t.Fatalf("unrepresentable root caption error = %v", err)
	}
}
