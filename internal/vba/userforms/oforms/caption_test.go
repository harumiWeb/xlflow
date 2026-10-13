package oforms

import (
	"bytes"
	"errors"
	"strings"
	"testing"
)

func TestRewriteVBFrameTypeInfoBindsOnlyRootVersion(t *testing.T) {
	raw := []byte("VERSION 5.00\r\nBegin Form\r\n   Caption = \"keep\"\r\n   TypeInfoVer = 7\r\n   BeginProperty Opaque\r\n      TypeInfoVer = 99\r\n   EndProperty\r\nEnd\r\n")
	got, err := rewriteVBFrameTypeInfo(raw, 9, 932)
	if err != nil {
		t.Fatal(err)
	}
	want := bytes.Replace(raw, []byte("   TypeInfoVer = 7\r\n"), []byte("   TypeInfoVer = 9\r\n"), 1)
	if !bytes.Equal(got, want) {
		t.Fatalf("unrelated bytes changed: %q", got)
	}
	unchanged, err := rewriteVBFrameTypeInfo(got, 9, 932)
	if err != nil || !bytes.Equal(unchanged, got) {
		t.Fatal("same version must preserve bytes", err)
	}
	absent := []byte("VERSION 5.00\r\nBegin Form\r\nEnd\r\n")
	got, err = rewriteVBFrameTypeInfo(absent, 9, 932)
	if err != nil || !bytes.Equal(got, absent) {
		t.Fatal("omitted version must remain omitted", err)
	}
	for _, value := range []string{"-1", "bad", "4294967296"} {
		invalid := bytes.Replace(raw, []byte("TypeInfoVer = 7"), []byte("TypeInfoVer = "+value), 1)
		if _, err := rewriteVBFrameTypeInfo(invalid, 9, 932); !errors.Is(err, ErrInvalidEdit) {
			t.Fatalf("invalid version %s: %v", value, err)
		}
	}
}

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

func TestRewriteVBFrameClientSizeAcceptsExcelFractionalClientWidth(t *testing.T) {
	raw := []byte("VERSION 5.00\r\nBegin Form\r\n   ClientHeight    =   8000\r\n   ClientLeft      =   0\r\n   ClientTop       =   0\r\n   ClientWidth     =   9996.001\r\n   StartUpPosition =   1  ' Excel normalized\r\nEnd\r\n")
	got, err := rewriteVBFrameClientSize(raw, Size{Width: 18344, Height: 14817}, 932)
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range [][]byte{
		[]byte("   ClientWidth     =   10400\r\n"),
		[]byte("   ClientHeight    =   8400\r\n"),
		[]byte("   ClientLeft      =   0\r\n"),
		[]byte("   ClientTop       =   0\r\n"),
		[]byte("   StartUpPosition =   1  ' Excel normalized\r\n"),
	} {
		if !bytes.Contains(got, want) {
			t.Fatalf("rewritten VBFrame is missing preserved or updated line %q: %q", want, got)
		}
	}
}

func TestRewriteVBFrameClientSizeRejectsInvalidNumericValues(t *testing.T) {
	for _, value := range []string{"malformed", "-1", "NaN", "Inf", "0x1p2", "1217471045", "2147483648"} {
		t.Run(value, func(t *testing.T) {
			raw := []byte("VERSION 5.00\r\nBegin Form\r\n   ClientHeight = 8000\r\n   ClientWidth = " + value + "\r\nEnd\r\n")
			_, err := rewriteVBFrameClientSize(raw, Size{Width: 18344, Height: 14817}, 932)
			dimensionErr, ok := errors.AsType[*vbFrameDimensionError](err)
			if !ok || dimensionErr.Property != "Width" || !errors.Is(err, ErrInvalidEdit) {
				t.Fatalf("invalid ClientWidth %q error = %v, want typed invalid Width", value, err)
			}
		})
	}
}

func TestRewriteVBFrameClientSizeAcceptsLargestTwipsValueWithinHimetricRange(t *testing.T) {
	raw := []byte("VERSION 5.00\r\nBegin Form\r\n   ClientHeight = 8000\r\n   ClientWidth = 1217471044\r\nEnd\r\n")

	if _, err := rewriteVBFrameClientSize(raw, Size{Width: 18344, Height: 14817}, 932); err != nil {
		t.Fatalf("largest HIMETRIC-representable ClientWidth was rejected: %v", err)
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

func TestApplyEditsRepeatedRootCaptionsUseLastValue(t *testing.T) {
	fixture, err := ReadForm(openFixture(t, "p4_form.bin"), "UserForm1", 932)
	if err != nil {
		t.Fatal(err)
	}
	base, err := ApplyEdits(fixture, []Edit{{Property: "Caption", Value: "Original"}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	for _, final := range []string{"Original", "Changed", "Second"} {
		t.Run(final, func(t *testing.T) {
			got, err := ApplyEdits(base, []Edit{
				{Property: "Caption", Value: "Changed"},
				{Property: "Caption", Value: final},
			}, 932)
			if err != nil {
				t.Fatal(err)
			}
			if caption := got.Levels[0].Record.Strings["Caption"].Text; caption != final {
				t.Fatalf("record Caption = %q, want %q", caption, final)
			}
			caption, found, err := rootVBFrameCaption(got.Levels[0].VBFrameRaw, 932)
			if err != nil || !found || caption != final {
				t.Fatalf("VBFrame Caption = %q, found=%v, error=%v; want %q", caption, found, err, final)
			}
			if caption := base.Levels[0].Record.Strings["Caption"].Text; caption != "Original" {
				t.Fatalf("input Caption was mutated: %q", caption)
			}
		})
	}
}
