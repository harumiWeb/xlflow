package oforms

import (
	"bytes"
	"encoding/binary"
	"errors"
	"reflect"
	"testing"
)

func TestTextFontPartialEditRetainsUneditedFields(t *testing.T) {
	r := newGenerationRecord(&textPropsSpec)
	for key, value := range map[string]int64{"FontEffects": 0x4000200b, "FontHeight": 165, "FontWeight": 650, "FontPitchAndFamily": 34, "ParagraphAlign": 2} {
		if err := setGenerationProperty(r, &textPropsSpec, key, value); err != nil {
			t.Fatal(err)
		}
	}
	if err := setGenerationProperty(r, &textPropsSpec, "FontName", "Tahoma"); err != nil {
		t.Fatal(err)
	}
	if err := editTextFont(r, Font{Italic: new(false), Size: new(12.5)}, 932); err != nil {
		t.Fatal(err)
	}
	read, _, err := parseRecord(r.Raw, &textPropsSpec, 932, len(r.Raw))
	if err != nil {
		t.Fatal(err)
	}
	if read.Values["FontEffects"] != 0x40002009 || read.Values["FontHeight"] != 250 || read.Values["FontWeight"] != 650 || read.Values["FontPitchAndFamily"] != 34 || read.Values["ParagraphAlign"] != 2 || read.Strings["FontName"].Text != "Tahoma" {
		t.Fatalf("partial edit lost state: %+v", read)
	}
	if ReadTextFont(newGenerationRecord(&textPropsSpec)) != nil {
		t.Fatal("omitted font became materialized defaults")
	}
}

func TestStdFontPartialEditPreservesRawStyleAndWeight(t *testing.T) {
	raw := make([]byte, 27)
	binary.LittleEndian.PutUint32(raw, guidStdFont)
	raw[16] = 1
	binary.LittleEndian.PutUint16(raw[17:], 128)
	raw[19] = 0x88
	binary.LittleEndian.PutUint16(raw[20:], 650)
	binary.LittleEndian.PutUint32(raw[22:], 82500)
	raw[26] = 6
	raw = append(raw, []byte("Tahoma")...)
	level := &Level{Record: newGenerationRecord(&formSpec), FontRaw: bytes.Clone(raw)}
	if err := editLevelFont(level, Font{Underline: new(true)}, 932); err != nil {
		t.Fatal(err)
	}
	want := bytes.Clone(raw)
	want[19] |= 4
	if !bytes.Equal(want, level.FontRaw) {
		t.Fatal("partial StdFont edit replaced unrelated bytes")
	}
	f, err := ReadLevelFont(level)
	if err != nil || *f.Size != 8.25 || *f.Charset != 128 || !*f.Underline || !*f.StrikeThrough || !*f.Bold {
		t.Fatalf("wrong StdFont projection: %+v %v", f, err)
	}
	if err := editLevelFont(level, Font{Name: new("日本語")}, 932); !errors.Is(err, ErrUnsupportedEdit) {
		t.Fatalf("unsafe StdFont conversion accepted: %v", err)
	}
}

func TestFontEditsRoundTripAllStorageLayouts(t *testing.T) {
	form, err := NewForm(Definition{Name: "Fonts", Size: Size{5000, 4000}, Controls: []ControlDefinition{
		{Name: "LabelOne", Class: 21, Visible: true, Size: Size{1000, 500}},
		{Name: "FrameOne", Class: 14, Visible: true, Size: Size{2000, 1000}},
		{Name: "Pages", Class: 57, Visible: true, Size: Size{2000, 1000}, Tabs: &TabStrip{Tabs: []Tab{{Name: "TabOne", Caption: "PageOne", Enabled: true, Visible: true}}}, Controls: []ControlDefinition{{Name: "PageOne", Class: 7, Visible: true}}},
	}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	want := Font{Name: new("ＭＳ ゴシック"), Size: new(12.5), Bold: new(true), Italic: new(false), Charset: new(128)}
	var edits []Edit
	for _, name := range []string{"", "LabelOne", "FrameOne", "Pages", "PageOne"} {
		edits = append(edits, Edit{Control: name, Property: "Font", Value: want})
	}
	result, err := ApplyEdits(form, edits, 932)
	if err != nil {
		t.Fatal(err)
	}
	got, err := ReadLevelFont(result.Levels[0])
	if err != nil || got == nil || *got.Name != *want.Name {
		t.Fatalf("root font missing: %v %v", got, err)
	}
	for _, level := range result.Levels {
		for _, control := range level.Controls {
			if control.Name == "" {
				continue
			}
			got, err := ReadControlFont(control)
			if err != nil || got == nil || *got.Name != *want.Name || *got.Size != 12.5 || !*got.Bold || *got.Italic || *got.Charset != 128 {
				t.Fatalf("%s font differs: %+v %v", control.Name, got, err)
			}
		}
	}
	a, _ := SerializeForm(result, 932)
	noop, err := ApplyEdits(result, nil, 932)
	if err != nil {
		t.Fatal(err)
	}
	b, _ := SerializeForm(noop, 932)
	if !reflect.DeepEqual(a, b) {
		t.Fatal("no-op changes font streams")
	}
	if len(form.Levels[0].FontRaw) != 0 {
		t.Fatal("edit mutated input")
	}
}

func TestFontUnsafeValuesAndUnknownMaskFail(t *testing.T) {
	for _, font := range []Font{{Size: new(0.0)}, {Size: new(12.333)}, {Charset: new(256)}, {Name: new("")}} {
		r := newGenerationRecord(&textPropsSpec)
		if err := editTextFont(r, font, 932); err == nil {
			t.Fatalf("accepted %+v", font)
		}
	}
	r := newGenerationRecord(&textPropsSpec)
	r.Mask = 1 << 30
	if err := editTextFont(r, Font{Bold: new(true)}, 932); !errors.Is(err, ErrUnsupportedEdit) {
		t.Fatalf("unknown mask replaced: %v", err)
	}
}

// Real Excel regression: the generated Label loaded Size=12.5 but Bold=False
// when FontEffects bit0 was set and FontWeight omitted. Bold must explicitly
// persist the weight; this test alone is not an Excel oracle.
func TestTextFontBoldPersistsWeightAndProjectsWeightPrecedence(t *testing.T) {
	legacy := newGenerationRecord(&textPropsSpec)
	if err := setGenerationProperty(legacy, &textPropsSpec, "FontEffects", int64(1)); err != nil {
		t.Fatal(err)
	}
	if font := ReadTextFont(legacy); font.Bold == nil || *font.Bold {
		t.Fatal("omitted weight400 incorrectly projected Bold=true")
	}
	for _, bold := range []bool{true, false} {
		r := newGenerationRecord(&textPropsSpec)
		if err := editTextFont(r, Font{Bold: new(bold)}, 932); err != nil {
			t.Fatal(err)
		}
		read, _, err := parseRecord(r.Raw, &textPropsSpec, 932, len(r.Raw))
		if err != nil {
			t.Fatal(err)
		}
		weight := int64(400)
		if bold {
			weight = 700
		}
		if read.Mask&(1<<7) == 0 || read.Values["FontWeight"] != weight {
			t.Fatalf("bold=%v missing weight %d", bold, weight)
		}
		// Deliberately contradictory effect bit: persisted weight wins.
		read.Values["FontEffects"] ^= 1
		if got := ReadTextFont(read); got.Bold == nil || *got.Bold != bold {
			t.Fatalf("weight precedence failed: %+v", got)
		}
		if err := editTextFont(read, Font{Italic: new(true)}, 932); err != nil {
			t.Fatal(err)
		}
		if read.Values["FontWeight"] != weight {
			t.Fatal("unrelated edit changed weight")
		}
	}
	for _, weight := range []int64{0, 400, 550, 551, 600, 650, 700, 1000} {
		r := newGenerationRecord(&textPropsSpec)
		if err := setGenerationProperty(r, &textPropsSpec, "FontWeight", weight); err != nil {
			t.Fatal(err)
		}
		if f := ReadTextFont(r); f == nil || f.Bold == nil || *f.Bold != (weight > 550) {
			t.Fatalf("weight %d bold differs: %+v", weight, f)
		}
		if err := editTextFont(r, Font{Italic: new(true)}, 932); err != nil {
			t.Fatal(err)
		}
		if r.Values["FontWeight"] != weight {
			t.Fatalf("partial edit changed weight%d", weight)
		}
	}
}
