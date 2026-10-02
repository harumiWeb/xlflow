package oforms

import (
	"errors"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
)

func TestNewFormAuthorsIndependentValidatedStreams(t *testing.T) {
	input := Definition{Name: "FreshForm", Caption: "日本語😀", Size: Size{8467, 6350}, Controls: []ControlDefinition{{Name: "TextMain", Class: 23, Size: Size{4233, 635}, Position: Position{100, 200}, TabIndex: 4, Visible: true, Properties: map[string]any{"Value": "日本語😀", "Tag": "tag", "VariousPropertyBits": int64(0x2c804819)}}}}
	form, err := NewForm(input, 1252)
	if err != nil {
		t.Fatal(err)
	}
	stored, err := SerializeForm(form, 1252)
	if err != nil {
		t.Fatal(err)
	}
	if len(stored.Streams) != 4 || len(stored.Storages) != 1 || stored.Storages["FreshForm"] != (cfb.StorageMeta{}) {
		t.Fatal("generated stream/storage inventory differs")
	}
	if form.Controls[0].Record.Strings["Value"].Compressed {
		t.Fatal("Unicode must use uncompressed UTF16 regardless of code page")
	}
	if form.Controls[0].TabIndex == nil || *form.Controls[0].TabIndex != 4 || form.Controls[0].Site.Strings["Tag"].Text != "tag" {
		t.Fatal("site state differs")
	}
	input.Controls[0].Properties["Value"] = "mutated input"
	if form.Controls[0].Record.Strings["Value"].Text != "日本語😀" {
		t.Fatal("returned model aliases caller input")
	}
	form.Controls[0].Record.Strings["Value"] = StoredString{Text: "arbitrary change"}
	if _, err := SerializeForm(form, 1252); !errors.Is(err, ErrUnsupportedMutation) {
		t.Fatalf("generated signature is not enforced: %v", err)
	}
}

func TestNewFormRejectsUnsupportedLayoutsAndInvalidRecords(t *testing.T) {
	for _, tc := range []struct {
		name   string
		change func(*Definition)
		want   error
	}{
		{"reserved-name", func(d *Definition) { d.Name = "VBA" }, ErrInvalidEdit},
		{"path-name", func(d *Definition) { d.Name = "Foo/Bar" }, ErrInvalidEdit},
		{"container", func(d *Definition) { d.Controls[0].Class = 14 }, ErrUnsupportedEdit},
		{"tabstrip", func(d *Definition) { d.Controls[0].Class = 18 }, ErrUnsupportedEdit},
		{"picture", func(d *Definition) { d.Controls[0].Properties = map[string]any{"Picture": int64(0xffff)} }, ErrUnsupportedEdit},
		{"bad-field-width", func(d *Definition) { d.Controls[0].Properties = map[string]any{"BorderStyle": int64(1 << 20)} }, ErrInvalidEdit},
		{"duplicate-name", func(d *Definition) { d.Controls = append(d.Controls, d.Controls[0]) }, ErrInvalidEdit},
		{"negative-tab", func(d *Definition) { d.Controls[0].TabIndex = -1 }, ErrInvalidEdit},
	} {
		t.Run(tc.name, func(t *testing.T) {
			d := Definition{Name: "FreshForm", Size: Size{100, 100}, Controls: []ControlDefinition{{Name: "LabelMain", Class: 21, Size: Size{100, 100}, Visible: true}}}
			tc.change(&d)
			form, err := NewForm(d, 1252)
			if form != nil || !errors.Is(err, tc.want) {
				t.Fatalf("got %v %v, want %v", form, err, tc.want)
			}
		})
	}
}
