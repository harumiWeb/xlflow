package oforms

import (
	"errors"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/pack/cfb"
)

func TestNewFormAuthorsIndependentValidatedStreams(t *testing.T) {
	input := Definition{Name: "FreshForm", Caption: "Fresh Form", Size: Size{8467, 6350}, Controls: []ControlDefinition{{Name: "TextMain", Class: 23, Size: Size{4233, 635}, Position: Position{100, 200}, TabIndex: 4, Visible: true, Properties: map[string]any{"Value": "日本語😀", "Tag": "tag", "VariousPropertyBits": int64(0x2c804819)}}}}
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
	if !strings.Contains(form.DesignerSource.Text, `Caption = "Fresh Form"`) {
		t.Fatalf("VBFrame root caption missing: %q", form.DesignerSource.Text)
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

func TestNewFormAuthorsTabStripAndMultiPage(t *testing.T) {
	form, err := NewForm(Definition{Name: "Tabbed", Size: Size{8467, 6350}, Controls: []ControlDefinition{
		{
			Name: "Tabs", Class: 18, Size: Size{3000, 400}, Visible: true,
			Tabs: &TabStrip{SelectedIndex: 1, Tabs: []Tab{
				{Name: "TabOne", Caption: "一", Enabled: true, Visible: true},
				{Name: "TabTwo", Caption: "二", Enabled: true, Visible: true},
			}},
		},
		{
			Name: "Pages", Class: 57, Size: Size{5000, 3000}, Visible: true,
			Tabs: &TabStrip{SelectedIndex: 1, Tabs: []Tab{
				{Name: "TabPageA", Caption: "Alpha", Enabled: true, Visible: true},
				{Name: "TabPageB", Caption: "日本語", Enabled: false, Visible: true},
			}},
			Controls: []ControlDefinition{
				{Name: "PageA", Class: 7, Visible: true},
				{Name: "PageB", Class: 7, Visible: true},
			},
		},
	}}, 932)
	if err != nil {
		t.Fatal(err)
	}
	if len(form.Controls) != 2 || form.Controls[0].TabStrip == nil || form.Controls[1].MultiPage == nil {
		t.Fatal("generated TabStrip or MultiPage state was not bound")
	}
	standalone := form.Controls[0].TabStrip
	if len(standalone.Tabs) != 2 || standalone.SelectedIndex != 1 || standalone.Tabs[1].Caption != "二" {
		t.Fatalf("standalone tabs = %#v", standalone)
	}
	multi := form.Controls[1].MultiPage
	if len(multi.Pages) != 2 || multi.Pages[0].Name != "PageA" || multi.Pages[1].Name != "PageB" {
		t.Fatalf("generated pages = %#v", multi.Pages)
	}
	if multi.Hidden.Name != "" || multi.Hidden.TabStrip == nil || multi.Hidden.TabStrip.SelectedIndex != 1 || multi.Hidden.TabStrip.Tabs[1].Caption != "日本語" {
		t.Fatalf("generated hidden strip = %#v", multi.Hidden)
	}
	stored, err := SerializeForm(form, 932)
	if err != nil {
		t.Fatal(err)
	}
	reopened, err := reparseEditedForm(stored, "Tabbed", 932)
	if err != nil {
		t.Fatal(err)
	}
	if reopened.Controls[1].MultiPage.Hidden.TabStrip.Tabs[1].Enabled || reopened.Controls[1].MultiPage.Pages[1].Name != "PageB" {
		t.Fatal("generated tab flags or Page order did not survive round-trip")
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
		{"picture", func(d *Definition) { d.Controls[0].Properties = map[string]any{"Picture": int64(0xffff)} }, ErrUnsupportedEdit},
		{"root-caption-codepage", func(d *Definition) { d.Caption = "😀" }, ErrInvalidEdit},
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
