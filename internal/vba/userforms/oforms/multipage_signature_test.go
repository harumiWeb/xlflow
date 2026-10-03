package oforms

import (
	"errors"
	"testing"
)

func TestSerializerRejectsDerivedTabAndPageModelMutation(t *testing.T) {
	for _, change := range []struct {
		name   string
		mutate func(*Control)
	}{
		{"selection", func(c *Control) { c.MultiPage.Hidden.TabStrip.SelectedIndex = 0 }},
		{"tab caption", func(c *Control) { c.MultiPage.Hidden.TabStrip.Tabs[0].Caption = "changed" }},
		{"page order", func(c *Control) {
			c.MultiPage.Pages[0], c.MultiPage.Pages[1] = c.MultiPage.Pages[1], c.MultiPage.Pages[0]
		}},
		{"enabled", func(c *Control) { c.MultiPage.Properties.Mask ^= 1 << 3 }},
		{"reserved", func(c *Control) { c.MultiPage.Reserved[0] ^= 1 }},
		{"page reference", func(c *Control) { copy := *c.MultiPage.Pages[0]; c.MultiPage.Pages[0] = &copy }},
	} {
		t.Run(change.name, func(t *testing.T) {
			form, err := NewForm(Definition{Name: "Tabs", Size: Size{8000, 6000}, Controls: []ControlDefinition{{Name: "Multi", Class: 57, Visible: true, Size: Size{5000, 4000}, Tabs: &TabStrip{SelectedIndex: 1, Tabs: []Tab{{Name: "A", Caption: "A", Enabled: true, Visible: true}, {Name: "B", Caption: "B", Enabled: true, Visible: true}}}, Controls: []ControlDefinition{{Name: "PageA", Class: 7}, {Name: "PageB", Class: 7}}}}}, 932)
			if err != nil {
				t.Fatal(err)
			}
			change.mutate(form.Controls[0])
			if got, err := SerializeForm(form, 932); got != nil || !errors.Is(err, ErrUnsupportedMutation) {
				t.Fatalf("got=%v err=%v", got, err)
			}
		})
	}
}
