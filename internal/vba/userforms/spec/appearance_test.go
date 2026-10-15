package spec

import (
	"fmt"
	"testing"
)

func TestColorSemantics(t *testing.T) {
	for _, value := range []uint32{0, 0xffffff, 0x00332211, 0x8000000f, 0x80000005, 0x800000ff, 0x01010203, 0xffffffff} {
		color := ColorFromOLE(value)
		got, err := ColorOLE(color)
		if err != nil || got != value {
			t.Fatalf("%08x -> %+v -> %08x: %v", value, color, got, err)
		}
	}
	if got, _ := ColorOLE(FormSpecColor{Kind: "rgb", Value: "#123456"}); got != 0x00563412 {
		t.Fatalf("RGB byte order: %08x", got)
	}
	if ColorFromOLE(0x8000000f).Kind != "system" || ColorFromOLE(0x8100000f).Kind != "ole" {
		t.Fatal("system/raw semantics lost")
	}
}

func TestTypedAppearanceSource(t *testing.T) {
	for _, tc := range []struct {
		name, fields string
		valid        bool
	}{
		{"partial", `"font":{"bold":false,"charset":0,"name":"","size":9},"backColor":{"kind":"system","value":"buttonFace"}`, true},
		{"null", `"font":null,"foreColor":null,"textAlign":null`, true},
		{"unknownOLE", `"foreColor":{"kind":"ole","value":"0x81234567"}`, true},
		{"equalLegacy", `"backColor":{"kind":"system","value":"buttonFace"},"properties":{"BackColor":-2147483633}`, true},
		{"conflict", `"backColor":{"kind":"system","value":"buttonFace"},"properties":{"BackColor":0}`, false},
		{"color", `"foreColor":{"kind":"rgb","value":"red"}`, false},
		{"size", `"font":{"size":0}`, false},
		{"charset", `"font":{"charset":256}`, false},
		{"member", `"font":{"unknown":true}`, false},
		{"enum", `"textAlign":0`, false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			source := fmt.Sprintf(`{"schemaVersion":1,"kind":"xlflow.userform","basis":"designer","form":{"name":"Main"},"controls":[{"id":"a","name":"Label1","type":"Label",%s}]}`, tc.fields)
			parsed, err := ParseFormSpec(SpecInput{Format: "json"}, []byte(source))
			if (err == nil) != tc.valid {
				t.Fatalf("valid=%v error=%v", tc.valid, err)
			}
			if tc.name == "partial" && (parsed.Controls[0].Font.Bold == nil || *parsed.Controls[0].Font.Bold || parsed.Controls[0].Font.Italic != nil) {
				t.Fatal("partial/false state lost")
			}
		})
	}
}

func TestTypedAppearanceApplicability(t *testing.T) {
	for _, control := range []FormSpecControl{
		{ID: "a", Name: "Image1", Type: "Image", Font: &FormSpecFont{Bold: new(true)}},
		{ID: "a", Name: "Scroll1", Type: "ScrollBar", Font: &FormSpecFont{Bold: new(true)}},
		{ID: "a", Name: "Page1", Type: "Page", Font: &FormSpecFont{Bold: new(true)}},
		{ID: "a", Name: "Button1", Type: "CommandButton", TextAlign: new(2)},
		{ID: "a", Name: "Label1", Type: "Label", ProgID: "Vendor.Label.1", Font: &FormSpecFont{Bold: new(true)}},
	} {
		if err := ValidateFormSpecControl(control, "controls[0]"); err == nil {
			t.Fatalf("accepted unsupported appearance: %+v", control)
		}
	}
	contract := UserFormContract()
	p := contract.FormProperties["font"]
	p.Children["name"] = PropertyContract{}
	if UserFormContract().FormProperties["font"].Children["name"].ValueType != ValueTypeString {
		t.Fatal("contract child map aliases")
	}
}
