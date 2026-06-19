package vbaproject

import "testing"

func TestDecodeCP932(t *testing.T) {
	// CP932 bytes 0x82 0xA0 decode to U+3042 (Japanese hiragana "a").
	got, err := decodeMBCS([]byte{0x82, 0xA0}, 932)
	if err != nil || got != "\u3042" {
		t.Errorf("decodeMBCS(932) = %q, %v; want %q", got, err, "\u3042")
	}
}

func TestDecodeASCIIAnyCodepage(t *testing.T) {
	got, _ := decodeMBCS([]byte("Sub X"), 932)
	if got != "Sub X" {
		t.Errorf("ASCII decode = %q", got)
	}
}

func TestDecodeRejectsUTF16(t *testing.T) {
	// 1200 (UTF-16) is not passed through and returns an error (prevents silent corruption from misreading as UTF-8).
	if _, err := decodeMBCS([]byte{0x41, 0x00}, 1200); err == nil {
		t.Error("codepage 1200 should return an unsupported-codepage error")
	}
}

func TestEncodeMBCSRoundTrip(t *testing.T) {
	cases := []struct {
		name     string
		codepage uint16
		s        string
	}{
		{"cp932-ascii", 932, "Attribute VB_Name = \"M\"\r\nSub X()\r\nEnd Sub\r\n"},
		{"cp932-japanese", 932, "Attribute VB_Name = \"M\"\r\n' \u65e5\u672c\u8a9e\u30b3\u30e1\u30f3\u30c8 \u30c6\u30b9\u30c8\r\n"},
		{"cp1252-latin", 1252, "Attribute VB_Name = \"M\"\r\n' café façade\r\n"},
		{"utf8-passthrough", 65001, "plain ascii"},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			enc, err := encodeMBCS(c.s, c.codepage)
			if err != nil {
				t.Fatalf("encode: %v", err)
			}
			back, err := decodeMBCS(enc, c.codepage)
			if err != nil {
				t.Fatalf("decode: %v", err)
			}
			if back != c.s {
				t.Errorf("round-trip drift:\n want %q\n got  %q", c.s, back)
			}
		})
	}
}

func TestEncodeMBCSErrors(t *testing.T) {
	if _, err := encodeMBCS("x", 1200); err == nil {
		t.Error("codepage 1200 (UTF-16) should return an error")
	}
	// Characters that CP1252 cannot represent (Japanese) return an error.
	if _, err := encodeMBCS("\u65e5\u672c\u8a9e", 1252); err == nil {
		t.Error("characters not representable in CP1252 should return an error")
	}
}
