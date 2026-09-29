package ovba

import "testing"

func TestSupportedCodePagesRoundTrip(t *testing.T) {
	cases := []struct {
		name     string
		codepage uint16
		text     string
	}{
		{"cp874-thai", 874, "\u0e01\u0e02"},
		{"cp932-japanese", 932, "\u6a19\u6e96\u30e2\u30b8\u30e5\u30fc\u30eb"},
		{"cp936-chinese", 936, "\u6a21\u5757"},
		{"cp949-korean", 949, "\ubaa8\ub4c8"},
		{"cp950-big5", 950, "\u6a21\u7d44"},
		{"cp1250-czech", 1250, "\u017e\u00fd"},
		{"cp1251-cyrillic", 1251, "\u041c\u043e\u0434\u0443\u043b\u044c"},
		{"cp1252-latin", 1252, "caf\u00e9"},
		{"cp1253-greek", 1253, "\u03b1\u03b2"},
		{"cp1254-turkish", 1254, "\u011f\u00fc"},
		{"cp1255-hebrew", 1255, "\u05d0\u05d1"},
		{"cp1256-arabic", 1256, "\u0627\u0628"},
		{"cp1257-baltic", 1257, "\u0105\u0117"},
		{"cp1258-vietnamese", 1258, "\u00e2\u0103"},
		{"utf8", 65001, "caf\u00e9 \u65e5\u672c\u8a9e"},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			enc, err := EncodeMBCS(c.text, c.codepage)
			if err != nil {
				t.Fatalf("encode: %v", err)
			}
			back, err := DecodeMBCS(enc, c.codepage)
			if err != nil {
				t.Fatalf("decode: %v", err)
			}
			if back != c.text {
				t.Errorf("round-trip drift: want %q got %q", c.text, back)
			}
		})
	}
}

func TestDecodeMBCSRejections(t *testing.T) {
	if _, err := DecodeMBCS([]byte{0x82, 0xA0}, 932); err != nil {
		t.Fatalf("cp932 decode: %v", err)
	}
	if _, err := DecodeMBCS([]byte("plain"), 1200); err == nil {
		t.Error("codepage 1200 (UTF-16) must stay unsupported")
	}
	if _, err := DecodeMBCS([]byte{0xFF, 0xFE}, 65001); err == nil {
		t.Error("invalid UTF-8 must be rejected under codepage 65001")
	}
	if _, err := DecodeMBCS([]byte{0x82}, 932); err == nil {
		t.Error("truncated cp932 sequence must be rejected")
	}
}

func TestEncodeMBCSRejections(t *testing.T) {
	if _, err := EncodeMBCS("x", 1200); err == nil {
		t.Error("codepage 1200 should return an error")
	}
	if _, err := EncodeMBCS("\u65e5\u672c\u8a9e", 1252); err == nil {
		t.Error("characters not representable in CP1252 should return an error")
	}
	if _, err := EncodeMBCS("x \U0001F600 y", 932); err == nil {
		t.Error("emoji not representable in CP932 should return an error")
	}
	if _, err := EncodeMBCS("x", 37); err == nil {
		t.Error("unsupported codepage 37 should return an error")
	}
}
