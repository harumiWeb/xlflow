package intel

import (
	"strings"
	"testing"
	"unicode/utf16"
)

func TestParseJSONIndexesValidationIssuePathsWithUTF16Ranges(t *testing.T) {
	const source = `{"schemaVersion":1,"form":{"caption":"😀"},"warnings":[{"code":"unsupported_properties","message":"Unsupported \"caption\" 😀.\nCaptured.","control":"😀"}]}`
	for _, test := range []struct {
		name   string
		source string
	}{
		{name: "single-line-unicode-and-escapes", source: source},
		{name: "crlf", source: strings.ReplaceAll(source, `,"`, ",\r\n\"")},
	} {
		t.Run(test.name, func(t *testing.T) {
			doc := ParseJSON(test.source)
			if doc.ParseError != nil {
				t.Fatal(doc.ParseError)
			}
			field, ok := doc.Field("warnings[0].message")
			if !ok {
				t.Fatal("warnings[0].message was not indexed")
			}
			start, end := jsonTestStringRange(test.source, `"Unsupported`)
			want := Range{Start: jsonTestPositionAt(test.source, start), End: jsonTestPositionAt(test.source, end)}
			if field.ValueRange != want {
				t.Fatalf("message range = %#v, want %#v", field.ValueRange, want)
			}
			if field.KeyRange.Start.Line != field.ValueRange.Start.Line {
				t.Fatalf("message key/value lines differ unexpectedly: %#v", field)
			}
		})
	}
	if doc := ParseJSON(`{"warnings":[]} {}`); doc.ParseError == nil {
		t.Fatal("multiple JSON values were accepted")
	}
}

func jsonTestStringRange(source, marker string) (int, int) {
	start := strings.Index(source, marker)
	for offset := start + 1; offset < len(source); offset++ {
		switch source[offset] {
		case '\\':
			offset++
		case '"':
			return start, offset + 1
		}
	}
	return start, len(source)
}

func jsonTestPositionAt(source string, offset int) Position {
	prefix := source[:offset]
	line := strings.Count(prefix, "\n")
	lineStart := strings.LastIndex(prefix, "\n") + 1
	return Position{Line: line, Character: len(utf16.Encode([]rune(prefix[lineStart:])))}
}
