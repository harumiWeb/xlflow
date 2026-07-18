package lspserver

import (
	"testing"

	protocol "github.com/tliron/glsp/protocol_3_16"
)

func TestApplyDocumentContentChanges(t *testing.T) {
	cases := []struct {
		name    string
		source  string
		changes []documentContentChange
		want    string
	}{
		{
			name:   "ascii insertion deletion and replacement",
			source: "alpha\nbeta\n",
			changes: []documentContentChange{
				{rng: protocolRange(0, 5, 0, 5), text: "!"},
				{rng: protocolRange(1, 0, 1, 1), text: ""},
				{rng: protocolRange(1, 0, 1, 3), text: "ETA"},
			},
			want: "alpha!\nETA\n",
		},
		{
			name:   "multiline delete complete line and final line replacement",
			source: "one\ntwo\nthree",
			changes: []documentContentChange{
				{rng: protocolRange(0, 3, 1, 0), text: "\n"},
				{rng: protocolRange(1, 0, 2, 5), text: "last"},
			},
			want: "one\nlast",
		},
		{
			name:   "japanese and supplementary unicode",
			source: "' 日本語 😀\nDim 名前 As String\n",
			changes: []documentContentChange{
				{rng: protocolRange(0, 6, 0, 8), text: "🚀"},
				{rng: protocolRange(1, 4, 1, 6), text: "値"},
			},
			want: "' 日本語 🚀\nDim 値 As String\n",
		},
		{
			name:    "crlf preserves delimiters",
			source:  "first\r\nsecond\r\nthird",
			changes: []documentContentChange{{rng: protocolRange(1, 0, 1, 6), text: "middle"}},
			want:    "first\r\nmiddle\r\nthird",
		},
		{
			name:    "full replacement followed by range",
			source:  "ignored",
			changes: []documentContentChange{{text: "abc"}, {rng: protocolRange(0, 1, 0, 2), text: "Z"}},
			want:    "aZc",
		},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			got, _, ok := applyDocumentContentChanges(tc.source, newLineOffsetIndex(tc.source), tc.changes)
			if !ok || got != tc.want {
				t.Fatalf("apply = (%q, ok=%v), want %q", got, ok, tc.want)
			}
		})
	}
}

func TestApplyDocumentContentChangesRejectsInvalidUTF16AndRange(t *testing.T) {
	for _, change := range []documentContentChange{
		{rng: protocolRange(0, 1, 0, 1), text: "x"}, // middle of 😀 surrogate pair
		{rng: protocolRange(1, 0, 1, 0), text: "x"},
		{rng: protocolRange(0, 2, 0, 1), text: "x"},
	} {
		if _, _, ok := applyDocumentContentChanges("😀", newLineOffsetIndex("😀"), []documentContentChange{change}); ok {
			t.Fatalf("invalid change %+v was accepted", change)
		}
	}
}

func protocolRange(startLine, startCharacter, endLine, endCharacter int) *protocol.Range {
	return &protocol.Range{
		Start: protocol.Position{Line: protocol.UInteger(startLine), Character: protocol.UInteger(startCharacter)},
		End:   protocol.Position{Line: protocol.UInteger(endLine), Character: protocol.UInteger(endCharacter)},
	}
}
