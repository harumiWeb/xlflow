package lspserver

import (
	"reflect"
	"testing"

	protocol "github.com/tliron/glsp/protocol_3_16"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
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

func TestPrepareDocumentContentChangesBuildsTreeSitterEdits(t *testing.T) {
	cases := []struct {
		name     string
		source   string
		change   documentContentChange
		want     string
		wantEdit tree_sitter.InputEdit
	}{
		{
			name:   "ascii",
			source: "abc",
			change: documentContentChange{rng: protocolRange(0, 1, 0, 2), text: "XYZ"},
			want:   "aXYZc",
			wantEdit: tree_sitter.InputEdit{StartByte: 1, OldEndByte: 2, NewEndByte: 4,
				StartPosition: tree_sitter.Point{Row: 0, Column: 1}, OldEndPosition: tree_sitter.Point{Row: 0, Column: 2}, NewEndPosition: tree_sitter.Point{Row: 0, Column: 4}},
		},
		{
			name:   "crlf",
			source: "aa\r\nbb",
			change: documentContentChange{rng: protocolRange(1, 1, 1, 2), text: "é"},
			want:   "aa\r\nbé",
			wantEdit: tree_sitter.InputEdit{StartByte: 5, OldEndByte: 6, NewEndByte: 7,
				StartPosition: tree_sitter.Point{Row: 1, Column: 1}, OldEndPosition: tree_sitter.Point{Row: 1, Column: 2}, NewEndPosition: tree_sitter.Point{Row: 1, Column: 3}},
		},
		{
			name:   "multiline",
			source: "one\ntwo\nthree",
			change: documentContentChange{rng: protocolRange(0, 3, 1, 0), text: "\nX\n"},
			want:   "one\nX\ntwo\nthree",
			wantEdit: tree_sitter.InputEdit{StartByte: 3, OldEndByte: 4, NewEndByte: 6,
				StartPosition: tree_sitter.Point{Row: 0, Column: 3}, OldEndPosition: tree_sitter.Point{Row: 1, Column: 0}, NewEndPosition: tree_sitter.Point{Row: 2, Column: 0}},
		},
		{
			name:   "bmp unicode",
			source: "あx",
			change: documentContentChange{rng: protocolRange(0, 1, 0, 2), text: "い"},
			want:   "あい",
			wantEdit: tree_sitter.InputEdit{StartByte: 3, OldEndByte: 4, NewEndByte: 6,
				StartPosition: tree_sitter.Point{Row: 0, Column: 3}, OldEndPosition: tree_sitter.Point{Row: 0, Column: 4}, NewEndPosition: tree_sitter.Point{Row: 0, Column: 6}},
		},
		{
			name:   "supplementary unicode",
			source: "😀x",
			change: documentContentChange{rng: protocolRange(0, 2, 0, 3), text: "y"},
			want:   "😀y",
			wantEdit: tree_sitter.InputEdit{StartByte: 4, OldEndByte: 5, NewEndByte: 5,
				StartPosition: tree_sitter.Point{Row: 0, Column: 4}, OldEndPosition: tree_sitter.Point{Row: 0, Column: 5}, NewEndPosition: tree_sitter.Point{Row: 0, Column: 5}},
		},
		{
			name:   "bare cr is a byte column not a tree sitter row",
			source: "a\rb",
			change: documentContentChange{rng: protocolRange(1, 0, 1, 1), text: "Z"},
			want:   "a\rZ",
			wantEdit: tree_sitter.InputEdit{StartByte: 2, OldEndByte: 3, NewEndByte: 3,
				StartPosition: tree_sitter.Point{Row: 0, Column: 2}, OldEndPosition: tree_sitter.Point{Row: 0, Column: 3}, NewEndPosition: tree_sitter.Point{Row: 0, Column: 3}},
		},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			got, _, edits, incremental, ok := prepareDocumentContentChanges(tc.source, newLineOffsetIndex(tc.source), []documentContentChange{tc.change})
			if !ok || !incremental || got != tc.want {
				t.Fatalf("prepare = (%q, incremental=%v, ok=%v), want (%q, true, true)", got, incremental, ok, tc.want)
			}
			if len(edits) != 1 || !reflect.DeepEqual(edits[0], tc.wantEdit) {
				t.Fatalf("edit = %#v, want %#v", edits, tc.wantEdit)
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
