package intel

import (
	"context"
	"errors"
	"slices"
	"strings"
	"testing"
)

func TestApplyCodeActionEditsSortsOriginalRangesAndPreservesSource(t *testing.T) {
	for _, eol := range []string{"\n", "\r\n"} {
		name := "LF"
		if eol == "\r\n" {
			name = "CRLF"
		}
		t.Run(name, func(t *testing.T) {
			source := strings.Join([]string{"first 😀", "second 日本語", "third"}, eol)
			doc := Document{Source: source}
			edits := []CodeActionEdit{
				{Range: Range{Start: Position{Line: 2}, End: Position{Line: 2, Character: 5}}, NewText: "THIRD"},
				{Range: Range{Start: Position{Line: 1, Character: 7}, End: Position{Line: 1, Character: 10}}, NewText: "日本語"},
				{Range: Range{Start: Position{Line: 0, Character: 6}, End: Position{Line: 0, Character: 8}}, NewText: "🌟"},
				{Range: Range{Start: Position{Line: 1, Character: 6}, End: Position{Line: 1, Character: 6}}, NewText: "!"},
				{Range: Range{Start: Position{Line: 0}, End: Position{Line: 0, Character: 5}}, NewText: "FIRST"},
			}
			originalEdits := slices.Clone(edits)

			got, err := ApplyCodeActionEdits(t.Context(), doc, edits)
			if err != nil {
				t.Fatal(err)
			}
			want := strings.Join([]string{"FIRST 🌟", "second! 日本語", "THIRD"}, eol)
			if got != want {
				t.Fatalf("applied source = %q, want %q", got, want)
			}
			if !slices.Equal(edits, originalEdits) {
				t.Fatalf("ApplyCodeActionEdits mutated input edits: %+v", edits)
			}
			repeated, err := ApplyCodeActionEdits(t.Context(), doc, edits)
			if err != nil || repeated != got {
				t.Fatalf("repeated application = %q, err = %v; want deterministic result %q", repeated, err, got)
			}
		})
	}
}

func TestApplyCodeActionEditsEmptyIsNoop(t *testing.T) {
	source := "日本語\r\n😀\r"
	got, err := ApplyCodeActionEdits(t.Context(), Document{Source: source}, nil)
	if err != nil || got != source {
		t.Fatalf("empty edit result = %q, err = %v; want original source %q", got, err, source)
	}
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	got, err = ApplyCodeActionEdits(ctx, Document{Source: source}, nil)
	if err != nil || got != source {
		t.Fatalf("empty edit result with canceled context = %q, err = %v; want original source %q", got, err, source)
	}
}

func TestApplyCodeActionEditsUsesDocumentSnapshot(t *testing.T) {
	doc := Document{Source: "a😀b"}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc = snapshot.Document()
	edits := []CodeActionEdit{{
		Range:   Range{Start: Position{Character: 1}, End: Position{Character: 3}},
		NewText: "x",
	}}
	got, err := ApplyCodeActionEdits(t.Context(), doc, edits)
	if err != nil || got != "axb" {
		t.Fatalf("snapshot-backed application = %q, err = %v; want %q", got, err, "axb")
	}
}

func TestApplyCodeActionEditsRejectsInvalidRanges(t *testing.T) {
	validEdit := func(start, end Position) CodeActionEdit {
		return CodeActionEdit{Range: Range{Start: start, End: end}, NewText: "x"}
	}
	for _, tc := range []struct {
		name   string
		source string
		edits  []CodeActionEdit
	}{
		{name: "negative line", source: "abc", edits: []CodeActionEdit{validEdit(Position{Line: -1}, Position{})}},
		{name: "negative character", source: "abc", edits: []CodeActionEdit{validEdit(Position{Character: -1}, Position{})}},
		{name: "line beyond document", source: "abc", edits: []CodeActionEdit{validEdit(Position{Line: 1}, Position{Line: 1})}},
		{name: "character beyond line", source: "abc", edits: []CodeActionEdit{validEdit(Position{Character: 2}, Position{Character: 4})}},
		{name: "surrogate split at start", source: "a😀b", edits: []CodeActionEdit{validEdit(Position{Character: 2}, Position{Character: 3})}},
		{name: "surrogate split at end", source: "a😀b", edits: []CodeActionEdit{validEdit(Position{Character: 1}, Position{Character: 2})}},
		{name: "reversed range", source: "abc", edits: []CodeActionEdit{validEdit(Position{Character: 2}, Position{Character: 1})}},
		{name: "overlap", source: "abcdef", edits: []CodeActionEdit{
			validEdit(Position{}, Position{Character: 3}),
			validEdit(Position{Character: 2}, Position{Character: 4}),
		}},
		{name: "duplicate", source: "abcdef", edits: []CodeActionEdit{
			validEdit(Position{Character: 1}, Position{Character: 3}),
			validEdit(Position{Character: 1}, Position{Character: 3}),
		}},
		{name: "same-point insertions", source: "abcdef", edits: []CodeActionEdit{
			validEdit(Position{Character: 2}, Position{Character: 2}),
			validEdit(Position{Character: 2}, Position{Character: 2}),
		}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			if _, err := ApplyCodeActionEdits(t.Context(), Document{Source: tc.source}, tc.edits); !errors.Is(err, ErrInvalidCodeActionEdits) {
				t.Fatalf("ApplyCodeActionEdits error = %v, want ErrInvalidCodeActionEdits", err)
			}
		})
	}
}

func TestApplyCodeActionEditsHonorsCancellation(t *testing.T) {
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	_, err := ApplyCodeActionEdits(ctx, Document{Source: "abc"}, []CodeActionEdit{{
		Range: Range{Start: Position{}, End: Position{Character: 1}}, NewText: "x",
	}})
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("ApplyCodeActionEdits error = %v, want context.Canceled", err)
	}
}
