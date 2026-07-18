package lspserver

import (
	"unicode/utf8"

	protocol "github.com/tliron/glsp/protocol_3_16"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

// lineOffsetIndex maps LSP line numbers to byte offsets in the original source.
// It retains CRLF source bytes while treating CRLF as one line separator.
type lineOffsetIndex struct {
	starts []int
}

func newLineOffsetIndex(source string) *lineOffsetIndex {
	starts := []int{0}
	for offset := 0; offset < len(source); {
		switch source[offset] {
		case '\r':
			offset++
			if offset < len(source) && source[offset] == '\n' {
				offset++
			}
			starts = append(starts, offset)
		case '\n':
			offset++
			starts = append(starts, offset)
		default:
			_, size := utf8.DecodeRuneInString(source[offset:])
			offset += size
		}
	}
	return &lineOffsetIndex{starts: starts}
}

func (i *lineOffsetIndex) byteOffset(source string, pos protocol.Position) (int, bool) {
	line := int(pos.Line)
	if i == nil || line < 0 || line >= len(i.starts) {
		return 0, false
	}
	start := i.starts[line]
	end := len(source)
	if line+1 < len(i.starts) {
		end = i.starts[line+1]
		if end > start && source[end-1] == '\n' {
			end--
		}
		if end > start && source[end-1] == '\r' {
			end--
		}
	}
	column := int(pos.Character)
	if column < 0 {
		return 0, false
	}
	units := 0
	for offset := start; offset < end; {
		if units == column {
			return offset, true
		}
		runeValue, size := utf8.DecodeRuneInString(source[offset:end])
		if runeValue == utf8.RuneError && size == 1 {
			return 0, false
		}
		next := units + 1
		if runeValue > 0xffff {
			next++
		}
		if column < next { // A UTF-16 position may not split a surrogate pair.
			return 0, false
		}
		units = next
		offset += size
	}
	return end, units == column
}

type documentContentChange struct {
	rng  *protocol.Range
	text string
}

func decodeDocumentContentChanges(changes []any) ([]documentContentChange, bool) {
	out := make([]documentContentChange, 0, len(changes))
	for _, change := range changes {
		switch typed := change.(type) {
		case protocol.TextDocumentContentChangeEventWhole:
			out = append(out, documentContentChange{text: typed.Text})
		case protocol.TextDocumentContentChangeEvent:
			out = append(out, documentContentChange{rng: typed.Range, text: typed.Text})
		default:
			return nil, false
		}
	}
	return out, true
}

// prepareDocumentContentChanges applies changes in LSP client order and, when
// possible, returns the equivalent tree-sitter edits. LSP positions are
// UTF-16; tree-sitter points are UTF-8 byte columns in the raw source.
func prepareDocumentContentChanges(source string, index *lineOffsetIndex, changes []documentContentChange) (string, *lineOffsetIndex, []tree_sitter.InputEdit, bool, bool) {
	if index == nil {
		index = newLineOffsetIndex(source)
	}
	edits := make([]tree_sitter.InputEdit, 0, len(changes))
	canIncrementallyParse := true
	for _, change := range changes {
		if change.rng == nil {
			source = change.text
			index = newLineOffsetIndex(source)
			edits = nil
			canIncrementallyParse = false
			continue
		}
		start, ok := index.byteOffset(source, change.rng.Start)
		if !ok {
			return "", nil, nil, false, false
		}
		end, ok := index.byteOffset(source, change.rng.End)
		if !ok || end < start {
			return "", nil, nil, false, false
		}
		oldSource := source
		source = oldSource[:start] + change.text + oldSource[end:]
		if canIncrementallyParse {
			edits = append(edits, tree_sitter.InputEdit{
				StartByte:      uint(start),
				OldEndByte:     uint(end),
				NewEndByte:     uint(start + len(change.text)),
				StartPosition:  treeSitterPoint(oldSource, start),
				OldEndPosition: treeSitterPoint(oldSource, end),
				NewEndPosition: treeSitterPoint(source, start+len(change.text)),
			})
		}
		index = newLineOffsetIndex(source)
	}
	return source, index, edits, canIncrementallyParse && len(edits) > 0, true
}

// treeSitterPoint counts a point exactly as tree-sitter does: rows advance on
// LF and columns count UTF-8 bytes from the last LF. In particular, CRLF
// leaves the CR in the preceding byte column and then resets at LF; a bare CR
// is not a tree-sitter line break.
func treeSitterPoint(source string, offset int) tree_sitter.Point {
	var point tree_sitter.Point
	for index := 0; index < offset; index++ {
		if source[index] == '\n' {
			point.Row++
			point.Column = 0
			continue
		}
		point.Column++
	}
	return point
}

func applyDocumentContentChanges(source string, index *lineOffsetIndex, changes []documentContentChange) (string, *lineOffsetIndex, bool) {
	source, index, _, _, ok := prepareDocumentContentChanges(source, index, changes)
	return source, index, ok
}
