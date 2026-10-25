package intel

import (
	"cmp"
	"context"
	"errors"
	"fmt"
	"slices"
	"strings"
	"unicode/utf8"
)

// ErrInvalidCodeActionEdits indicates that code action edits cannot be applied
// together to the captured document revision.
var ErrInvalidCodeActionEdits = errors.New("invalid code action edits")

type codeActionEditWithOffsets struct {
	edit       CodeActionEdit
	start, end int
}

type codeActionSourceLine struct {
	start, end   int
	utf16Offsets []int
}

type codeActionSourceIndex struct {
	snapshot *AnalysisSnapshot
	lines    []codeActionSourceLine
}

// ApplyCodeActionEdits applies edits expressed in UTF-16 coordinates from the
// original document. It preserves all source bytes outside the edit ranges.
func ApplyCodeActionEdits(ctx context.Context, doc Document, edits []CodeActionEdit) (string, error) {
	if len(edits) == 0 {
		return doc.Source, nil
	}
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return "", err
	}

	validated, err := validateAndSortCodeActionEdits(ctx, doc, edits)
	if err != nil {
		return "", err
	}

	var result strings.Builder
	result.Grow(len(doc.Source))
	cursor := 0
	for _, edit := range validated {
		if err := ctx.Err(); err != nil {
			return "", err
		}
		result.WriteString(doc.Source[cursor:edit.start])
		result.WriteString(edit.edit.NewText)
		cursor = edit.end
	}
	if err := ctx.Err(); err != nil {
		return "", err
	}
	result.WriteString(doc.Source[cursor:])
	return result.String(), nil
}

func validateAndSortCodeActionEdits(ctx context.Context, doc Document, edits []CodeActionEdit) ([]codeActionEditWithOffsets, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	if len(edits) == 0 {
		return nil, nil
	}

	ordered := slices.Clone(edits)
	for i, edit := range ordered {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		if edit.Range.Start.Line < 0 || edit.Range.Start.Character < 0 || edit.Range.End.Line < 0 || edit.Range.End.Character < 0 {
			return nil, fmt.Errorf("%w: edit[%d] has a negative position", ErrInvalidCodeActionEdits, i)
		}
		if comparePosition(edit.Range.Start, edit.Range.End) > 0 {
			return nil, fmt.Errorf("%w: edit[%d] has a reversed range", ErrInvalidCodeActionEdits, i)
		}
	}

	index, err := newCodeActionSourceIndex(ctx, doc)
	if err != nil {
		return nil, err
	}
	validated := make([]codeActionEditWithOffsets, 0, len(ordered))
	for i, edit := range ordered {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		start, ok := index.offset(edit.Range.Start)
		if !ok {
			return nil, fmt.Errorf("%w: edit[%d] start is outside a UTF-16 boundary in the document", ErrInvalidCodeActionEdits, i)
		}
		end, ok := index.offset(edit.Range.End)
		if !ok {
			return nil, fmt.Errorf("%w: edit[%d] end is outside a UTF-16 boundary in the document", ErrInvalidCodeActionEdits, i)
		}
		if start > end {
			return nil, fmt.Errorf("%w: edit[%d] has a reversed range", ErrInvalidCodeActionEdits, i)
		}
		validated = append(validated, codeActionEditWithOffsets{edit: edit, start: start, end: end})
	}

	slices.SortFunc(validated, func(left, right codeActionEditWithOffsets) int {
		if result := cmp.Compare(left.start, right.start); result != 0 {
			return result
		}
		if result := cmp.Compare(left.end, right.end); result != 0 {
			return result
		}
		return strings.Compare(left.edit.NewText, right.edit.NewText)
	})
	for i := 1; i < len(validated); i++ {
		previous, current := validated[i-1], validated[i]
		if current.start < previous.end || current.start == previous.start {
			return nil, fmt.Errorf("%w: edits %d and %d overlap or share a start position", ErrInvalidCodeActionEdits, i-1, i)
		}
	}
	return validated, nil
}

func newCodeActionSourceIndex(ctx context.Context, doc Document) (codeActionSourceIndex, error) {
	if snapshot := analysisSnapshotForDocument(doc); snapshot != nil {
		return codeActionSourceIndex{snapshot: snapshot}, nil
	}

	index := codeActionSourceIndex{}
	for lineStart := 0; ; {
		if err := ctx.Err(); err != nil {
			return codeActionSourceIndex{}, err
		}
		lineEnd, nextLineStart := rawLineBounds(doc.Source, lineStart)
		line := doc.Source[lineStart:lineEnd]
		offsets := make([]int, 1, len(line)+1)
		offsets[0] = lineStart
		for byteOffset := 0; byteOffset < len(line); {
			if byteOffset&4095 == 0 {
				if err := ctx.Err(); err != nil {
					return codeActionSourceIndex{}, err
				}
			}
			r, size := utf8.DecodeRuneInString(line[byteOffset:])
			nextOffset := lineStart + byteOffset + size
			if r > 0xFFFF {
				offsets = append(offsets, -1)
			}
			offsets = append(offsets, nextOffset)
			byteOffset += size
		}
		index.lines = append(index.lines, codeActionSourceLine{start: lineStart, end: lineEnd, utf16Offsets: offsets})
		if nextLineStart == len(doc.Source) {
			if nextLineStart > lineEnd {
				index.lines = append(index.lines, codeActionSourceLine{start: nextLineStart, end: nextLineStart, utf16Offsets: []int{nextLineStart}})
			}
			break
		}
		lineStart = nextLineStart
	}
	return index, nil
}

func (index codeActionSourceIndex) offset(position Position) (int, bool) {
	if position.Line < 0 || position.Character < 0 {
		return 0, false
	}
	if index.snapshot != nil {
		offset := index.snapshot.byteOffset(position)
		return offset, index.snapshot.position(offset) == position
	}
	if position.Line >= len(index.lines) {
		return 0, false
	}
	line := index.lines[position.Line]
	if position.Character >= len(line.utf16Offsets) {
		return 0, false
	}
	offset := line.utf16Offsets[position.Character]
	return offset, offset >= line.start && offset <= line.end
}
