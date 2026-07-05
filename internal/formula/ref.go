package formula

import (
	"fmt"
	"strconv"
	"strings"
	"unicode"

	"github.com/xuri/excelize/v2"
)

type CellRef struct {
	Row int `json:"row"`
	Col int `json:"col"`
}

type ReferenceKind string

const (
	ReferenceKindCell        ReferenceKind = "cell"
	ReferenceKindRange       ReferenceKind = "range"
	ReferenceKindColumnRange ReferenceKind = "column_range"
	ReferenceKindRowRange    ReferenceKind = "row_range"
)

type RefEndpoint struct {
	Row    int  `json:"row,omitempty"`
	Col    int  `json:"col,omitempty"`
	RowAbs bool `json:"row_abs,omitempty"`
	ColAbs bool `json:"col_abs,omitempty"`
}

type Reference struct {
	Kind       ReferenceKind `json:"kind"`
	Raw        string        `json:"raw"`
	Normalized string        `json:"normalized"`
	Sheet      string        `json:"sheet,omitempty"`
	Start      RefEndpoint   `json:"start"`
	End        *RefEndpoint  `json:"end,omitempty"`
	Span       Span          `json:"span"`
}

type parseRefResult struct {
	ref      Reference
	consumed int
}

func parseReferenceAt(tokens []Token, pos int, base CellRef) (parseRefResult, bool) {
	start := pos
	sheet := ""
	if qualifier, consumed, ok := parseSheetQualifier(tokens, pos); ok {
		sheet = qualifier
		pos += consumed
	}

	if endpoint, consumed, ok := parseCellEndpoint(tokens, pos); ok {
		if isFunctionCall(tokens, pos+consumed) && sheet == "" {
			return parseRefResult{}, false
		}
		pos += consumed
		kind := ReferenceKindCell
		endpoint2 := (*RefEndpoint)(nil)
		if tokenText(tokens, pos) == ":" {
			if second, secondConsumed, ok := parseCellEndpoint(tokens, pos+1); ok {
				endpoint2 = &second
				kind = ReferenceKindRange
				pos += 1 + secondConsumed
			}
		}
		ref := buildReference(tokens, start, pos, sheet, kind, endpoint, endpoint2, base)
		return parseRefResult{ref: ref, consumed: pos - start}, true
	}

	if first, consumed, ok := parseColumnEndpoint(tokens, pos); ok && tokenText(tokens, pos+consumed) == ":" {
		if second, secondConsumed, ok := parseColumnEndpoint(tokens, pos+consumed+1); ok {
			pos += consumed + 1 + secondConsumed
			ref := buildReference(tokens, start, pos, sheet, ReferenceKindColumnRange, first, &second, base)
			return parseRefResult{ref: ref, consumed: pos - start}, true
		}
	}

	if first, consumed, ok := parseRowEndpoint(tokens, pos); ok && tokenText(tokens, pos+consumed) == ":" {
		if second, secondConsumed, ok := parseRowEndpoint(tokens, pos+consumed+1); ok {
			pos += consumed + 1 + secondConsumed
			ref := buildReference(tokens, start, pos, sheet, ReferenceKindRowRange, first, &second, base)
			return parseRefResult{ref: ref, consumed: pos - start}, true
		}
	}

	return parseRefResult{}, false
}

func parseSheetQualifier(tokens []Token, pos int) (string, int, bool) {
	if pos+1 >= len(tokens) || tokenText(tokens, pos+1) != "!" {
		return "", 0, false
	}
	switch tokens[pos].Kind {
	case TokenIdentifier, TokenQuotedName:
		return tokens[pos].Text, 2, true
	default:
		return "", 0, false
	}
}

func parseCellEndpoint(tokens []Token, pos int) (RefEndpoint, int, bool) {
	start := pos
	colAbs := false
	rowAbs := false
	if tokenText(tokens, pos) == "$" {
		colAbs = true
		pos++
	}
	if pos >= len(tokens) || tokens[pos].Kind != TokenIdentifier {
		return RefEndpoint{}, 0, false
	}
	col, rowText, combined := splitCellText(tokens[pos].Text)
	if col == "" {
		return RefEndpoint{}, 0, false
	}
	pos++
	if rowText == "" {
		if tokenText(tokens, pos) == "$" {
			rowAbs = true
			pos++
		}
		if pos >= len(tokens) || tokens[pos].Kind != TokenNumber || !isIntegerText(tokens[pos].Text) {
			return RefEndpoint{}, 0, false
		}
		rowText = tokens[pos].Text
		pos++
	} else if combined && strings.HasPrefix(rowText, "$") {
		rowAbs = true
		rowText = strings.TrimPrefix(rowText, "$")
	}
	colNumber, err := excelize.ColumnNameToNumber(strings.ToUpper(col))
	if err != nil {
		return RefEndpoint{}, 0, false
	}
	rowNumber, err := strconv.Atoi(rowText)
	if err != nil || rowNumber <= 0 {
		return RefEndpoint{}, 0, false
	}
	return RefEndpoint{Row: rowNumber, Col: colNumber, RowAbs: rowAbs, ColAbs: colAbs}, pos - start, true
}

func parseColumnEndpoint(tokens []Token, pos int) (RefEndpoint, int, bool) {
	start := pos
	abs := false
	if tokenText(tokens, pos) == "$" {
		abs = true
		pos++
	}
	if pos >= len(tokens) || tokens[pos].Kind != TokenIdentifier || !isColumnText(tokens[pos].Text) {
		return RefEndpoint{}, 0, false
	}
	col, err := excelize.ColumnNameToNumber(strings.ToUpper(tokens[pos].Text))
	if err != nil {
		return RefEndpoint{}, 0, false
	}
	return RefEndpoint{Col: col, ColAbs: abs}, pos - start + 1, true
}

func parseRowEndpoint(tokens []Token, pos int) (RefEndpoint, int, bool) {
	start := pos
	abs := false
	if tokenText(tokens, pos) == "$" {
		abs = true
		pos++
	}
	if pos >= len(tokens) || tokens[pos].Kind != TokenNumber || !isIntegerText(tokens[pos].Text) {
		return RefEndpoint{}, 0, false
	}
	row, err := strconv.Atoi(tokens[pos].Text)
	if err != nil || row <= 0 {
		return RefEndpoint{}, 0, false
	}
	return RefEndpoint{Row: row, RowAbs: abs}, pos - start + 1, true
}

func buildReference(tokens []Token, start, end int, sheet string, kind ReferenceKind, first RefEndpoint, second *RefEndpoint, base CellRef) Reference {
	raw := rawTokens(tokens[start:end])
	normalized := renderReference(kind, sheet, first, second, base)
	return Reference{
		Kind:       kind,
		Raw:        raw,
		Normalized: normalized,
		Sheet:      sheet,
		Start:      first,
		End:        second,
		Span:       Span{Start: tokens[start].Span.Start, End: tokens[end-1].Span.End},
	}
}

func renderReference(kind ReferenceKind, sheet string, first RefEndpoint, second *RefEndpoint, base CellRef) string {
	prefix := ""
	if sheet != "" {
		prefix = sheet + "!"
	}
	switch kind {
	case ReferenceKindCell:
		return prefix + renderCell(first, base)
	case ReferenceKindRange:
		return prefix + renderCell(first, base) + ":" + renderCell(*second, base)
	case ReferenceKindColumnRange:
		return prefix + renderColumn(first, base) + ":" + renderColumn(*second, base)
	case ReferenceKindRowRange:
		return prefix + renderRow(first, base) + ":" + renderRow(*second, base)
	default:
		return prefix + fmt.Sprintf("%v", first)
	}
}

func renderCell(endpoint RefEndpoint, base CellRef) string {
	return renderRow(endpoint, base) + renderColumn(endpoint, base)
}

func renderRow(endpoint RefEndpoint, base CellRef) string {
	if endpoint.RowAbs {
		return fmt.Sprintf("R%d", endpoint.Row)
	}
	offset := endpoint.Row - base.Row
	if offset == 0 {
		return "R"
	}
	return fmt.Sprintf("R[%d]", offset)
}

func renderColumn(endpoint RefEndpoint, base CellRef) string {
	if endpoint.ColAbs {
		return fmt.Sprintf("C%d", endpoint.Col)
	}
	offset := endpoint.Col - base.Col
	if offset == 0 {
		return "C"
	}
	return fmt.Sprintf("C[%d]", offset)
}

func splitCellText(text string) (col string, row string, combined bool) {
	i := 0
	for i < len(text) {
		r := rune(text[i])
		if r > unicode.MaxASCII || !unicode.IsLetter(r) {
			break
		}
		i++
	}
	if i == 0 {
		return "", "", false
	}
	if i == len(text) {
		return text, "", false
	}
	for j := i; j < len(text); j++ {
		if text[j] < '0' || text[j] > '9' {
			return "", "", false
		}
	}
	return text[:i], text[i:], true
}

func isColumnText(text string) bool {
	if text == "" {
		return false
	}
	for _, r := range text {
		if r > unicode.MaxASCII || !unicode.IsLetter(r) {
			return false
		}
	}
	return true
}

func isIntegerText(text string) bool {
	if text == "" {
		return false
	}
	for _, r := range text {
		if r < '0' || r > '9' {
			return false
		}
	}
	return true
}

func isFunctionCall(tokens []Token, pos int) bool {
	return tokenText(tokens, pos) == "("
}

func tokenText(tokens []Token, pos int) string {
	if pos < 0 || pos >= len(tokens) {
		return ""
	}
	return tokens[pos].Text
}

func rawTokens(tokens []Token) string {
	var b strings.Builder
	for _, token := range tokens {
		b.WriteString(token.Text)
	}
	return b.String()
}
