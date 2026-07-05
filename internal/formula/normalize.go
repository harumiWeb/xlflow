package formula

import "strings"

type NormalizeOptions struct {
	BaseCell CellRef
}

type ParseStatus string

const (
	ParseStatusOK      ParseStatus = "ok"
	ParseStatusPartial ParseStatus = "partial"
	ParseStatusFailed  ParseStatus = "failed"
)

type NormalizeResult struct {
	Formula     string       `json:"formula"`
	Status      ParseStatus  `json:"status"`
	References  []Reference  `json:"references"`
	Features    []Feature    `json:"features"`
	Diagnostics []Diagnostic `json:"diagnostics"`
}

func NormalizeA1ToR1C1Pattern(formula string, opts NormalizeOptions) NormalizeResult {
	result := NormalizeResult{
		Formula:    formula,
		Status:     ParseStatusOK,
		References: []Reference{},
		Features:   []Feature{},
	}
	if opts.BaseCell.Row <= 0 || opts.BaseCell.Col <= 0 {
		result.Status = ParseStatusFailed
		result.Diagnostics = append(result.Diagnostics, Diagnostic{
			Code:     DiagnosticInvalidBaseCell,
			Severity: DiagnosticSeverityError,
			Message:  "base cell row and column must be positive",
			Span:     Span{},
		})
		return result
	}

	tokens, diagnostics := Lex(formula)
	result.Diagnostics = append(result.Diagnostics, diagnostics...)
	var out strings.Builder
	for i := 0; i < len(tokens); {
		if feature, consumed, ok := detectExternalReference(tokens, i); ok {
			result.Features = append(result.Features, feature)
			out.WriteString(rawTokens(tokens[i : i+consumed]))
			i += consumed
			continue
		}
		if feature, consumed, ok := detect3DReference(tokens, i); ok {
			result.Features = append(result.Features, feature)
			out.WriteString(rawTokens(tokens[i : i+consumed]))
			i += consumed
			continue
		}
		if feature, diagnostic, consumed, ok := detectStructuredReference(tokens, i); ok {
			result.Features = append(result.Features, feature)
			if diagnostic != nil {
				result.Diagnostics = append(result.Diagnostics, *diagnostic)
			}
			out.WriteString(rawTokens(tokens[i : i+consumed]))
			i += consumed
			continue
		}
		if tokenText(tokens, i) == "@" {
			if ref, ok := parseReferenceAt(tokens, i+1, opts.BaseCell); ok {
				span := Span{Start: tokens[i].Span.Start, End: ref.ref.Span.End}
				result.Features = append(result.Features, Feature{
					Code: FeatureImplicitIntersection,
					Raw:  rawTokens(tokens[i : i+1+ref.consumed]),
					Span: span,
				})
				out.WriteString(rawTokens(tokens[i : i+1+ref.consumed]))
				i += 1 + ref.consumed
				continue
			}
		}
		if ref, ok := parseReferenceAt(tokens, i, opts.BaseCell); ok {
			if tokenText(tokens, i+ref.consumed) == "#" {
				result.Features = append(result.Features, Feature{
					Code: FeatureSpillReference,
					Raw:  rawTokens(tokens[i : i+ref.consumed+1]),
					Span: Span{Start: tokens[i].Span.Start, End: tokens[i+ref.consumed].Span.End},
				})
				out.WriteString(rawTokens(tokens[i : i+ref.consumed+1]))
				i += ref.consumed + 1
				continue
			}
			result.References = append(result.References, ref.ref)
			out.WriteString(ref.ref.Normalized)
			i += ref.consumed
			continue
		}
		out.WriteString(tokens[i].Text)
		i++
	}
	result.Formula = out.String()
	if len(result.Diagnostics) > 0 || len(result.Features) > 0 {
		result.Status = ParseStatusPartial
	}
	return result
}

func detectExternalReference(tokens []Token, pos int) (Feature, int, bool) {
	if tokenText(tokens, pos) == "[" {
		end := pos + 1
		for end < len(tokens) && tokenText(tokens, end) != "!" {
			end++
		}
		if end < len(tokens) && tokenText(tokens, end) == "!" {
			if ref, ok := parseReferenceAt(tokens, end+1, CellRef{Row: 1, Col: 1}); ok {
				consumed := end + 1 + ref.consumed - pos
				return featureFromTokens(FeatureExternalReference, tokens[pos:pos+consumed]), consumed, true
			}
		}
	}
	if pos+1 < len(tokens) && tokens[pos].Kind == TokenQuotedName && tokenText(tokens, pos+1) == "!" && strings.Contains(tokens[pos].Text, "[") {
		if ref, ok := parseReferenceAt(tokens, pos+2, CellRef{Row: 1, Col: 1}); ok {
			consumed := 2 + ref.consumed
			return featureFromTokens(FeatureExternalReference, tokens[pos:pos+consumed]), consumed, true
		}
	}
	return Feature{}, 0, false
}

func detect3DReference(tokens []Token, pos int) (Feature, int, bool) {
	if pos+4 >= len(tokens) {
		return Feature{}, 0, false
	}
	if tokens[pos].Kind != TokenIdentifier && tokens[pos].Kind != TokenQuotedName {
		return Feature{}, 0, false
	}
	if tokenText(tokens, pos+1) != ":" {
		return Feature{}, 0, false
	}
	if tokens[pos+2].Kind != TokenIdentifier && tokens[pos+2].Kind != TokenQuotedName {
		return Feature{}, 0, false
	}
	if tokenText(tokens, pos+3) != "!" {
		return Feature{}, 0, false
	}
	if ref, ok := parseReferenceAt(tokens, pos+4, CellRef{Row: 1, Col: 1}); ok {
		consumed := 4 + ref.consumed
		return featureFromTokens(Feature3DReference, tokens[pos:pos+consumed]), consumed, true
	}
	return Feature{}, 0, false
}

func detectStructuredReference(tokens []Token, pos int) (Feature, *Diagnostic, int, bool) {
	if pos+2 >= len(tokens) || tokens[pos].Kind != TokenIdentifier || tokenText(tokens, pos+1) != "[" {
		return Feature{}, nil, 0, false
	}
	depth := 0
	for i := pos + 1; i < len(tokens); i++ {
		switch tokenText(tokens, i) {
		case "[":
			depth++
		case "]":
			depth--
			if depth == 0 {
				consumed := i - pos + 1
				return featureFromTokens(FeatureStructuredReference, tokens[pos:pos+consumed]), nil, consumed, true
			}
		}
	}
	span := Span{Start: tokens[pos].Span.Start, End: tokens[len(tokens)-1].Span.End}
	feature := Feature{
		Code: FeatureStructuredReference,
		Raw:  rawTokens(tokens[pos:]),
		Span: span,
	}
	diagnostic := Diagnostic{
		Code:     DiagnosticUnterminatedStructuredReference,
		Severity: DiagnosticSeverityWarning,
		Message:  "unterminated structured reference",
		Span:     span,
	}
	return feature, &diagnostic, len(tokens) - pos, true
}

func featureFromTokens(code string, tokens []Token) Feature {
	return Feature{
		Code: code,
		Raw:  rawTokens(tokens),
		Span: Span{Start: tokens[0].Span.Start, End: tokens[len(tokens)-1].Span.End},
	}
}
