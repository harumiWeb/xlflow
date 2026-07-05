package formula

import (
	"strings"
	"unicode"
	"unicode/utf8"
)

func Lex(formula string) ([]Token, []Diagnostic) {
	tokens := make([]Token, 0, len(formula)/2)
	diagnostics := []Diagnostic{}
	for i := 0; i < len(formula); {
		r, size := utf8.DecodeRuneInString(formula[i:])
		if r == utf8.RuneError && size == 1 {
			tokens = append(tokens, Token{Kind: TokenUnknown, Text: formula[i : i+1], Span: Span{Start: i, End: i + 1}})
			i++
			continue
		}
		start := i
		switch {
		case unicode.IsSpace(r):
			i += size
			for i < len(formula) {
				next, nextSize := utf8.DecodeRuneInString(formula[i:])
				if !unicode.IsSpace(next) {
					break
				}
				i += nextSize
			}
			tokens = append(tokens, Token{Kind: TokenWhitespace, Text: formula[start:i], Span: Span{Start: start, End: i}})
		case r == '"':
			end, ok := scanDoubleQuoted(formula, i)
			i = end
			tokens = append(tokens, Token{Kind: TokenString, Text: formula[start:i], Span: Span{Start: start, End: i}})
			if !ok {
				diagnostics = append(diagnostics, Diagnostic{
					Code:     DiagnosticUnterminatedString,
					Severity: DiagnosticSeverityWarning,
					Message:  "unterminated formula string literal",
					Span:     Span{Start: start, End: i},
				})
			}
		case r == '\'':
			end, ok := scanSingleQuoted(formula, i)
			i = end
			tokens = append(tokens, Token{Kind: TokenQuotedName, Text: formula[start:i], Span: Span{Start: start, End: i}})
			if !ok {
				diagnostics = append(diagnostics, Diagnostic{
					Code:     DiagnosticUnterminatedQuoted,
					Severity: DiagnosticSeverityWarning,
					Message:  "unterminated quoted sheet name",
					Span:     Span{Start: start, End: i},
				})
			}
		case r == '#':
			end := scanErrorLiteral(formula, i)
			if end > i+size {
				i = end
				tokens = append(tokens, Token{Kind: TokenError, Text: formula[start:i], Span: Span{Start: start, End: i}})
			} else {
				i += size
				tokens = append(tokens, Token{Kind: TokenPunctuation, Text: formula[start:i], Span: Span{Start: start, End: i}})
			}
		case isIdentifierStart(r):
			i += size
			for i < len(formula) {
				next, nextSize := utf8.DecodeRuneInString(formula[i:])
				if !isIdentifierPart(next) {
					break
				}
				i += nextSize
			}
			tokens = append(tokens, Token{Kind: TokenIdentifier, Text: formula[start:i], Span: Span{Start: start, End: i}})
		case unicode.IsDigit(r):
			i += size
			for i < len(formula) {
				next, nextSize := utf8.DecodeRuneInString(formula[i:])
				if !unicode.IsDigit(next) && next != '.' {
					break
				}
				i += nextSize
			}
			tokens = append(tokens, Token{Kind: TokenNumber, Text: formula[start:i], Span: Span{Start: start, End: i}})
		case strings.ContainsRune("(){}[],;:!$@", r):
			i += size
			tokens = append(tokens, Token{Kind: TokenPunctuation, Text: formula[start:i], Span: Span{Start: start, End: i}})
		case strings.ContainsRune("+-*/^&=<>%", r):
			i += size
			if i < len(formula) {
				next, nextSize := utf8.DecodeRuneInString(formula[i:])
				if (r == '<' && (next == '>' || next == '=')) || (r == '>' && next == '=') {
					i += nextSize
				}
			}
			tokens = append(tokens, Token{Kind: TokenOperator, Text: formula[start:i], Span: Span{Start: start, End: i}})
		default:
			i += size
			tokens = append(tokens, Token{Kind: TokenUnknown, Text: formula[start:i], Span: Span{Start: start, End: i}})
		}
	}
	return tokens, diagnostics
}

func scanDoubleQuoted(s string, start int) (int, bool) {
	for i := start + 1; i < len(s); {
		if s[i] != '"' {
			_, size := utf8.DecodeRuneInString(s[i:])
			i += size
			continue
		}
		if i+1 < len(s) && s[i+1] == '"' {
			i += 2
			continue
		}
		return i + 1, true
	}
	return len(s), false
}

func scanSingleQuoted(s string, start int) (int, bool) {
	for i := start + 1; i < len(s); {
		if s[i] != '\'' {
			_, size := utf8.DecodeRuneInString(s[i:])
			i += size
			continue
		}
		if i+1 < len(s) && s[i+1] == '\'' {
			i += 2
			continue
		}
		return i + 1, true
	}
	return len(s), false
}

func scanErrorLiteral(s string, start int) int {
	i := start + 1
	for i < len(s) {
		r, size := utf8.DecodeRuneInString(s[i:])
		if !unicode.IsLetter(r) && r != '/' && r != '?' && !unicode.IsDigit(r) {
			break
		}
		i += size
	}
	if i < len(s) && s[i] == '!' {
		return i + 1
	}
	if i > start+1 {
		return i
	}
	return start + 1
}

func isIdentifierStart(r rune) bool {
	return unicode.IsLetter(r) || r == '_' || r == '\\'
}

func isIdentifierPart(r rune) bool {
	return unicode.IsLetter(r) || unicode.IsDigit(r) || r == '_' || r == '.' || r == '\\'
}
