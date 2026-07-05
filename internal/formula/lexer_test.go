package formula

import "testing"

func TestLexPreservesStringsQuotedNamesErrorsAndPunctuation(t *testing.T) {
	tokens, diagnostics := Lex(`=IF(A1="B2",#N/A,'売上 集計'!C3)+SUM(A1:B2)`)
	if len(diagnostics) != 0 {
		t.Fatalf("diagnostics = %#v", diagnostics)
	}
	assertToken(t, tokens, TokenString, `"B2"`)
	assertToken(t, tokens, TokenError, "#N/A")
	assertToken(t, tokens, TokenQuotedName, `'売上 集計'`)
	for _, want := range []string{"=", "(", ",", "!", ":", ")"} {
		assertTokenText(t, tokens, want)
	}
	if got := rawTokens(tokens); got != `=IF(A1="B2",#N/A,'売上 集計'!C3)+SUM(A1:B2)` {
		t.Fatalf("raw tokens = %q", got)
	}
}

func TestLexEscapedQuotes(t *testing.T) {
	tokens, diagnostics := Lex(`="A1 ""quoted"" B2"`)
	if len(diagnostics) != 0 {
		t.Fatalf("diagnostics = %#v", diagnostics)
	}
	assertToken(t, tokens, TokenString, `"A1 ""quoted"" B2"`)
}

func TestLexEscapedQuotedSheetName(t *testing.T) {
	tokens, diagnostics := Lex(`='O''Brien'!A1`)
	if len(diagnostics) != 0 {
		t.Fatalf("diagnostics = %#v", diagnostics)
	}
	assertToken(t, tokens, TokenQuotedName, `'O''Brien'`)
}

func TestLexUnterminatedStringDiagnostic(t *testing.T) {
	_, diagnostics := Lex(`="A1`)
	if len(diagnostics) != 1 || diagnostics[0].Code != DiagnosticUnterminatedString {
		t.Fatalf("diagnostics = %#v", diagnostics)
	}
}

func assertToken(t *testing.T, tokens []Token, kind TokenKind, text string) {
	t.Helper()
	for _, token := range tokens {
		if token.Kind == kind && token.Text == text {
			return
		}
	}
	t.Fatalf("missing token kind=%s text=%q in %#v", kind, text, tokens)
}

func assertTokenText(t *testing.T, tokens []Token, text string) {
	t.Helper()
	for _, token := range tokens {
		if token.Text == text {
			return
		}
	}
	t.Fatalf("missing token text=%q in %#v", text, tokens)
}
