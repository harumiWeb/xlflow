package formula

type TokenKind string

const (
	TokenIdentifier  TokenKind = "identifier"
	TokenNumber      TokenKind = "number"
	TokenString      TokenKind = "string"
	TokenQuotedName  TokenKind = "quoted_name"
	TokenError       TokenKind = "error"
	TokenPunctuation TokenKind = "punctuation"
	TokenOperator    TokenKind = "operator"
	TokenWhitespace  TokenKind = "whitespace"
	TokenUnknown     TokenKind = "unknown"
)

type Span struct {
	Start int `json:"start"`
	End   int `json:"end"`
}

type Token struct {
	Kind TokenKind `json:"kind"`
	Text string    `json:"text"`
	Span Span      `json:"span"`
}
