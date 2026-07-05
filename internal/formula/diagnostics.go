package formula

type DiagnosticSeverity string

const (
	DiagnosticSeverityWarning DiagnosticSeverity = "warning"
	DiagnosticSeverityError   DiagnosticSeverity = "error"
)

const (
	DiagnosticInvalidBaseCell    = "invalid_base_cell"
	DiagnosticUnterminatedString = "unterminated_string"
	DiagnosticUnterminatedQuoted = "unterminated_quoted_name"
	DiagnosticUnsupportedSyntax  = "unsupported_formula_syntax"
	DiagnosticInvalidReference   = "invalid_reference"
)

type Diagnostic struct {
	Code     string             `json:"code"`
	Severity DiagnosticSeverity `json:"severity"`
	Message  string             `json:"message"`
	Span     Span               `json:"span"`
}
