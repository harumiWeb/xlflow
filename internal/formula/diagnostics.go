package formula

type DiagnosticSeverity string

const (
	DiagnosticSeverityWarning DiagnosticSeverity = "warning"
	DiagnosticSeverityError   DiagnosticSeverity = "error"
)

type DiagnosticCode string

const (
	DiagnosticInvalidBaseCell                 DiagnosticCode = "invalid_base_cell"
	DiagnosticUnterminatedString              DiagnosticCode = "unterminated_string"
	DiagnosticUnterminatedQuoted              DiagnosticCode = "unterminated_quoted_name"
	DiagnosticUnterminatedStructuredReference DiagnosticCode = "unterminated_structured_reference"
	DiagnosticUnsupportedSyntax               DiagnosticCode = "unsupported_formula_syntax"
	DiagnosticInvalidReference                DiagnosticCode = "invalid_reference"
)

type Diagnostic struct {
	Code     DiagnosticCode     `json:"code"`
	Severity DiagnosticSeverity `json:"severity"`
	Message  string             `json:"message"`
	Span     Span               `json:"span"`
}
