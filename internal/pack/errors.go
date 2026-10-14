package pack

import (
	"errors"
	"fmt"
)

var (
	// ErrProtectedProject reports that the template contains a protected VBA project.
	ErrProtectedProject = errors.New("pack: protected VBA project")

	// ErrSignedProject reports that the template contains VBA signature streams or OOXML parts.
	ErrSignedProject = errors.New("pack: signed VBA project")

	// ErrAmbiguousLayout reports an unknown, unsupported, or ambiguous VBA project layout.
	ErrAmbiguousLayout = errors.New("pack: ambiguous VBA project layout")
)

// SourceEncodingError reports one source character that the target VBA
// project's code page cannot represent. Columns are one-based Unicode
// character columns; ByteOffset and ByteColumn retain the UTF-8 byte position.
type SourceEncodingError struct {
	Path       string
	Field      string
	Line       int
	Column     int
	ByteColumn int
	ByteOffset int
	Rune       rune
	CodePage   uint16
}

func (e *SourceEncodingError) Error() string {
	if e == nil {
		return "pack: source encoding is invalid"
	}
	return fmt.Sprintf("%s:%d:%d: rune U+%04X cannot be encoded with codepage %d", e.Path, e.Line, e.Column, e.Rune, e.CodePage)
}

// Details returns stable fields suitable for a CLI error envelope.
func (e *SourceEncodingError) Details() map[string]any {
	if e == nil {
		return nil
	}
	return map[string]any{
		"path":        e.Path,
		"field":       e.Field,
		"line":        e.Line,
		"column":      e.Column,
		"byte_column": e.ByteColumn,
		"byte_offset": e.ByteOffset,
		"rune":        string(e.Rune),
		"rune_code":   fmt.Sprintf("U+%04X", e.Rune),
		"codepage":    e.CodePage,
	}
}

// SourceEncodingErrors retains independent source encoding failures found
// while planning a pack operation.
type SourceEncodingErrors struct {
	Errors []*SourceEncodingError
}

func (e *SourceEncodingErrors) Error() string {
	if e == nil || len(e.Errors) == 0 {
		return "pack: source encoding is invalid"
	}
	if len(e.Errors) == 1 {
		return e.Errors[0].Error()
	}
	return fmt.Sprintf("pack: %d source encoding issue(s)", len(e.Errors))
}

func (e *SourceEncodingErrors) Unwrap() []error {
	if e == nil {
		return nil
	}
	issues := make([]error, 0, len(e.Errors))
	for _, issue := range e.Errors {
		if issue != nil {
			issues = append(issues, issue)
		}
	}
	return issues
}

// DiagnosticsError combines independent validation and compiler failures.
// Each underlying error remains available through errors.AsType.
type DiagnosticsError struct {
	Errors []error
}

func (e *DiagnosticsError) Error() string {
	if e == nil || len(e.Errors) == 0 {
		return "pack: validation failed"
	}
	return fmt.Sprintf("pack: validation found %d issue(s)", len(e.Errors))
}

func (e *DiagnosticsError) Unwrap() []error {
	if e == nil {
		return nil
	}
	return e.Errors
}
