package pack

import "errors"

var (
	// ErrProtectedProject reports that the template contains a protected VBA project.
	ErrProtectedProject = errors.New("pack: protected VBA project")

	// ErrSignedProject reports that the template contains VBA signature streams or OOXML parts.
	ErrSignedProject = errors.New("pack: signed VBA project")

	// ErrAmbiguousLayout reports an unknown, unsupported, or ambiguous VBA project layout.
	ErrAmbiguousLayout = errors.New("pack: ambiguous VBA project layout")
)
