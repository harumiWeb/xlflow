package pack

import "errors"

var (
	// ErrProtectedProject reports that the template contains a protected VBA project.
	ErrProtectedProject = errors.New("pack: protected VBA project")

	// ErrSignedProject reports that the template contains VBA signature streams or OOXML parts.
	ErrSignedProject = errors.New("pack: signed VBA project")

	// ErrUserFormGenerationUnsupported reports that the requested source update would require UserForm or .frx generation.
	ErrUserFormGenerationUnsupported = errors.New("pack: UserForm generation unsupported")

	// ErrBlankUserFormUnsupported reports that blank mode cannot author a UserForm designer.
	ErrBlankUserFormUnsupported = errors.New("pack: blank mode does not support UserForms")

	// ErrAmbiguousLayout reports an unknown, unsupported, or ambiguous VBA project layout.
	ErrAmbiguousLayout = errors.New("pack: ambiguous VBA project layout")
)
