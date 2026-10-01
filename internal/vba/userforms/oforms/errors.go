package oforms

import (
	"errors"
	"fmt"
)

// ErrMalformed identifies structurally inconsistent MS-OFORMS data.
var ErrMalformed = errors.New("oforms: malformed designer data")

// ParseError reports the exact storage, stream, and byte offset at which a
// designer structure failed validation.
type ParseError struct {
	Path      string
	Stream    string
	Offset    int
	Structure string
	Err       error
}

func (e *ParseError) Error() string {
	location := e.Path
	if e.Stream != "" {
		location += "/" + e.Stream
	}
	return fmt.Sprintf("oforms: %s at byte %d: %s: %v", location, e.Offset, e.Structure, e.Err)
}

func (e *ParseError) Unwrap() error { return errors.Join(ErrMalformed, e.Err) }

func parseError(path, stream string, offset int, structure, format string, args ...any) error {
	return &ParseError{
		Path:      path,
		Stream:    stream,
		Offset:    offset,
		Structure: structure,
		Err:       fmt.Errorf(format, args...),
	}
}
