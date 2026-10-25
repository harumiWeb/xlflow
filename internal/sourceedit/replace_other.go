//go:build !unix && !windows

package sourceedit

import "fmt"

func platformAtomicReplace(_, _ string) error {
	return fmt.Errorf("%w: this operating system has no replacement implementation", ErrUnsupported)
}

func platformIsReparsePoint(string) (bool, error) {
	return false, nil
}
