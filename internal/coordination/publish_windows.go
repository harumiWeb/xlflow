//go:build windows

package coordination

import (
	"errors"

	"golang.org/x/sys/windows"
)

// platformAtomicCreate moves the staged file over the destination only when no
// destination exists yet; MoveFileEx without MOVEFILE_REPLACE_EXISTING fails
// instead of overwriting a concurrent creator.
func platformAtomicCreate(source, destination string) error {
	from, err := windows.UTF16PtrFromString(source)
	if err != nil {
		return err
	}
	to, err := windows.UTF16PtrFromString(destination)
	if err != nil {
		return err
	}
	return windows.MoveFileEx(from, to, windows.MOVEFILE_WRITE_THROUGH)
}

func platformBusyError(err error) bool {
	return errors.Is(err, windows.ERROR_SHARING_VIOLATION) ||
		errors.Is(err, windows.ERROR_ACCESS_DENIED) ||
		errors.Is(err, windows.ERROR_LOCK_VIOLATION)
}
