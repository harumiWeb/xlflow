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

func platformProbePublishTarget(target string) error {
	path, err := windows.UTF16PtrFromString(target)
	if err != nil {
		return err
	}
	handle, err := windows.CreateFile(
		path,
		windows.GENERIC_READ|windows.GENERIC_WRITE,
		0,
		nil,
		windows.OPEN_EXISTING,
		windows.FILE_ATTRIBUTE_NORMAL,
		0,
	)
	if err != nil {
		return err
	}
	return windows.CloseHandle(handle)
}

func platformBusyError(err error) bool {
	return errors.Is(err, windows.ERROR_SHARING_VIOLATION) ||
		errors.Is(err, windows.ERROR_ACCESS_DENIED) ||
		errors.Is(err, windows.ERROR_LOCK_VIOLATION)
}
