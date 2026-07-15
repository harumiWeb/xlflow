//go:build windows

package coordination

import (
	"errors"
	"os"

	"golang.org/x/sys/windows"
)

const lockFlags = windows.LOCKFILE_EXCLUSIVE_LOCK | windows.LOCKFILE_FAIL_IMMEDIATELY

func platformTryLock(file *os.File, offset int64) (bool, error) {
	overlapped := windows.Overlapped{
		Offset:     uint32(offset),
		OffsetHigh: uint32(uint64(offset) >> 32),
	}
	err := windows.LockFileEx(windows.Handle(file.Fd()), lockFlags, 0, 1, 0, &overlapped)
	if err == nil {
		return true, nil
	}
	if errors.Is(err, windows.ERROR_LOCK_VIOLATION) || errors.Is(err, windows.ERROR_IO_PENDING) {
		return false, nil
	}
	return false, err
}

func platformUnlock(file *os.File, offset int64) error {
	overlapped := windows.Overlapped{
		Offset:     uint32(offset),
		OffsetHigh: uint32(uint64(offset) >> 32),
	}
	return windows.UnlockFileEx(windows.Handle(file.Fd()), 0, 1, 0, &overlapped)
}

func platformAtomicReplace(source, destination string) error {
	from, err := windows.UTF16PtrFromString(source)
	if err != nil {
		return err
	}
	to, err := windows.UTF16PtrFromString(destination)
	if err != nil {
		return err
	}
	return windows.MoveFileEx(from, to, windows.MOVEFILE_REPLACE_EXISTING|windows.MOVEFILE_WRITE_THROUGH)
}
