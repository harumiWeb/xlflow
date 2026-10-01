//go:build unix

package coordination

import (
	"errors"
	"io"
	"os"

	"golang.org/x/sys/unix"
)

func platformTryLock(file *os.File, offset int64) (bool, error) {
	lock := unix.Flock_t{Type: unix.F_WRLCK, Whence: io.SeekStart, Start: offset, Len: 1}
	err := unix.FcntlFlock(file.Fd(), unix.F_SETLK, &lock)
	if err == nil {
		return true, nil
	}
	if errors.Is(err, unix.EACCES) || errors.Is(err, unix.EAGAIN) {
		return false, nil
	}
	return false, err
}

func platformUnlock(file *os.File, offset int64) error {
	lock := unix.Flock_t{Type: unix.F_UNLCK, Whence: io.SeekStart, Start: offset, Len: 1}
	return unix.FcntlFlock(file.Fd(), unix.F_SETLK, &lock)
}

func platformAtomicReplace(source, destination string) error {
	return os.Rename(source, destination)
}
