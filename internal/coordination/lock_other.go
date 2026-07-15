//go:build !windows

package coordination

import (
	"errors"
	"os"
)

var errPlatformLockUnsupported = errors.New("workbook coordination locks are supported only on Windows")

func platformTryLock(_ *os.File, _ int64) (bool, error) {
	return false, errPlatformLockUnsupported
}

func platformUnlock(_ *os.File, _ int64) error {
	return errPlatformLockUnsupported
}

func platformAtomicReplace(source, destination string) error {
	return os.Rename(source, destination)
}
