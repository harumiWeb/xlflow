//go:build unix

package coordination

import (
	"errors"
	"fmt"
	"os"
	"sync"

	"golang.org/x/sys/unix"
)

type unixLockKey struct {
	fd     uintptr
	offset int64
}

var unixLocks = struct {
	sync.Mutex
	files map[unixLockKey]*os.File
}{files: map[unixLockKey]*os.File{}}

func platformTryLock(file *os.File, offset int64, shared bool) (bool, error) {
	key := unixLockKey{fd: file.Fd(), offset: offset}
	unixLocks.Lock()
	defer unixLocks.Unlock()
	if _, exists := unixLocks.files[key]; exists {
		return true, nil
	}

	lockFile, err := os.OpenFile(fmt.Sprintf("%s.%d", file.Name(), offset), os.O_CREATE|os.O_RDWR, 0o600)
	if err != nil {
		return false, err
	}
	mode := unix.LOCK_EX
	if shared {
		mode = unix.LOCK_SH
	}
	err = unix.Flock(int(lockFile.Fd()), mode|unix.LOCK_NB)
	if err == nil {
		unixLocks.files[key] = lockFile
		return true, nil
	}
	_ = lockFile.Close()
	if errors.Is(err, unix.EACCES) || errors.Is(err, unix.EAGAIN) {
		return false, nil
	}
	return false, err
}

func platformUnlock(file *os.File, offset int64) error {
	key := unixLockKey{fd: file.Fd(), offset: offset}
	unixLocks.Lock()
	defer unixLocks.Unlock()
	lockFile, exists := unixLocks.files[key]
	if !exists {
		return fmt.Errorf("unix lock is not held for offset %d", offset)
	}
	if err := unix.Flock(int(lockFile.Fd()), unix.LOCK_UN); err != nil {
		return err
	}
	delete(unixLocks.files, key)
	return lockFile.Close()
}

func platformAtomicReplace(source, destination string) error {
	return os.Rename(source, destination)
}
