//go:build windows

package oracle

import (
	"context"
	"errors"
	"fmt"
	"os"
	"path/filepath"

	"golang.org/x/sys/windows"
)

type fileBatchLock struct {
	path string
}

func newBatchLock() (BatchLock, error) {
	cacheDir, err := os.UserCacheDir()
	if err != nil {
		return nil, fmt.Errorf("locate user cache directory for VBE oracle lock: %w", err)
	}
	stateDir := filepath.Join(cacheDir, "xlflow")
	if err := os.MkdirAll(stateDir, 0o700); err != nil {
		return nil, fmt.Errorf("create VBE oracle lock directory: %w", err)
	}
	return fileBatchLock{path: filepath.Join(stateDir, "vbe-oracle.lock")}, nil
}

func (l fileBatchLock) Acquire(ctx context.Context) (func(), error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}

	file, err := os.OpenFile(l.path, os.O_CREATE|os.O_RDWR, 0o600)
	if err != nil {
		return nil, fmt.Errorf("open VBE oracle lock: %w", err)
	}
	overlapped := windows.Overlapped{}
	err = windows.LockFileEx(
		windows.Handle(file.Fd()),
		windows.LOCKFILE_EXCLUSIVE_LOCK|windows.LOCKFILE_FAIL_IMMEDIATELY,
		0,
		1,
		0,
		&overlapped,
	)
	if err != nil {
		_ = file.Close()
		if errors.Is(err, windows.ERROR_LOCK_VIOLATION) {
			return nil, errOracleAlreadyRunning
		}
		return nil, fmt.Errorf("acquire VBE oracle lock: %w", err)
	}

	return func() {
		_ = windows.UnlockFileEx(windows.Handle(file.Fd()), 0, 1, 0, &overlapped)
		_ = file.Close()
	}, nil
}
