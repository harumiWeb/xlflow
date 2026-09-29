//go:build linux

package coordination

import (
	"errors"
	"os"
	"syscall"

	"golang.org/x/sys/unix"
)

// platformAtomicCreate uses Linux's no-replace rename so publication remains
// atomic on filesystems that support rename but reject hard links. Older
// kernels and filesystems may reject the flag; retain the hard-link strategy
// as a no-clobber fallback for those cases.
func platformAtomicCreate(source, destination string) error {
	err := unix.Renameat2(unix.AT_FDCWD, source, unix.AT_FDCWD, destination, unix.RENAME_NOREPLACE)
	if err == nil {
		return nil
	}
	if errors.Is(err, syscall.ENOSYS) || errors.Is(err, syscall.EINVAL) {
		return os.Link(source, destination)
	}
	return err
}
