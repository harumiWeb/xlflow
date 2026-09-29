//go:build !windows

package coordination

import (
	"errors"
	"os"
	"syscall"
)

// platformAtomicCreate publishes through a hard link, which fails with EEXIST
// instead of overwriting a concurrent creator. The staged name is removed by
// the caller's cleanup step after the link is established.
func platformAtomicCreate(source, destination string) error {
	return os.Link(source, destination)
}

func platformBusyError(err error) bool {
	return errors.Is(err, syscall.EBUSY) ||
		errors.Is(err, syscall.EACCES) ||
		errors.Is(err, syscall.EPERM)
}
