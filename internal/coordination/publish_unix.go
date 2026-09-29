//go:build !windows

package coordination

import (
	"errors"
	"syscall"
)

// platformProbePublishTarget is intentionally a no-op on Unix. Atomic rename
// is controlled by directory permissions and may validly replace a read-only
// file; the final rename is the authoritative publication check.
func platformProbePublishTarget(string) error {
	return nil
}

func platformBusyError(err error) bool {
	return errors.Is(err, syscall.EBUSY)
}
