// Package sourcepath defines host path keys for source artifact roles.
package sourcepath

import (
	"path/filepath"
	"runtime"
	"strings"
)

// Key cleans a logical host path, folding case only on Windows. VBA component
// names remain case-insensitive independently of filesystem artifact identity.
func Key(path string) string {
	path = filepath.Clean(path)
	if runtime.GOOS == "windows" {
		return strings.ToLower(path)
	}
	return path
}
