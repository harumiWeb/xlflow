//go:build unix

package sourceedit

import "os"

func platformAtomicReplace(tempPath, targetPath string) error {
	return os.Rename(tempPath, targetPath)
}
