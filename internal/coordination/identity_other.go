//go:build !windows

package coordination

import "path/filepath"

func normalizePlatformPath(path string) string {
	return filepath.Clean(path)
}

func platformComparisonKey(path string) string {
	return normalizePlatformPath(path)
}

func resolvePlatformPath(path string) (string, error) {
	return filepath.EvalSymlinks(path)
}
