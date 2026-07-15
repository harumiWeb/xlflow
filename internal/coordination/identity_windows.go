//go:build windows

package coordination

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"golang.org/x/sys/windows"
)

func normalizePlatformPath(path string) string {
	path = stripWindowsExtendedPrefix(path)
	path = filepath.Clean(path)
	path = strings.ReplaceAll(path, "/", `\`)

	volume := filepath.VolumeName(path)
	if len(volume) == 2 && volume[1] == ':' {
		path = strings.ToUpper(volume[:1]) + path[1:]
	}
	return path
}

func platformComparisonKey(path string) string {
	// Windows path matching is case-insensitive for xlflow's coordination
	// contract. Upper-casing also normalizes drive-letter casing.
	return strings.ToUpper(normalizePlatformPath(path))
}

func resolvePlatformPath(path string) (string, error) {
	// EvalSymlinks handles ordinary symbolic links, but can leave a junction
	// component unresolved on Windows. Asking the filesystem for the final DOS
	// path through an open handle covers both forms.
	file, err := os.Open(path)
	if err != nil {
		return "", err
	}
	defer func() { _ = file.Close() }()

	buffer := make([]uint16, 512)
	for {
		length, err := windows.GetFinalPathNameByHandle(windows.Handle(file.Fd()), &buffer[0], uint32(len(buffer)), 0)
		if err != nil {
			return "", err
		}
		if length == 0 {
			return "", fmt.Errorf("GetFinalPathNameByHandle returned an empty path")
		}
		if length < uint32(len(buffer)) {
			return windows.UTF16ToString(buffer[:length]), nil
		}
		buffer = make([]uint16, length+1)
	}
}

func stripWindowsExtendedPrefix(path string) string {
	path = strings.ReplaceAll(path, "/", `\`)
	if len(path) < 4 || !strings.EqualFold(path[:4], `\\?\`) {
		return path
	}

	rest := path[4:]
	if len(rest) >= 4 && strings.EqualFold(rest[:4], `UNC\`) {
		return `\\` + rest[4:]
	}
	return rest
}
