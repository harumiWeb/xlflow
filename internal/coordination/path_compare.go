package coordination

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// SamePath compares two host paths using the path-equivalence behavior that
// predates workbook lock identities. It remains suitable for CLI file/session
// matching where the caller has no explicit project base directory.
func SamePath(left, right string) bool {
	leftCanonical, leftErr := canonicalHostPath(left)
	rightCanonical, rightErr := canonicalHostPath(right)
	if leftErr == nil && rightErr == nil {
		return strings.EqualFold(leftCanonical, rightCanonical)
	}
	if leftInfo, err := os.Stat(left); err == nil {
		if rightInfo, err := os.Stat(right); err == nil {
			return os.SameFile(leftInfo, rightInfo)
		}
	}
	return strings.EqualFold(filepath.Clean(left), filepath.Clean(right))
}

func canonicalHostPath(path string) (string, error) {
	if strings.TrimSpace(path) == "" {
		return "", fmt.Errorf("path is required")
	}
	absPath, err := filepath.Abs(path)
	if err != nil {
		return "", err
	}
	resolvedPath, err := filepath.EvalSymlinks(absPath)
	if err == nil {
		return filepath.Clean(resolvedPath), nil
	}
	if errors.Is(err, os.ErrNotExist) {
		return filepath.Clean(absPath), nil
	}
	return "", err
}
