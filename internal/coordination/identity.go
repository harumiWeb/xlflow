package coordination

import (
	"crypto/sha256"
	"encoding/hex"
	"fmt"
	"path/filepath"
	"strings"
)

const (
	workbookLockIDPrefix = "xlflow-workbook-v1-"
	workbookHashDomain   = "xlflow/workbook-coordination/v1\x00"
)

// WorkbookIdentity identifies a workbook for process-wide coordination.
// CanonicalPath is retained for diagnostics. LockID is the opaque value that
// should be used when naming an operating-system synchronization primitive.
type WorkbookIdentity struct {
	CanonicalPath string
	LockID        string
}

// NewWorkbookIdentity returns a stable coordination identity for workbookPath.
// baseDir must be absolute and is used to resolve relative workbook paths. The
// workbook does not need to exist.
func NewWorkbookIdentity(baseDir, workbookPath string) (WorkbookIdentity, error) {
	if strings.TrimSpace(baseDir) == "" {
		return WorkbookIdentity{}, fmt.Errorf("base directory is required")
	}
	if strings.TrimSpace(workbookPath) == "" {
		return WorkbookIdentity{}, fmt.Errorf("workbook path is required")
	}

	baseDir = normalizePlatformPath(baseDir)
	if !filepath.IsAbs(baseDir) {
		return WorkbookIdentity{}, fmt.Errorf("base directory must be absolute: %q", baseDir)
	}

	workbookPath = normalizePlatformPath(workbookPath)
	if !filepath.IsAbs(workbookPath) {
		workbookPath = filepath.Join(baseDir, workbookPath)
	}
	canonicalPath := normalizePlatformPath(filepath.Clean(workbookPath))

	// A real operating-system lock, rather than path metadata, is the eventual
	// source of truth. Symlink resolution is therefore best effort: a missing or
	// inaccessible workbook still receives a deterministic lexical identity.
	if resolved, err := resolvePlatformPath(canonicalPath); err == nil {
		canonicalPath = normalizePlatformPath(filepath.Clean(resolved))
	}

	comparisonKey := platformComparisonKey(canonicalPath)
	sum := sha256.Sum256([]byte(workbookHashDomain + comparisonKey))

	return WorkbookIdentity{
		CanonicalPath: canonicalPath,
		LockID:        workbookLockIDPrefix + hex.EncodeToString(sum[:]),
	}, nil
}
