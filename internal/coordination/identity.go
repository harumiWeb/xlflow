package coordination

import (
	"crypto/sha256"
	"encoding/hex"
	"fmt"
	"path/filepath"
	"strings"
)

const (
	workbookLockIDPrefix   = "xlflow-workbook-v1-"
	workbookHashDomain     = "xlflow/workbook-coordination/v1\x00"
	sourceTreeLockIDPrefix = "xlflow-source-tree-v1-"
	sourceTreeHashDomain   = "xlflow/source-tree-coordination/v1\x00"
)

// ResourceIdentity identifies a filesystem resource for process-wide coordination.
// CanonicalPath is retained for diagnostics. LockID is the opaque value that
// should be used when naming an operating-system synchronization primitive.
type ResourceIdentity struct {
	CanonicalPath string
	LockID        string
}

// WorkbookIdentity retains the workbook-specific API name while coordination
// uses the same identity shape for other filesystem resources.
type WorkbookIdentity = ResourceIdentity

// NewWorkbookIdentity returns a stable coordination identity for workbookPath.
// baseDir must be absolute and is used to resolve relative workbook paths. The
// workbook does not need to exist.
func NewWorkbookIdentity(baseDir, workbookPath string) (WorkbookIdentity, error) {
	return newPathIdentity(baseDir, workbookPath, workbookLockIDPrefix, workbookHashDomain, "workbook")
}

// NewSourceTreeIdentity returns a stable identity for one managed source root.
func NewSourceTreeIdentity(baseDir, sourceRoot string) (ResourceIdentity, error) {
	return newPathIdentity(baseDir, sourceRoot, sourceTreeLockIDPrefix, sourceTreeHashDomain, "source root")
}

func newPathIdentity(baseDir, resourcePath, prefix, domain, label string) (ResourceIdentity, error) {
	if strings.TrimSpace(baseDir) == "" {
		return ResourceIdentity{}, fmt.Errorf("base directory is required")
	}
	if strings.TrimSpace(resourcePath) == "" {
		return ResourceIdentity{}, fmt.Errorf("%s path is required", label)
	}

	baseDir = normalizePlatformPath(baseDir)
	if !filepath.IsAbs(baseDir) {
		return ResourceIdentity{}, fmt.Errorf("base directory must be absolute: %q", baseDir)
	}

	resourcePath = normalizePlatformPath(resourcePath)
	if !filepath.IsAbs(resourcePath) {
		resourcePath = filepath.Join(baseDir, resourcePath)
	}
	canonicalPath := normalizePlatformPath(filepath.Clean(resourcePath))

	// A real operating-system lock, rather than path metadata, is the eventual
	// source of truth. Resolve the nearest existing ancestor so a workbook that
	// has not been created yet still shares an identity across symlinked or
	// junctioned project paths. If no ancestor can be resolved, retain the
	// deterministic lexical identity.
	if resolved, err := resolveNearestExistingAncestor(canonicalPath); err == nil {
		canonicalPath = normalizePlatformPath(filepath.Clean(resolved))
	}

	comparisonKey := platformComparisonKey(canonicalPath)
	sum := sha256.Sum256([]byte(domain + comparisonKey))

	return ResourceIdentity{
		CanonicalPath: canonicalPath,
		LockID:        prefix + hex.EncodeToString(sum[:]),
	}, nil
}

func resolveNearestExistingAncestor(path string) (string, error) {
	candidate := path
	missingTail := make([]string, 0, 4)
	var lastErr error
	for {
		resolved, err := resolvePlatformPath(candidate)
		if err == nil {
			for i := len(missingTail) - 1; i >= 0; i-- {
				resolved = filepath.Join(resolved, missingTail[i])
			}
			return resolved, nil
		}
		lastErr = err

		parent := filepath.Dir(candidate)
		if parent == candidate {
			return "", lastErr
		}
		missingTail = append(missingTail, filepath.Base(candidate))
		candidate = parent
	}
}
