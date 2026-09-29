package coordination

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
)

var (
	// ErrPublishTargetBusy reports that the publish destination is locked by
	// another process or cannot be opened for replacement. The existing
	// destination is never modified when this error is returned.
	ErrPublishTargetBusy = errors.New("publish target is busy")

	// ErrPublishReplaceFailed reports that the staged temporary artifact could
	// not be atomically moved over the destination. The existing destination is
	// never modified when this error is returned.
	ErrPublishReplaceFailed = errors.New("could not atomically publish the staged artifact")
)

// CleanupResult describes the post-publish state of the staged temporary
// artifact. Status is "clean" when nothing remains to remove and "failed" when
// the staged file could not be removed after a successful publication.
type CleanupResult struct {
	Status       string
	ResidualPath string
	Error        string
}

// PublishResult describes one artifact publication.
type PublishResult struct {
	ReplacedExisting bool
	Publication      string // "atomic_create" or "atomic_replace"
	Cleanup          CleanupResult
}

// PublishFile atomically publishes data at target through a temporary sibling
// artifact in the same directory (same volume), mirroring the publication
// contract used by the Excel bridge build path:
//
//  1. probe the destination so a locked or busy file fails before staging
//  2. write, sync, and close a temporary sibling artifact
//  3. run validate on the closed temporary artifact when provided
//  4. publish with the platform atomic create/replace primitive
//  5. remove the temporary artifact on every failure path
//
// If any step fails before publication, an existing destination is left
// unchanged. There is no delete-then-copy fallback: when the atomic move is
// impossible the error wraps ErrPublishReplaceFailed (or ErrPublishTargetBusy
// when the destination is in use).
func PublishFile(target string, data []byte, validate func(path string) error) (PublishResult, error) {
	target = filepath.Clean(target)
	dir := filepath.Dir(target)
	if dir == "" {
		dir = "."
	}

	exists, err := statExists(target)
	if err != nil {
		return PublishResult{}, fmt.Errorf("stat publish target: %w", err)
	}
	if err := probePublishTarget(target, exists); err != nil {
		return PublishResult{}, err
	}

	tmp, err := os.CreateTemp(dir, ".xlflow-publish-*.tmp")
	if err != nil {
		return PublishResult{}, fmt.Errorf("create temporary artifact: %w", err)
	}
	tmpPath := tmp.Name()
	// On success the temporary artifact is consumed by the atomic move. The
	// deferred remove is a no-op for the platform primitives that rename the
	// file away, and covers every failure before that point.
	defer func() { _ = os.Remove(tmpPath) }()

	if _, err := tmp.Write(data); err != nil {
		_ = tmp.Close()
		return PublishResult{}, fmt.Errorf("write temporary artifact: %w", err)
	}
	if err := tmp.Sync(); err != nil {
		_ = tmp.Close()
		return PublishResult{}, fmt.Errorf("flush temporary artifact: %w", err)
	}
	if err := tmp.Close(); err != nil {
		return PublishResult{}, fmt.Errorf("close temporary artifact: %w", err)
	}

	if validate != nil {
		if err := validate(tmpPath); err != nil {
			return PublishResult{}, fmt.Errorf("validate temporary artifact: %w", err)
		}
	}

	if !exists {
		// The temporary artifact shares the destination directory, so this is a
		// same-volume atomic create. A concurrent creator is a publication
		// failure, never a fallback to overwrite.
		if err := platformAtomicCreate(tmpPath, target); err != nil {
			return PublishResult{}, classifyPublishMoveError(err)
		}
		return PublishResult{
			ReplacedExisting: false,
			Publication:      "atomic_create",
			Cleanup:          cleanupTemporaryArtifact(tmpPath),
		}, nil
	}

	if err := platformAtomicReplace(tmpPath, target); err != nil {
		return PublishResult{}, classifyPublishMoveError(err)
	}
	return PublishResult{
		ReplacedExisting: true,
		Publication:      "atomic_replace",
		Cleanup:          cleanupTemporaryArtifact(tmpPath),
	}, nil
}

// SameFileIdentity reports whether two host paths identify the same
// filesystem object after canonicalizing each with the workbook-identity
// model: nearest-existing-ancestor resolution, symlink/junction resolution,
// and platform normalization (drive-letter case, short names, UNC/extended
// prefixes). Paths that do not exist yet still resolve through their nearest
// existing ancestor, so aliased directories cannot hide the comparison.
func SameFileIdentity(baseDir, left, right string) (bool, error) {
	leftIdentity, err := NewWorkbookIdentity(baseDir, left)
	if err != nil {
		return false, fmt.Errorf("resolve %q: %w", left, err)
	}
	rightIdentity, err := NewWorkbookIdentity(baseDir, right)
	if err != nil {
		return false, fmt.Errorf("resolve %q: %w", right, err)
	}
	return platformComparisonKey(leftIdentity.CanonicalPath) ==
		platformComparisonKey(rightIdentity.CanonicalPath), nil
}

func statExists(path string) (bool, error) {
	info, err := os.Stat(path)
	if err == nil {
		return !info.IsDir(), nil
	}
	if errors.Is(err, os.ErrNotExist) {
		return false, nil
	}
	return false, err
}

// probePublishTarget opens an existing destination exclusively to fail fast
// when another process prevents replacement, such as an open workbook. The
// platform atomic move performs the same enforcement at publish time; the
// probe only avoids staging work that cannot be published.
func probePublishTarget(target string, exists bool) error {
	if !exists {
		return nil
	}
	file, err := os.OpenFile(target, os.O_RDWR, 0)
	if err != nil {
		if platformBusyError(err) {
			return fmt.Errorf("%w: %w", ErrPublishTargetBusy, err)
		}
		return fmt.Errorf("%w: open existing output: %w", ErrPublishReplaceFailed, err)
	}
	return file.Close()
}

func classifyPublishMoveError(err error) error {
	if platformBusyError(err) {
		return fmt.Errorf("%w: %w", ErrPublishTargetBusy, err)
	}
	return fmt.Errorf("%w: %w", ErrPublishReplaceFailed, err)
}

func cleanupTemporaryArtifact(tmpPath string) CleanupResult {
	if err := os.Remove(tmpPath); err != nil {
		if errors.Is(err, os.ErrNotExist) {
			return CleanupResult{Status: "clean"}
		}
		return CleanupResult{Status: "failed", ResidualPath: tmpPath, Error: err.Error()}
	}
	return CleanupResult{Status: "clean"}
}
