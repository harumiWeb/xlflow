package filepull

import (
	"bytes"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"slices"

	"github.com/harumiWeb/xlflow/internal/coordination"
)

var (
	publishArtifact = coordination.PublishFile
	removeArtifact  = os.Remove
)

type originalFile struct {
	path   string
	body   []byte
	mode   os.FileMode
	exists bool
}

func publish(p plan) error {
	paths := make([]string, 0, len(p.files)+len(p.stale))
	assetBodies := make(map[string][]byte)
	for _, file := range p.files {
		paths = append(paths, file.path)
		if file.preserveExistingAsset {
			assetBodies[filepath.Clean(file.path)] = file.body
		}
	}
	paths = append(paths, p.stale...)
	slices.Sort(paths)
	paths = slices.Compact(paths)
	originals := make([]originalFile, 0, len(paths))
	var createdDirs []string
	createdDirSet := map[string]bool{}
	for _, path := range paths {
		original, err := snapshot(path)
		if err != nil {
			return fmt.Errorf("%w: snapshot %s: %v", ErrPublish, path, err)
		}
		if body, asset := assetBodies[filepath.Clean(path)]; asset && original.exists && !bytes.Equal(original.body, body) {
			return fmt.Errorf("%w: existing path %s has different contents", ErrPictureAssetConflict, path)
		}
		originals = append(originals, original)
	}

	for _, file := range p.files {
		dirs, err := ensureParentDirectories(filepath.Dir(file.path))
		for _, dir := range dirs {
			if !createdDirSet[dir] {
				createdDirs = append(createdDirs, dir)
				createdDirSet[dir] = true
			}
		}
		if err != nil {
			return rollbackPublish(originals, createdDirs, fmt.Errorf("create source directory: %w", err))
		}
		if _, err := publishArtifact(file.path, file.body, nil); err != nil {
			return rollbackPublish(originals, createdDirs, err)
		}
	}
	for _, path := range p.stale {
		if err := removeArtifact(path); err != nil && !errors.Is(err, os.ErrNotExist) {
			return rollbackPublish(originals, createdDirs, err)
		}
	}
	return nil
}

func ensureParentDirectories(path string) ([]string, error) {
	var missing []string
	for current := filepath.Clean(path); ; current = filepath.Dir(current) {
		info, err := os.Stat(current)
		if err == nil {
			if !info.IsDir() {
				return missing, fmt.Errorf("source parent is not a directory: %s", current)
			}
			break
		}
		if !errors.Is(err, os.ErrNotExist) {
			return missing, err
		}
		missing = append(missing, current)
		parent := filepath.Dir(current)
		if parent == current {
			return missing, fmt.Errorf("no existing parent for %s", path)
		}
	}
	slices.Reverse(missing)
	if err := os.MkdirAll(path, 0o755); err != nil {
		return missing, err
	}
	return missing, nil
}

func snapshot(path string) (originalFile, error) {
	info, err := os.Lstat(path)
	if errors.Is(err, os.ErrNotExist) {
		return originalFile{path: path}, nil
	}
	if err != nil {
		return originalFile{}, err
	}
	if info.Mode()&os.ModeSymlink != 0 || !info.Mode().IsRegular() {
		return originalFile{}, fmt.Errorf("source target is not a regular file")
	}
	body, err := os.ReadFile(path)
	if err != nil {
		return originalFile{}, err
	}
	return originalFile{path: path, body: body, mode: info.Mode().Perm(), exists: true}, nil
}

func rollbackPublish(originals []originalFile, createdDirs []string, primary error) error {
	var rollbackErrors []error
	for i := len(originals) - 1; i >= 0; i-- {
		original := originals[i]
		if original.exists {
			if err := os.MkdirAll(filepath.Dir(original.path), 0o755); err != nil {
				rollbackErrors = append(rollbackErrors, err)
				continue
			}
			if _, err := publishArtifact(original.path, original.body, nil); err != nil {
				rollbackErrors = append(rollbackErrors, fmt.Errorf("restore %s: %w", original.path, err))
				continue
			}
			if err := os.Chmod(original.path, original.mode); err != nil {
				rollbackErrors = append(rollbackErrors, fmt.Errorf("restore mode %s: %w", original.path, err))
			}
			continue
		}
		if err := removeArtifact(original.path); err != nil && !errors.Is(err, os.ErrNotExist) {
			rollbackErrors = append(rollbackErrors, fmt.Errorf("remove new file %s: %w", original.path, err))
		}
	}
	for i := len(createdDirs) - 1; i >= 0; i-- {
		if err := os.Remove(createdDirs[i]); err != nil && !errors.Is(err, os.ErrNotExist) {
			rollbackErrors = append(rollbackErrors, fmt.Errorf("remove new directory %s: %w", createdDirs[i], err))
		}
	}
	return fmt.Errorf("%w: %w", ErrPublish, errors.Join(append([]error{primary}, rollbackErrors...)...))
}
