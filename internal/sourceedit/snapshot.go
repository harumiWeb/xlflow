// Package sourceedit provides best-effort stale-write protection and atomic
// publication for edits to one existing regular file.
//
// The identity, mode, and content checks are cooperative protection against
// stale edits. Another process can still race the final check and publication,
// and this package does not promise crash durability.
package sourceedit

import (
	"bytes"
	"context"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"sync"
)

var (
	// ErrStale indicates that the captured file or its parent path changed.
	ErrStale = errors.New("sourceedit: source changed since capture")
	// ErrUnsupported indicates that atomic replacement is unavailable on this platform.
	ErrUnsupported = errors.New("sourceedit: atomic replacement is unsupported")
	// ErrReadOnly indicates that the captured source has no write permission bits.
	ErrReadOnly = errors.New("sourceedit: source is read-only")
)

var errSymlinkPath = errors.New("sourceedit: path contains a symlink or reparse point")

// Snapshot holds an absolute source path and the bytes read by Capture.
// Source is a convenience copy; modifying it does not change the private bytes
// used to validate a later Apply.
type Snapshot struct {
	// Path is the absolute path resolved by Capture. Apply continues to use that
	// captured path even if this display field is later changed.
	Path string
	// Source is a copy of the captured file contents.
	Source []byte

	mu             sync.Mutex
	path           string
	original       []byte
	identity       os.FileInfo
	parentIdentity os.FileInfo
	mode           os.FileMode
	ops            fileOperations
}

type fileOperations struct {
	beforePublish func()
	publish       func(tempPath, targetPath string) error
}

// Capture opens path as an existing regular file and records its absolute
// path, identity, permission bits, and contents. The target and every existing
// path component must not be a symlink or reparse point.
func Capture(ctx context.Context, path string) (*Snapshot, error) {
	if ctx == nil {
		return nil, errors.New("sourceedit: nil context")
	}
	if err := ctx.Err(); err != nil {
		return nil, fmt.Errorf("sourceedit: capture canceled: %w", err)
	}
	if path == "" {
		return nil, errors.New("sourceedit: path is empty")
	}

	absPath, err := filepath.Abs(path)
	if err != nil {
		return nil, fmt.Errorf("sourceedit: resolve path: %w", err)
	}
	absPath = filepath.Clean(absPath)
	pathInfo, err := lstatNoLinks(absPath)
	if err != nil {
		return nil, fmt.Errorf("sourceedit: inspect target: %w", err)
	}
	if !pathInfo.Mode().IsRegular() {
		return nil, fmt.Errorf("sourceedit: target is not a regular file: %s", absPath)
	}

	parentPath := filepath.Dir(absPath)
	parentInfo, err := lstatNoLinks(parentPath)
	if err != nil {
		return nil, fmt.Errorf("sourceedit: inspect parent: %w", err)
	}
	if !parentInfo.IsDir() {
		return nil, fmt.Errorf("sourceedit: parent is not a directory: %s", parentPath)
	}

	file, err := os.Open(absPath)
	if err != nil {
		return nil, fmt.Errorf("sourceedit: open target: %w", err)
	}
	openedInfo, statErr := file.Stat()
	if statErr != nil {
		_ = file.Close()
		return nil, fmt.Errorf("sourceedit: stat opened target: %w", statErr)
	}
	if !openedInfo.Mode().IsRegular() || !os.SameFile(pathInfo, openedInfo) {
		_ = file.Close()
		return nil, fmt.Errorf("sourceedit: target identity changed while opening: %s", absPath)
	}

	source, readErr := io.ReadAll(file)
	afterInfo, afterStatErr := file.Stat()
	closeErr := file.Close()
	if err := errors.Join(readErr, afterStatErr, closeErr); err != nil {
		return nil, fmt.Errorf("sourceedit: read target: %w", err)
	}
	if !os.SameFile(openedInfo, afterInfo) || !sameCaptureMetadata(openedInfo, afterInfo) || int64(len(source)) != afterInfo.Size() {
		return nil, fmt.Errorf("%w: source changed while being captured: %s", ErrStale, absPath)
	}

	pathAfter, err := lstatNoLinks(absPath)
	if err != nil {
		return nil, fmt.Errorf("sourceedit: verify target path: %w", err)
	}
	parentAfter, err := lstatNoLinks(parentPath)
	if err != nil {
		return nil, fmt.Errorf("sourceedit: verify parent path: %w", err)
	}
	if !os.SameFile(openedInfo, pathAfter) || !os.SameFile(parentInfo, parentAfter) {
		return nil, fmt.Errorf("%w: path identity changed while being captured: %s", ErrStale, absPath)
	}
	if err := ctx.Err(); err != nil {
		return nil, fmt.Errorf("sourceedit: capture canceled: %w", err)
	}

	return &Snapshot{
		Path:           absPath,
		Source:         bytes.Clone(source),
		path:           absPath,
		original:       bytes.Clone(source),
		identity:       openedInfo,
		parentIdentity: parentInfo,
		mode:           openedInfo.Mode().Perm(),
	}, nil
}

// Apply atomically replaces the captured file with replacement after
// rechecking its identity, permission bits, contents, and parent path. A
// byte-for-byte no-op does not create or publish a temporary file.
func (s *Snapshot) Apply(ctx context.Context, replacement []byte) (resultErr error) {
	if s == nil {
		return errors.New("sourceedit: nil snapshot")
	}
	if ctx == nil {
		return errors.New("sourceedit: nil context")
	}
	s.mu.Lock()
	defer s.mu.Unlock()

	if err := ctx.Err(); err != nil {
		return fmt.Errorf("sourceedit: apply canceled: %w", err)
	}
	if s.mode&0o222 == 0 {
		return fmt.Errorf("%w: %s", ErrReadOnly, s.path)
	}
	replacement = bytes.Clone(replacement)
	if err := s.verifyCurrent(); err != nil {
		return err
	}
	if bytes.Equal(replacement, s.original) {
		return nil
	}

	temp, err := os.CreateTemp(filepath.Dir(s.path), ".sourceedit-*.tmp")
	if err != nil {
		return fmt.Errorf("sourceedit: create temporary file: %w", err)
	}
	tempPath := temp.Name()
	defer func() {
		if err := os.Remove(tempPath); err != nil && !errors.Is(err, os.ErrNotExist) {
			resultErr = errors.Join(resultErr, fmt.Errorf("sourceedit: remove temporary file: %w", err))
		}
	}()

	written, writeErr := temp.Write(replacement)
	if writeErr == nil && written != len(replacement) {
		writeErr = io.ErrShortWrite
	}
	if writeErr != nil {
		_ = temp.Close()
		return fmt.Errorf("sourceedit: write temporary file: %w", writeErr)
	}
	if err := temp.Chmod(s.mode); err != nil {
		_ = temp.Close()
		return fmt.Errorf("sourceedit: preserve source permissions: %w", err)
	}
	if err := temp.Sync(); err != nil {
		_ = temp.Close()
		return fmt.Errorf("sourceedit: sync temporary file: %w", err)
	}
	if err := temp.Close(); err != nil {
		return fmt.Errorf("sourceedit: close temporary file: %w", err)
	}

	if s.ops.beforePublish != nil {
		s.ops.beforePublish()
	}
	if err := ctx.Err(); err != nil {
		return fmt.Errorf("sourceedit: apply canceled before publication: %w", err)
	}
	if err := s.verifyCurrent(); err != nil {
		return err
	}
	if err := ctx.Err(); err != nil {
		return fmt.Errorf("sourceedit: apply canceled before publication: %w", err)
	}

	publish := s.ops.publish
	if publish == nil {
		publish = platformAtomicReplace
	}
	if err := publish(tempPath, s.path); err != nil {
		return fmt.Errorf("sourceedit: publish replacement: %w", err)
	}
	return nil
}

func (s *Snapshot) verifyCurrent() error {
	parentPath := filepath.Dir(s.path)
	parentInfo, err := lstatNoLinks(parentPath)
	if err != nil {
		return staleOrInspectionError(s.path, err)
	}
	if !parentInfo.IsDir() || !os.SameFile(s.parentIdentity, parentInfo) {
		return fmt.Errorf("%w: parent directory identity changed: %s", ErrStale, parentPath)
	}

	pathInfo, err := lstatNoLinks(s.path)
	if err != nil {
		return staleOrInspectionError(s.path, err)
	}
	if !pathInfo.Mode().IsRegular() || !os.SameFile(s.identity, pathInfo) {
		return fmt.Errorf("%w: target identity changed: %s", ErrStale, s.path)
	}
	if pathInfo.Mode().Perm() != s.mode {
		return fmt.Errorf("%w: source permission bits changed: %s", ErrStale, s.path)
	}

	file, err := os.Open(s.path)
	if err != nil {
		return staleOrInspectionError(s.path, err)
	}
	openedInfo, statErr := file.Stat()
	if statErr != nil {
		_ = file.Close()
		return fmt.Errorf("sourceedit: stat current source: %w", statErr)
	}
	if !openedInfo.Mode().IsRegular() || !os.SameFile(s.identity, openedInfo) || openedInfo.Mode().Perm() != s.mode {
		_ = file.Close()
		return fmt.Errorf("%w: opened target identity or mode changed: %s", ErrStale, s.path)
	}

	current, readErr := io.ReadAll(file)
	afterInfo, afterStatErr := file.Stat()
	closeErr := file.Close()
	if err := errors.Join(readErr, afterStatErr, closeErr); err != nil {
		return fmt.Errorf("sourceedit: read current source: %w", err)
	}
	if !os.SameFile(s.identity, afterInfo) || afterInfo.Mode().Perm() != s.mode {
		return fmt.Errorf("%w: source identity or mode changed while reading: %s", ErrStale, s.path)
	}
	if !bytes.Equal(current, s.original) {
		return fmt.Errorf("%w: source contents changed: %s", ErrStale, s.path)
	}

	pathAfter, err := lstatNoLinks(s.path)
	if err != nil {
		return staleOrInspectionError(s.path, err)
	}
	parentAfter, err := lstatNoLinks(parentPath)
	if err != nil {
		return staleOrInspectionError(parentPath, err)
	}
	if !os.SameFile(s.identity, pathAfter) || pathAfter.Mode().Perm() != s.mode ||
		!os.SameFile(s.parentIdentity, parentAfter) {
		return fmt.Errorf("%w: target or parent path changed during verification: %s", ErrStale, s.path)
	}
	return nil
}

func lstatNoLinks(path string) (os.FileInfo, error) {
	volume := filepath.VolumeName(path)
	root := volume + string(os.PathSeparator)
	if volume == "" {
		root = string(os.PathSeparator)
	}
	remainder := strings.TrimLeft(strings.TrimPrefix(path, volume), `/\`)
	components := strings.FieldsFunc(remainder, func(r rune) bool { return r == '/' || r == '\\' })
	current := root
	info, err := os.Lstat(root)
	if err != nil {
		return nil, err
	}
	if info.Mode()&os.ModeSymlink != 0 {
		return nil, fmt.Errorf("%w: %s", errSymlinkPath, root)
	}
	reparsePoint, err := platformIsReparsePoint(root)
	if err != nil {
		return nil, fmt.Errorf("inspect path component %s: %w", root, err)
	}
	if reparsePoint {
		return nil, fmt.Errorf("%w: %s", errSymlinkPath, root)
	}
	for _, component := range components {
		current = filepath.Join(current, component)
		info, err = os.Lstat(current)
		if err != nil {
			return nil, err
		}
		if info.Mode()&os.ModeSymlink != 0 {
			return nil, fmt.Errorf("%w: %s", errSymlinkPath, current)
		}
		reparsePoint, err := platformIsReparsePoint(current)
		if err != nil {
			return nil, fmt.Errorf("inspect path component %s: %w", current, err)
		}
		if reparsePoint {
			return nil, fmt.Errorf("%w: %s", errSymlinkPath, current)
		}
	}
	return info, nil
}

func sameCaptureMetadata(before, after os.FileInfo) bool {
	return before.Size() == after.Size() && before.Mode().Perm() == after.Mode().Perm() && before.ModTime().Equal(after.ModTime())
}

func staleOrInspectionError(path string, err error) error {
	if errors.Is(err, os.ErrNotExist) || errors.Is(err, errSymlinkPath) {
		return fmt.Errorf("%w: cannot safely inspect %s: %v", ErrStale, path, err)
	}
	return fmt.Errorf("sourceedit: inspect %s: %w", path, err)
}
