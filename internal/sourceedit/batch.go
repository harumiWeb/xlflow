package sourceedit

import (
	"bytes"
	"cmp"
	"context"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"slices"

	"github.com/harumiWeb/xlflow/internal/sourcepath"
)

// Change binds proposed bytes to a previously captured source revision.
type Change struct {
	Snapshot    *Snapshot
	Replacement []byte
}

// BatchFileResult describes writes made by this invocation, not the current
// contents of paths that another process may have edited independently.
type BatchFileResult struct {
	Path         string
	State        string
	RecoveryPath string
}

// BatchResult remains meaningful on failure. Committed means all proposed
// changes were published; a later cleanup error does not undo that fact.
type BatchResult struct {
	Committed bool
	Files     []BatchFileResult
	Artifacts []string
}

// BatchError separates a publication/preparation failure from recovery and
// cleanup failures. Recovery ignores cancellation to restore earlier writes.
type BatchError struct {
	Cause    error
	Rollback error
	Cleanup  error
}

func (e *BatchError) Error() string {
	return fmt.Sprintf("sourceedit: batch failed: %v", errors.Join(e.Cause, e.Rollback, e.Cleanup))
}

func (e *BatchError) Unwrap() []error {
	var children []error
	for _, err := range []error{e.Cause, e.Rollback, e.Cleanup} {
		if err != nil {
			children = append(children, err)
		}
	}
	return children
}

type incompleteStageError struct {
	path    string
	cause   error
	cleanup error
}

func (e *incompleteStageError) Error() string {
	return fmt.Sprintf("sourceedit: incomplete stage retained at %s: %v", e.path, errors.Join(e.cause, e.cleanup))
}

func (e *incompleteStageError) Unwrap() error { return errors.Join(e.cause, e.cleanup) }

type batchItem struct {
	source      *Snapshot
	replacement []byte
	staged      *Snapshot
	backup      *Snapshot
	published   *Snapshot
}

type batchOperations struct {
	stage   func(context.Context, *Snapshot, []byte) (*Snapshot, error)
	remove  func(string) error
	recover func([]byte) (*Snapshot, error)
}

// ApplyBatch stages all changed files and original-byte backups before any
// publication. Returned failures trigger reverse rollback, with retained
// backups for files that cannot safely be restored. Publication is atomic per
// file only; this is not crash recovery or a race-free multi-file transaction.
// Duplicate logical paths or physical file identities are rejected.
func ApplyBatch(ctx context.Context, changes []Change) (BatchResult, error) {
	return applyBatch(ctx, changes, batchOperations{})
}

func applyBatch(ctx context.Context, changes []Change, ops batchOperations) (result BatchResult, resultErr error) {
	result.Files = make([]BatchFileResult, len(changes))
	result.Artifacts = []string{}
	if ctx == nil {
		return result, &BatchError{Cause: errors.New("sourceedit: nil context")}
	}
	items := make([]batchItem, len(changes))
	for i, change := range changes {
		if change.Snapshot == nil || change.Snapshot.identity == nil {
			return result, &BatchError{Cause: errors.New("sourceedit: nil snapshot")}
		}
		items[i] = batchItem{source: change.Snapshot, replacement: bytes.Clone(change.Replacement)}
		result.Files[i] = BatchFileResult{Path: change.Snapshot.path, State: "not_applied"}
	}
	for i, item := range items {
		for j := range i {
			if sourcepath.Key(items[j].source.path) == sourcepath.Key(item.source.path) || os.SameFile(items[j].source.identity, item.source.identity) {
				return result, &BatchError{Cause: fmt.Errorf("sourceedit: duplicate source identity: %s", item.source.path)}
			}
		}
	}
	// Stable locking order avoids deadlock between batches sharing snapshots.
	locked := make([]*Snapshot, len(items))
	for i := range items {
		locked[i] = items[i].source
	}
	slices.SortFunc(locked, func(a, b *Snapshot) int { return cmp.Compare(sourcepath.Key(a.path), sourcepath.Key(b.path)) })
	for _, snapshot := range locked {
		snapshot.mu.Lock()
	}
	defer func() {
		for i := len(locked) - 1; i >= 0; i-- {
			locked[i].mu.Unlock()
		}
	}()
	if ops.stage == nil {
		ops.stage = stageBatchFile
	}
	if ops.remove == nil {
		ops.remove = os.Remove
	}
	if ops.recover == nil {
		ops.recover = func(data []byte) (*Snapshot, error) {
			return stageBatchBytes(context.Background(), os.TempDir(), 0o600, data)
		}
	}
	defer func() {
		cleanupErr := cleanupBatch(items, &result, ops.remove)
		if resultErr == nil && cleanupErr == nil {
			return
		}
		batchErr, ok := errors.AsType[*BatchError](resultErr)
		if !ok {
			batchErr = &BatchError{Cause: resultErr}
		}
		batchErr.Cleanup = errors.Join(batchErr.Cleanup, cleanupErr)
		if incomplete, ok := errors.AsType[*incompleteStageError](resultErr); ok {
			result.Artifacts = append(result.Artifacts, incomplete.path)
			batchErr.Cleanup = errors.Join(batchErr.Cleanup, incomplete.cleanup)
		}
		resultErr = batchErr
	}()
	if err := verifyBatch(ctx, items); err != nil {
		return result, &BatchError{Cause: err}
	}
	for i := range items {
		item := &items[i]
		if bytes.Equal(item.replacement, item.source.original) {
			result.Files[i].State = "unchanged"
			continue
		}
		if item.source.mode&0o222 == 0 {
			return result, &BatchError{Cause: fmt.Errorf("%w: %s", ErrReadOnly, item.source.path)}
		}
	}
	for i := range items {
		item := &items[i]
		if result.Files[i].State == "unchanged" {
			continue
		}
		var err error
		item.backup, err = ops.stage(ctx, item.source, item.source.original)
		if err != nil {
			return result, &BatchError{Cause: err}
		}
		item.staged, err = ops.stage(ctx, item.source, item.replacement)
		if err != nil {
			return result, &BatchError{Cause: err}
		}
	}
	// Include unchanged targets in the final pre-publication validation.
	if err := verifyBatch(ctx, items); err != nil {
		return result, &BatchError{Cause: err}
	}
	for i := range items {
		item := &items[i]
		if item.staged == nil {
			continue
		}
		err := publishBatchItem(ctx, item)
		if err != nil {
			rollbackErr := rollbackBatch(items, &result, ops.recover)
			return result, &BatchError{Cause: err, Rollback: rollbackErr}
		}
		result.Files[i].State = "applied"
	}
	// Publication of the last changed file is the commit point. Cancellation
	// afterwards must not report failure after successful writes.
	result.Committed = true
	return result, nil
}

func verifyBatch(ctx context.Context, items []batchItem) error {
	for _, item := range items {
		if err := ctx.Err(); err != nil {
			return err
		}
		if err := item.source.verifyCurrent(); err != nil {
			return err
		}
	}
	return ctx.Err()
}

func stageBatchFile(ctx context.Context, source *Snapshot, data []byte) (_ *Snapshot, resultErr error) {
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	if err := source.verifyCurrent(); err != nil {
		return nil, err
	}
	return stageBatchBytes(ctx, filepath.Dir(source.path), source.mode, data)
}

func stageBatchBytes(ctx context.Context, dir string, mode os.FileMode, data []byte) (_ *Snapshot, resultErr error) {
	temp, err := os.CreateTemp(dir, ".sourceedit-*.tmp")
	if err != nil {
		return nil, fmt.Errorf("sourceedit: stage in %s: %w", dir, err)
	}
	path := temp.Name()
	defer func() {
		if resultErr != nil {
			_ = temp.Close()
			if err := os.Remove(path); err != nil && !errors.Is(err, os.ErrNotExist) {
				resultErr = &incompleteStageError{path: path, cause: resultErr, cleanup: err}
			}
		}
	}()
	n, err := temp.Write(data)
	if err == nil && n != len(data) {
		err = io.ErrShortWrite
	}
	if err != nil {
		return nil, err
	}
	if err := temp.Chmod(mode); err != nil {
		return nil, err
	}
	if err := temp.Sync(); err != nil {
		return nil, err
	}
	if err := temp.Close(); err != nil {
		return nil, err
	}
	return Capture(ctx, path)
}

func publishBatchItem(ctx context.Context, item *batchItem) error {
	source := item.source
	if source.ops.beforePublish != nil {
		source.ops.beforePublish()
	}
	if err := ctx.Err(); err != nil {
		return err
	}
	for _, snapshot := range []*Snapshot{source, item.staged, item.backup} {
		if err := snapshot.verifyCurrent(); err != nil {
			return err
		}
	}
	if err := ctx.Err(); err != nil {
		return err
	}
	publish := source.ops.publish
	if publish == nil {
		publish = platformAtomicReplace
	}
	if err := publish(item.staged.path, source.path); err != nil {
		return fmt.Errorf("sourceedit: publish %s: %w", source.path, err)
	}
	item.published = &Snapshot{path: source.path, original: item.staged.original, identity: item.staged.identity, parentIdentity: source.parentIdentity, mode: source.mode}
	item.staged = nil // Its directory entry was consumed by atomic replacement.
	return nil
}

func rollbackBatch(items []batchItem, result *BatchResult, recoverOriginal func([]byte) (*Snapshot, error)) error {
	var failures []error
	for i := len(items) - 1; i >= 0; i-- {
		item := &items[i]
		if item.published == nil {
			continue
		}
		backupErr := item.backup.verifyCurrent()
		err := errors.Join(item.published.verifyCurrent(), backupErr)
		if err == nil {
			publish := item.source.ops.publish
			if publish == nil {
				publish = platformAtomicReplace
			}
			err = publish(item.backup.path, item.source.path)
		}
		if err != nil {
			// Restore itself can lose or alter the previously verified backup.
			// Select recovery material using its state after that failure.
			backupErr = item.backup.verifyCurrent()
			err = errors.Join(err, backupErr)
			result.Files[i].State = "recovery_required"
			result.Files[i].RecoveryPath = item.backup.path
			if backupErr != nil {
				result.Files[i].RecoveryPath = ""
				// A moved/replaced parent or altered backup makes its old path
				// unreliable. Persist private original bytes outside that parent
				// rather than claiming the old path is usable recovery material.
				recovery, recoveryErr := recoverOriginal(item.source.original)
				if recoveryErr != nil {
					err = errors.Join(err, fmt.Errorf("persist recovery copy: %w", recoveryErr))
				} else {
					result.Files[i].RecoveryPath = recovery.path
					result.Artifacts = append(result.Artifacts, recovery.path)
				}
			}
			failures = append(failures, fmt.Errorf("sourceedit: restore %s from %s: %w", item.source.path, item.backup.path, err))
			continue
		}
		item.backup = nil
		result.Files[i].State = "restored"
	}
	return errors.Join(failures...)
}

func cleanupBatch(items []batchItem, result *BatchResult, remove func(string) error) error {
	var failures []error
	for i, item := range items {
		for _, artifact := range []*Snapshot{item.staged, item.backup} {
			if artifact == nil {
				continue
			}
			if _, err := os.Lstat(artifact.path); errors.Is(err, os.ErrNotExist) {
				continue // It was consumed, moved away, or removed externally.
			}
			if result.Files[i].RecoveryPath == artifact.path {
				result.Artifacts = append(result.Artifacts, artifact.path)
				continue
			}
			err := artifact.verifyCurrent()
			if err == nil {
				err = remove(artifact.path)
			}
			if err != nil {
				result.Artifacts = append(result.Artifacts, artifact.path)
				failures = append(failures, fmt.Errorf("sourceedit: cleanup %s: %w", artifact.path, err))
			}
		}
	}
	return errors.Join(failures...)
}
