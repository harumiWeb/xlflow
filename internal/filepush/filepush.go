// Package filepush applies the tracked VBA source tree to a saved .xlsm
// workbook without starting Excel or using COM/VBIDE. It reuses the pure-Go
// pack engine to rebuild xl/vbaProject.bin, validates the staged artifact,
// and atomically replaces the workbook through coordination.PublishFile.
//
// The backend deliberately mirrors the Excel push contract where it matters:
// source fingerprinting, duplicate detection, folder annotation updates, Erl
// line-number instrumentation, backup creation, and .xlflow/state/push.json
// all use the same schema and semantics, so --changed-only interoperates with
// pushes performed through Excel. A successful file push means the artifact
// was reconstructed and structurally validated — not that the VBA compiled
// in VBE.
package filepush

import (
	"archive/zip"
	"context"
	"errors"
	"fmt"
	"io"
	"maps"
	"os"
	"path/filepath"
	"slices"
	"strings"
	"time"

	"github.com/harumiWeb/xlflow/internal/backup"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/coordination"
	packpkg "github.com/harumiWeb/xlflow/internal/pack"
)

// Options controls one file-backend push.
type Options struct {
	// BackupMode is "always" (default) or "never".
	BackupMode string
	// ChangedOnly skips the rebuild when the recorded push state still covers
	// the current source tree and saved workbook.
	ChangedOnly bool
	// Coordination is the lease manager for source-tree locks; nil creates the
	// default manager.
	Coordination *coordination.Manager
	// Wait/WaitTimeout mirror the CLI --wait contract for busy source trees.
	Wait        bool
	WaitTimeout time.Duration
	// Now overrides the backup timestamp (tests); zero uses the clock.
	Now time.Time
	// StatePath overrides .xlflow/state/push.json (tests).
	StatePath string
	// Guard rejects the mutation when the workbook appears open or owned by a
	// recorded session. It runs only when the push would actually modify the
	// workbook — a changed-only skip is evaluated first and never reaches it.
	Guard func(ctx context.Context, workbookPath string) error
}

// Result describes a completed file-backend push.
type Result struct {
	// WorkbookPath is the normalized configured workbook path.
	WorkbookPath string
	// PublishPath is the canonical path the artifact was published to.
	PublishPath string
	// Skipped reports a changed-only no-op: source state and saved-file stamp
	// still match the recorded push state.
	Skipped bool
	// Backup is the pre-push backup record, nil for --backup never or skips.
	Backup *backup.Record
	// Meta summarizes the module kinds applied to the rebuilt project.
	Meta packpkg.PackMeta
	// Publication describes the atomic artifact replacement.
	Publication coordination.PublishResult
	// SourceFiles is the number of files covered by the source fingerprint.
	SourceFiles int
	// StatePath is the push-state file path that was (or would be) written.
	StatePath string
	// StateError carries a post-publish push.json write failure. It never
	// fails the push; callers surface it as a warning.
	StateError error
}

// Push applies the source tree to the configured workbook using the default
// background context.
func Push(root string, cfg config.Config, workbookPath string, opts Options) (Result, error) {
	return PushContext(context.Background(), root, cfg, workbookPath, opts)
}

// PushContext runs one file-backend push while holding exclusive leases on the
// managed source roots and shared leases on their ancestors, matching the
// pull backend's coordination contract.
func PushContext(ctx context.Context, root string, cfg config.Config, workbookPath string, opts Options) (Result, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	manager := opts.Coordination
	if manager == nil {
		var err error
		manager, err = coordination.NewDefaultManager()
		if err != nil {
			return Result{}, fmt.Errorf("initialize source-tree coordination: %w", err)
		}
	}
	release, err := acquireSourceTrees(ctx, manager, root, cfg, opts)
	if err != nil {
		return Result{}, err
	}
	defer release()

	statePath := opts.StatePath
	if strings.TrimSpace(statePath) == "" {
		statePath = filepath.Join(root, ".xlflow", "state", "push.json")
	}

	files := discoverSourceFiles(resolvedRoots(root, cfg), cfg.UserForm.CodeSource)
	fingerprint := computeFingerprint(workbookPath, files, cfg.VBA.LineNumbers.Enabled, cfg.VBA.FolderAnnotation)
	if cfg.VBA.LineNumbers.Enabled {
		if err := validateLineNumberSources(files); err != nil {
			return Result{}, err
		}
	}
	if duplicates := findDuplicateModuleNames(files); len(duplicates) > 0 {
		details := make([]string, 0, len(duplicates))
		for _, dup := range duplicates {
			details = append(details, strings.Join(dup, ", "))
		}
		return Result{}, fmt.Errorf("%w: %s", ErrDuplicateModule, strings.Join(details, "; "))
	}

	if opts.ChangedOnly && shouldSkipUnchanged(statePath, fingerprint, workbookPath) {
		return Result{
			WorkbookPath: normalizeFingerprintPath(workbookPath),
			Skipped:      true,
			SourceFiles:  len(files),
			StatePath:    statePath,
		}, nil
	}

	if opts.Guard != nil {
		if err := opts.Guard(ctx, workbookPath); err != nil {
			return Result{}, err
		}
	}

	identity, err := coordination.NewWorkbookIdentity(root, workbookPath)
	if err != nil {
		return Result{}, fmt.Errorf("resolve workbook identity: %w", err)
	}
	// The mutation window is serialized with other workbook writers — a
	// concurrent rollback, pack, or session holding the workbook lease fails
	// this push fast instead of racing the atomic replace. The acquisition is
	// deliberately non-blocking: the source-tree leases are already held, and
	// waiting here while an Excel-backend push holds the workbook lease and
	// waits on the source tree would deadlock.
	lease, err := manager.Acquire(ctx, coordination.AcquireRequest{
		Identity:      identity,
		Command:       "push",
		OperationKind: coordination.OperationMutate,
		ResourceScope: coordination.ResourceWorkbook,
	})
	if err != nil {
		return Result{}, fmt.Errorf("%w: %w", ErrWorkbookLease, err)
	}
	defer func() { _ = lease.Release() }()
	if err := lease.RequireRecoveryAllowed(coordination.RecoveryBlock, false); err != nil {
		return Result{}, fmt.Errorf("%w: %w", ErrRecoveryCheck, err)
	}

	template, err := os.ReadFile(workbookPath)
	if err != nil {
		return Result{}, err
	}

	sources, err := collectSources(root, cfg)
	if err != nil {
		return Result{}, err
	}
	built, meta, err := packpkg.BuildWorkbook(template, sources)
	if err != nil {
		return Result{}, err
	}

	var record *backup.Record
	if opts.BackupMode != "never" {
		created, err := backup.CreateForBackend(root, workbookPath, "before-push", "file", opts.Now)
		if err != nil {
			return Result{}, fmt.Errorf("%w: %v", ErrBackup, err)
		}
		record = &created
	}

	publication, err := coordination.PublishFile(identity.CanonicalPath, built, validateWorkbookArtifact)
	if err != nil {
		// Double %w keeps both the filepush sentinel and any typed
		// coordination.PublishError reachable through errors.Is/As.
		return Result{}, fmt.Errorf("%w: %w", ErrPublish, err)
	}

	result := Result{
		WorkbookPath: normalizeFingerprintPath(workbookPath),
		PublishPath:  identity.CanonicalPath,
		Backup:       record,
		Meta:         meta,
		Publication:  publication,
		SourceFiles:  len(files),
		StatePath:    statePath,
	}
	if err := writePushState(statePath, fingerprint, buildFilePushAppliedTo(workbookPath)); err != nil {
		result.StateError = err
	}
	return result, nil
}

// acquireSourceTrees mirrors filepull's lease acquisition: every managed
// source root is taken exclusively while its ancestors are shared, so source
// mutations cannot interleave with the read of the tree being pushed.
func acquireSourceTrees(ctx context.Context, manager *coordination.Manager, root string, cfg config.Config, opts Options) (func(), error) {
	type lockTarget struct {
		identity coordination.ResourceIdentity
		shared   bool
	}
	roots := resolvedRoots(root, cfg)
	targets := map[string]lockTarget{}
	for _, path := range []string{roots.modules, roots.classes, roots.forms, roots.workbook} {
		identity, err := coordination.NewSourceTreeIdentity(root, path)
		if err != nil {
			return nil, fmt.Errorf("resolve source-tree identity %s: %w", path, err)
		}
		for index, hierarchyIdentity := range coordination.SourceTreeLockHierarchy(identity) {
			target := lockTarget{identity: hierarchyIdentity, shared: index != 0}
			if current, exists := targets[hierarchyIdentity.LockID]; !exists || current.shared {
				targets[hierarchyIdentity.LockID] = target
			}
		}
	}
	lockIDs := slices.Sorted(maps.Keys(targets))
	leases := make([]*coordination.Lease, 0, len(lockIDs))
	release := func() {
		for i := len(leases) - 1; i >= 0; i-- {
			_ = leases[i].Release()
		}
	}
	acquireCtx := ctx
	waitStarted := false
	for _, lockID := range lockIDs {
		target := targets[lockID]
		request := coordination.AcquireRequest{
			Identity:      target.identity,
			Command:       "push",
			OperationKind: coordination.OperationMutate,
			ResourceScope: coordination.ResourceSourceTree,
			Shared:        target.shared,
		}
		lease, err := manager.Acquire(acquireCtx, request)
		if opts.Wait && errors.Is(err, coordination.ErrSourceTreeBusy) {
			if !waitStarted && opts.WaitTimeout > 0 {
				var cancelWait context.CancelFunc
				acquireCtx, cancelWait = context.WithTimeout(ctx, opts.WaitTimeout)
				defer cancelWait()
				waitStarted = true
			}
			request.Wait = true
			lease, err = manager.Acquire(acquireCtx, request)
		}
		if err == nil && waitStarted && acquireCtx.Err() != nil {
			_ = lease.Release()
			err = acquireCtx.Err()
		}
		if err != nil {
			release()
			return nil, err
		}
		leases = append(leases, lease)
	}
	return release, nil
}

// validateWorkbookArtifact performs the structural check required before
// publication: the staged file must be a readable OOXML zip containing a
// non-empty xl/vbaProject.bin. It intentionally does not validate VBA
// semantics; VBE compilation is an Excel-backend guarantee.
func validateWorkbookArtifact(path string) error {
	reader, err := zip.OpenReader(path)
	if err != nil {
		return fmt.Errorf("staged artifact is not a readable zip: %w", err)
	}
	defer func() { _ = reader.Close() }()
	for _, entry := range reader.File {
		if entry.Name != "xl/vbaProject.bin" {
			continue
		}
		stream, err := entry.Open()
		if err != nil {
			return fmt.Errorf("open staged xl/vbaProject.bin: %w", err)
		}
		size, readErr := io.Copy(io.Discard, stream)
		closeErr := stream.Close()
		if readErr != nil {
			return fmt.Errorf("read staged xl/vbaProject.bin: %w", readErr)
		}
		if closeErr != nil {
			return fmt.Errorf("close staged xl/vbaProject.bin: %w", closeErr)
		}
		if size == 0 {
			return errors.New("staged xl/vbaProject.bin is empty")
		}
		return nil
	}
	return errors.New("staged artifact is missing xl/vbaProject.bin")
}
