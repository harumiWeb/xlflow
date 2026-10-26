package sourceedit

import (
	"context"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"slices"
	"testing"
)

func batchTestChanges(t *testing.T, dir string, names ...string) []Change {
	t.Helper()
	changes := make([]Change, len(names))
	for i, name := range names {
		path := writeTestFile(t, dir, name, []byte("original 😀\r\n"), 0o640)
		changes[i] = Change{Snapshot: captureTestSnapshot(t, path), Replacement: []byte("replacement 日本語\r\n")}
	}
	return changes
}

func TestBatchAppliesAndPreservesModes(t *testing.T) {
	dir := t.TempDir()
	changes := batchTestChanges(t, dir, "a.bas", "b.cls", "noop.bas")
	changes[2].Replacement = changes[2].Snapshot.Source
	info, err := os.Stat(changes[2].Snapshot.path)
	if err != nil {
		t.Fatal(err)
	}
	result, err := ApplyBatch(t.Context(), changes)
	if err != nil || !result.Committed || len(result.Artifacts) != 0 {
		t.Fatalf("result=%+v error=%v", result, err)
	}
	for i, change := range changes {
		assertFileContents(t, change.Snapshot.path, change.Replacement)
		after, err := os.Stat(change.Snapshot.path)
		if err != nil || after.Mode().Perm() != change.Snapshot.mode {
			t.Fatalf("mode/stat %s: %v %v", change.Snapshot.path, after, err)
		}
		want := "applied"
		if i == 2 {
			want = "unchanged"
			if !os.SameFile(info, after) {
				t.Fatal("no-op file identity changed")
			}
		}
		if result.Files[i].State != want {
			t.Fatalf("file[%d]=%+v", i, result.Files[i])
		}
	}
	assertNoTemporaryFiles(t, dir)
}

func TestBatchPreflightAndStagingFailuresPublishNothing(t *testing.T) {
	for _, scenario := range []string{"stale", "unchanged stale", "read-only", "stage failure", "cancelled", "stale after staging"} {
		t.Run(scenario, func(t *testing.T) {
			dir := t.TempDir()
			changes := batchTestChanges(t, dir, "a.bas", "b.bas")
			ctx, cancel := context.WithCancel(t.Context())
			defer cancel()
			want := errors.New("stage failure")
			ops := batchOperations{}
			switch scenario {
			case "stale", "unchanged stale":
				if scenario == "unchanged stale" {
					changes[1].Replacement = changes[1].Snapshot.Source
				}
				if err := os.WriteFile(changes[1].Snapshot.path, []byte("external"), 0o600); err != nil {
					t.Fatal(err)
				}
				want = ErrStale
			case "read-only":
				path := changes[1].Snapshot.path
				if err := os.Chmod(path, 0o444); err != nil {
					t.Fatal(err)
				}
				t.Cleanup(func() { _ = os.Chmod(path, 0o600) })
				changes[1].Snapshot = captureTestSnapshot(t, path)
				want = ErrReadOnly
			case "stage failure":
				calls := 0
				ops.stage = func(ctx context.Context, source *Snapshot, data []byte) (*Snapshot, error) {
					calls++
					if calls == 4 {
						return nil, want
					}
					return stageBatchFile(ctx, source, data)
				}
			case "cancelled":
				cancel()
				want = context.Canceled
			case "stale after staging":
				calls := 0
				ops.stage = func(ctx context.Context, source *Snapshot, data []byte) (*Snapshot, error) {
					stage, err := stageBatchFile(ctx, source, data)
					calls++
					if calls == 4 {
						if writeErr := os.WriteFile(changes[1].Snapshot.path, []byte("external"), 0o600); writeErr != nil {
							t.Fatal(writeErr)
						}
					}
					return stage, err
				}
				want = ErrStale
			}
			for _, change := range changes {
				change.Snapshot.ops.publish = func(string, string) error { t.Fatal("published during failed preflight"); return nil }
			}
			result, err := applyBatch(ctx, changes, ops)
			if !errors.Is(err, want) || result.Committed {
				t.Fatalf("result=%+v error=%v, want %v", result, err, want)
			}
			assertFileContents(t, changes[0].Snapshot.path, changes[0].Snapshot.Source)
			if scenario == "stale" || scenario == "unchanged stale" || scenario == "stale after staging" {
				assertFileContents(t, changes[1].Snapshot.path, []byte("external"))
			} else {
				assertFileContents(t, changes[1].Snapshot.path, changes[1].Snapshot.Source)
			}
			assertNoTemporaryFiles(t, dir)
		})
	}
}

func TestBatchFailureAndCancellationRollbackInReverseOrder(t *testing.T) {
	for _, cancelBeforeLast := range []bool{false, true} {
		t.Run(map[bool]string{false: "publish failure", true: "cancellation"}[cancelBeforeLast], func(t *testing.T) {
			dir := t.TempDir()
			changes := batchTestChanges(t, dir, "a.bas", "b.bas", "c.bas")
			ctx, cancel := context.WithCancel(t.Context())
			defer cancel()
			want := errors.New("last publication failed")
			var publications []string
			for _, change := range changes {
				change.Snapshot.ops.publish = func(temp, target string) error {
					publications = append(publications, filepath.Base(target))
					if target == changes[2].Snapshot.path {
						return want
					}
					return platformAtomicReplace(temp, target)
				}
			}
			if cancelBeforeLast {
				changes[2].Snapshot.ops.beforePublish = cancel
				want = context.Canceled
			}
			result, err := ApplyBatch(ctx, changes)
			if !errors.Is(err, want) || result.Committed {
				t.Fatalf("result=%+v error=%v", result, err)
			}
			batchErr, ok := errors.AsType[*BatchError](err)
			if !ok || batchErr.Rollback != nil || batchErr.Cleanup != nil {
				t.Fatalf("rollback/cleanup failed: %+v", err)
			}
			for i, change := range changes {
				assertFileContents(t, change.Snapshot.path, change.Snapshot.Source)
				wantState := "restored"
				if i == 2 {
					wantState = "not_applied"
				}
				if result.Files[i].State != wantState {
					t.Fatalf("file[%d]=%+v", i, result.Files[i])
				}
			}
			wantOrder := []string{"a.bas", "b.bas", "c.bas", "b.bas", "a.bas"}
			if cancelBeforeLast {
				wantOrder = []string{"a.bas", "b.bas", "b.bas", "a.bas"}
			}
			if !reflect.DeepEqual(publications, wantOrder) {
				t.Fatalf("publication order=%v, want %v", publications, wantOrder)
			}
			assertNoTemporaryFiles(t, dir)
		})
	}
}

func TestBatchFailedRollbackRetainsOriginalWithoutOverwritingExternalEdits(t *testing.T) {
	for _, external := range []bool{false, true} {
		t.Run(map[bool]string{false: "rollback I/O failure", true: "external conflict"}[external], func(t *testing.T) {
			dir := t.TempDir()
			changes := batchTestChanges(t, dir, "a.bas", "b.bas")
			count := 0
			changes[0].Snapshot.ops.publish = func(temp, target string) error {
				count++
				if count > 1 {
					if external {
						t.Fatal("attempted to overwrite external content")
					}
					return errors.New("rollback blocked")
				}
				return platformAtomicReplace(temp, target)
			}
			changes[1].Snapshot.ops.beforePublish = func() {
				if external {
					if err := os.WriteFile(changes[0].Snapshot.path, []byte("external latest"), 0o600); err != nil {
						t.Fatal(err)
					}
				}
			}
			changes[1].Snapshot.ops.publish = func(string, string) error { return errors.New("publish blocked") }
			result, err := ApplyBatch(t.Context(), changes)
			batchErr, ok := errors.AsType[*BatchError](err)
			if !ok || batchErr.Rollback == nil || result.Committed || result.Files[0].State != "recovery_required" || len(result.Artifacts) != 1 {
				t.Fatalf("result=%+v error=%v", result, err)
			}
			assertFileContents(t, result.Files[0].RecoveryPath, changes[0].Snapshot.Source)
			want := changes[0].Replacement
			if external {
				want = []byte("external latest")
			}
			assertFileContents(t, changes[0].Snapshot.path, want)
			assertFileContents(t, changes[1].Snapshot.path, changes[1].Snapshot.Source)
		})
	}
}

func TestBatchCleanupFailureReportsCommittedStateAndArtifacts(t *testing.T) {
	dir := t.TempDir()
	changes := batchTestChanges(t, dir, "a.bas", "b.bas")
	want := errors.New("cleanup blocked")
	result, err := applyBatch(t.Context(), changes, batchOperations{remove: func(string) error { return want }})
	batchErr, ok := errors.AsType[*BatchError](err)
	if !ok || batchErr.Cause != nil || batchErr.Rollback != nil || !errors.Is(batchErr.Cleanup, want) || !result.Committed || len(result.Artifacts) != 2 {
		t.Fatalf("result=%+v error=%v", result, err)
	}
	for i, change := range changes {
		assertFileContents(t, change.Snapshot.path, change.Replacement)
		assertFileContents(t, result.Artifacts[i], change.Snapshot.Source)
		if result.Files[i].State != "applied" {
			t.Fatalf("file=%+v", result.Files[i])
		}
	}
}

func TestBatchNoOpReadOnlyAndCommitPoint(t *testing.T) {
	t.Run("read-only no-op", func(t *testing.T) {
		dir := t.TempDir()
		changes := batchTestChanges(t, dir, "a.bas")
		path := changes[0].Snapshot.path
		if err := os.Chmod(path, 0o444); err != nil {
			t.Fatal(err)
		}
		t.Cleanup(func() { _ = os.Chmod(path, 0o600) })
		changes[0].Snapshot = captureTestSnapshot(t, path)
		changes[0].Replacement = changes[0].Snapshot.Source
		result, err := applyBatch(t.Context(), changes, batchOperations{stage: func(context.Context, *Snapshot, []byte) (*Snapshot, error) {
			t.Fatal("staged a no-op")
			return nil, nil
		}})
		if err != nil || !result.Committed || result.Files[0].State != "unchanged" {
			t.Fatalf("result=%+v error=%v", result, err)
		}
		assertNoTemporaryFiles(t, dir)
	})
	t.Run("cancel after final publication", func(t *testing.T) {
		dir := t.TempDir()
		changes := batchTestChanges(t, dir, "a.bas")
		ctx, cancel := context.WithCancel(t.Context())
		defer cancel()
		changes[0].Snapshot.ops.publish = func(temp, target string) error { err := platformAtomicReplace(temp, target); cancel(); return err }
		result, err := ApplyBatch(ctx, changes)
		if err != nil || !result.Committed {
			t.Fatalf("result=%+v error=%v", result, err)
		}
		assertFileContents(t, changes[0].Snapshot.path, changes[0].Replacement)
		assertNoTemporaryFiles(t, dir)
	})
}

func TestBatchRejectsDuplicatePhysicalSources(t *testing.T) {
	for _, hardLink := range []bool{false, true} {
		t.Run(map[bool]string{false: "same snapshot", true: "hard link"}[hardLink], func(t *testing.T) {
			dir := t.TempDir()
			changes := batchTestChanges(t, dir, "a.bas")
			second := changes[0]
			if hardLink {
				link := filepath.Join(dir, "alias.bas")
				if err := os.Link(changes[0].Snapshot.path, link); err != nil {
					t.Skipf("hard links unavailable: %v", err)
				}
				second.Snapshot = captureTestSnapshot(t, link)
			}
			result, err := ApplyBatch(t.Context(), append(changes, second))
			if err == nil || result.Committed {
				t.Fatalf("result=%+v error=%v", result, err)
			}
			assertFileContents(t, changes[0].Snapshot.path, changes[0].Snapshot.Source)
			assertNoTemporaryFiles(t, dir)
		})
	}
}

func TestBatchRollbackContinuesAfterOneRestoreFails(t *testing.T) {
	dir := t.TempDir()
	changes := batchTestChanges(t, dir, "a.bas", "b.bas", "c.bas")
	calls := 0
	changes[1].Snapshot.ops.publish = func(temp, target string) error {
		calls++
		if calls == 2 {
			return errors.New("restore b failed")
		}
		return platformAtomicReplace(temp, target)
	}
	changes[2].Snapshot.ops.publish = func(string, string) error { return errors.New("publish c failed") }
	result, err := ApplyBatch(t.Context(), changes)
	if err == nil || result.Files[0].State != "restored" || result.Files[1].State != "recovery_required" || result.Files[2].State != "not_applied" {
		t.Fatalf("result=%+v error=%v", result, err)
	}
	assertFileContents(t, changes[0].Snapshot.path, changes[0].Snapshot.Source)
	assertFileContents(t, changes[1].Snapshot.path, changes[1].Replacement)
	assertFileContents(t, result.Files[1].RecoveryPath, changes[1].Snapshot.Source)
}

func TestBatchRejectsChangedParentAndReplacementIdentityDuringRollback(t *testing.T) {
	for _, parentChanged := range []bool{false, true} {
		t.Run(map[bool]string{false: "replacement identity", true: "parent identity"}[parentChanged], func(t *testing.T) {
			root := t.TempDir()
			dir := filepath.Join(root, "sources")
			if err := os.Mkdir(dir, 0o700); err != nil {
				t.Fatal(err)
			}
			changes := batchTestChanges(t, dir, "a.bas", "b.bas")
			changes[1].Snapshot.ops.beforePublish = func() {
				if parentChanged {
					// Relocating the directory moves the original recovery material.
					// Keep this as a publication-preflight test instead of rollback,
					// below, where the original directory entry must be retained.
					return
				}
				path := changes[0].Snapshot.path
				if err := os.Remove(path); err != nil {
					t.Fatal(err)
				}
				writeTestFile(t, dir, "a.bas", changes[0].Replacement, 0o640)
			}
			if parentChanged {
				moved := filepath.Join(root, "moved")
				if err := os.Rename(dir, moved); err != nil {
					t.Fatal(err)
				}
				if err := os.Mkdir(dir, 0o700); err != nil {
					t.Fatal(err)
				}
				writeTestFile(t, dir, "a.bas", []byte("external"), 0o640)
				writeTestFile(t, dir, "b.bas", []byte("external"), 0o640)
				result, err := ApplyBatch(t.Context(), changes)
				if !errors.Is(err, ErrStale) || result.Committed {
					t.Fatalf("result=%+v error=%v", result, err)
				}
				assertFileContents(t, changes[0].Snapshot.path, []byte("external"))
				assertNoTemporaryFiles(t, moved)
				assertNoTemporaryFiles(t, dir)
				return
			}
			changes[1].Snapshot.ops.publish = func(string, string) error { return errors.New("publish blocked") }
			result, err := ApplyBatch(t.Context(), changes)
			if !errors.Is(err, ErrStale) || result.Files[0].State != "recovery_required" {
				t.Fatalf("result=%+v error=%v", result, err)
			}
			assertFileContents(t, result.Files[0].RecoveryPath, changes[0].Snapshot.Source)
			assertFileContents(t, changes[0].Snapshot.path, changes[0].Replacement)
		})
	}
}

func TestBatchReportsUnremovedIncompleteStage(t *testing.T) {
	dir := t.TempDir()
	changes := batchTestChanges(t, dir, "a.bas")
	orphan := writeTestFile(t, dir, ".sourceedit-incomplete.tmp", []byte("partial stage"), 0o600)
	want := errors.New("cleanup failed")
	result, err := applyBatch(t.Context(), changes, batchOperations{stage: func(context.Context, *Snapshot, []byte) (*Snapshot, error) {
		return nil, &incompleteStageError{path: orphan, cause: errors.New("incomplete write"), cleanup: want}
	}})
	batchErr, ok := errors.AsType[*BatchError](err)
	if !ok || !errors.Is(batchErr.Cleanup, want) || len(result.Artifacts) != 1 || result.Artifacts[0] != orphan {
		t.Fatalf("result=%+v error=%v", result, err)
	}
	assertFileContents(t, changes[0].Snapshot.path, changes[0].Snapshot.Source)
}

func TestBatchRecoveryCopySurvivesChangedParent(t *testing.T) {
	root := t.TempDir()
	dir := filepath.Join(root, "sources")
	if err := os.Mkdir(dir, 0o700); err != nil {
		t.Fatal(err)
	}
	changes := batchTestChanges(t, dir, "a.bas", "b.bas")
	moved := filepath.Join(root, "moved")
	changes[1].Snapshot.ops.beforePublish = func() {
		if err := os.Rename(dir, moved); err != nil {
			t.Fatal(err)
		}
		if err := os.Mkdir(dir, 0o700); err != nil {
			t.Fatal(err)
		}
		writeTestFile(t, dir, "a.bas", []byte("external"), 0o600)
		writeTestFile(t, dir, "b.bas", []byte("external"), 0o600)
	}
	result, err := ApplyBatch(t.Context(), changes)
	batchErr, ok := errors.AsType[*BatchError](err)
	if !ok || batchErr.Rollback == nil || result.Files[0].State != "recovery_required" {
		t.Fatalf("result=%+v error=%v", result, err)
	}
	recovery := result.Files[0].RecoveryPath
	if filepath.Dir(recovery) == dir {
		t.Fatalf("recovery points to replaced parent: %s", recovery)
	}
	t.Cleanup(func() { _ = os.Remove(recovery) })
	assertFileContents(t, recovery, changes[0].Snapshot.Source)
	assertFileContents(t, changes[0].Snapshot.path, []byte("external"))
	assertFileContents(t, changes[1].Snapshot.path, []byte("external"))
}

func TestBatchCorruptBackupIsNotRestored(t *testing.T) {
	dir := t.TempDir()
	changes := batchTestChanges(t, dir, "a.bas", "b.bas")
	stages := 0
	ops := batchOperations{stage: func(ctx context.Context, source *Snapshot, data []byte) (*Snapshot, error) {
		stage, err := stageBatchFile(ctx, source, data)
		stages++
		if stages == 1 {
			changes[1].Snapshot.ops.beforePublish = func() {
				if err := os.WriteFile(stage.path, []byte("corrupt backup"), 0o600); err != nil {
					t.Fatal(err)
				}
			}
		}
		return stage, err
	}}
	changes[1].Snapshot.ops.publish = func(string, string) error { return errors.New("publication failed") }
	result, err := applyBatch(t.Context(), changes, ops)
	if err == nil || result.Files[0].State != "recovery_required" {
		t.Fatalf("result=%+v error=%v", result, err)
	}
	recovery := result.Files[0].RecoveryPath
	t.Cleanup(func() { _ = os.Remove(recovery) })
	assertFileContents(t, recovery, changes[0].Snapshot.Source)
	assertFileContents(t, changes[0].Snapshot.path, changes[0].Replacement)
}

func TestBatchBackupLostDuringFailedRestore(t *testing.T) {
	for _, corrupt := range []bool{false, true} {
		for _, fallbackFails := range []bool{false, true} {
			t.Run(fmt.Sprintf("corrupt=%t/fallbackFails=%t", corrupt, fallbackFails), func(t *testing.T) {
				changes := batchTestChanges(t, t.TempDir(), "a.bas", "b.bas")
				var backupPath string
				stageCalls := 0
				ops := batchOperations{stage: func(ctx context.Context, source *Snapshot, data []byte) (*Snapshot, error) {
					staged, err := stageBatchFile(ctx, source, data)
					stageCalls++
					if stageCalls == 1 && err == nil {
						backupPath = staged.path
					}
					return staged, err
				}}
				storageErr := errors.New("recovery storage unavailable")
				if fallbackFails {
					ops.recover = func([]byte) (*Snapshot, error) { return nil, storageErr }
				}
				restoreErr := errors.New("restore failed after backup changed")
				changes[0].Snapshot.ops.publish = func(temp, target string) error {
					if temp != backupPath {
						return platformAtomicReplace(temp, target)
					}
					var err error
					if corrupt {
						err = os.WriteFile(backupPath, []byte("corrupted backup"), 0o600)
					} else {
						err = os.Remove(backupPath)
					}
					if err != nil {
						t.Fatal(err)
					}
					return restoreErr
				}
				changes[1].Snapshot.ops.publish = func(string, string) error { return errors.New("second publish failed") }
				result, err := applyBatch(t.Context(), changes, ops)
				if !errors.Is(err, restoreErr) || result.Committed || result.Files[0].State != "recovery_required" || result.Files[1].State != "not_applied" {
					t.Fatalf("result=%+v error=%v", result, err)
				}
				recovery := result.Files[0].RecoveryPath
				if fallbackFails {
					if !errors.Is(err, storageErr) || recovery != "" {
						t.Fatalf("unusable recovery path: result=%+v error=%v", result, err)
					}
				} else {
					if recovery == "" || recovery == backupPath || !slices.Contains(result.Artifacts, recovery) {
						t.Fatalf("no safe fallback: %+v", result)
					}
					t.Cleanup(func() { _ = os.Remove(recovery) })
					assertFileContents(t, recovery, changes[0].Snapshot.Source)
				}
				if slices.Contains(result.Artifacts, backupPath) != corrupt {
					t.Fatalf("incorrect backup artifact state: %+v", result)
				}
				assertFileContents(t, changes[0].Snapshot.path, changes[0].Replacement)
				assertFileContents(t, changes[1].Snapshot.path, changes[1].Snapshot.Source)
			})
		}
	}
}

func TestBatchUnavailableBackupAndRecoveryStorageReportNoUsableRecoveryPath(t *testing.T) {
	dir := t.TempDir()
	changes := batchTestChanges(t, dir, "a.bas", "b.bas")
	stages := 0
	want := errors.New("recovery storage unavailable")
	ops := batchOperations{
		recover: func([]byte) (*Snapshot, error) { return nil, want },
		stage: func(ctx context.Context, source *Snapshot, data []byte) (*Snapshot, error) {
			stage, err := stageBatchFile(ctx, source, data)
			stages++
			if stages == 1 {
				changes[1].Snapshot.ops.beforePublish = func() {
					if err := os.Remove(stage.path); err != nil {
						t.Fatal(err)
					}
				}
			}
			return stage, err
		},
	}
	changes[1].Snapshot.ops.publish = func(string, string) error { return errors.New("publication failed") }
	result, err := applyBatch(t.Context(), changes, ops)
	if !errors.Is(err, want) || result.Files[0].State != "recovery_required" || result.Files[0].RecoveryPath != "" || len(result.Artifacts) != 0 {
		t.Fatalf("result=%+v error=%v", result, err)
	}
	assertFileContents(t, changes[0].Snapshot.path, changes[0].Replacement)
}
