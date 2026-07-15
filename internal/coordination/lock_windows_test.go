//go:build windows

package coordination

import (
	"bufio"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
	"time"
)

const helperLockID = "xlflow-workbook-v1-aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"

func testIdentity(path, lockID string) WorkbookIdentity {
	return WorkbookIdentity{CanonicalPath: path, LockID: lockID}
}

func testRequest(identity WorkbookIdentity, wait bool) AcquireRequest {
	return AcquireRequest{
		Identity:      identity,
		Command:       CommandID("push"),
		OperationKind: OperationMutate,
		ResourceScope: ResourceWorkbook,
		Wait:          wait,
	}
}

func TestAcquirePublishesMetadataAndReleaseClearsIt(t *testing.T) {
	manager, err := NewManager(t.TempDir())
	if err != nil {
		t.Fatalf("NewManager: %v", err)
	}
	identity := testIdentity(filepath.Join(t.TempDir(), "book.xlsm"), helperLockID)
	lease, err := manager.Acquire(context.Background(), testRequest(identity, false))
	if err != nil {
		t.Fatalf("Acquire: %v", err)
	}

	result, err := manager.Probe(context.Background(), identity)
	if err != nil {
		t.Fatalf("Probe while held: %v", err)
	}
	if !result.Busy || result.Owner == nil {
		t.Fatalf("Probe = %#v, want busy owner", result)
	}
	if result.Owner.PID != os.Getpid() || result.Owner.Command != "push" || result.Owner.Workbook != identity.CanonicalPath {
		t.Fatalf("owner = %#v", result.Owner)
	}
	if result.Owner.Generation == "" || result.Owner.SchemaVersion != ownerSchemaV1 {
		t.Fatalf("owner lacks schema/generation: %#v", result.Owner)
	}

	if err := lease.Release(); err != nil {
		t.Fatalf("Release: %v", err)
	}
	if err := lease.Release(); err != nil {
		t.Fatalf("second Release: %v", err)
	}
	result, err = manager.Probe(context.Background(), identity)
	if err != nil {
		t.Fatalf("Probe after release: %v", err)
	}
	if result.Busy || result.Owner != nil {
		t.Fatalf("Probe after release = %#v", result)
	}
	if _, err := os.Stat(manager.ownerPath(identity)); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("owner metadata remains after release: %v", err)
	}
}

func TestAcquireDifferentWorkbooksDoesNotContend(t *testing.T) {
	manager, err := NewManager(t.TempDir())
	if err != nil {
		t.Fatalf("NewManager: %v", err)
	}
	first := testIdentity(`C:\work\first.xlsm`, helperLockID)
	second := testIdentity(`C:\work\second.xlsm`, "xlflow-workbook-v1-bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb")
	firstLease, err := manager.Acquire(context.Background(), testRequest(first, false))
	if err != nil {
		t.Fatalf("Acquire(first): %v", err)
	}
	defer func() { _ = firstLease.Release() }()
	secondLease, err := manager.Acquire(context.Background(), testRequest(second, false))
	if err != nil {
		t.Fatalf("Acquire(second): %v", err)
	}
	defer func() { _ = secondLease.Release() }()
}

func TestAcquireWaitHonorsContext(t *testing.T) {
	manager, err := NewManager(t.TempDir())
	if err != nil {
		t.Fatalf("NewManager: %v", err)
	}
	identity := testIdentity(`C:\work\book.xlsm`, helperLockID)
	lease, err := manager.Acquire(context.Background(), testRequest(identity, false))
	if err != nil {
		t.Fatalf("Acquire(owner): %v", err)
	}
	defer func() { _ = lease.Release() }()

	ctx, cancel := context.WithTimeout(context.Background(), 100*time.Millisecond)
	defer cancel()
	_, err = manager.Acquire(ctx, testRequest(identity, true))
	if !errors.Is(err, context.DeadlineExceeded) {
		t.Fatalf("Acquire(wait) error = %v, want deadline exceeded", err)
	}
}

func TestAcquireWithoutWaitDoesNotBlockOnPublicationGuard(t *testing.T) {
	manager, err := NewManager(t.TempDir())
	if err != nil {
		t.Fatalf("NewManager: %v", err)
	}
	identity := testIdentity(`C:\work\book.xlsm`, helperLockID)
	file, err := manager.openLock(identity)
	if err != nil {
		t.Fatalf("openLock: %v", err)
	}
	defer func() { _ = file.Close() }()
	acquired, err := platformTryLock(file, publicationByte)
	if err != nil || !acquired {
		t.Fatalf("lock publication guard = %v, %v", acquired, err)
	}
	defer func() { _ = platformUnlock(file, publicationByte) }()

	started := time.Now()
	_, err = manager.Acquire(context.Background(), testRequest(identity, false))
	if !errors.Is(err, ErrWorkbookBusy) {
		t.Fatalf("Acquire error = %v, want workbook busy", err)
	}
	if elapsed := time.Since(started); elapsed > 500*time.Millisecond {
		t.Fatalf("non-wait acquisition blocked for %s", elapsed)
	}
}

func TestAcquireRejectsAlreadyCancelledContext(t *testing.T) {
	manager, err := NewManager(t.TempDir())
	if err != nil {
		t.Fatalf("NewManager: %v", err)
	}
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	_, err = manager.Acquire(ctx, testRequest(testIdentity(`C:\work\book.xlsm`, helperLockID), false))
	if !errors.Is(err, context.Canceled) {
		t.Fatalf("Acquire error = %v, want context canceled", err)
	}
}

func TestMalformedMetadataDoesNotOverrideAuthoritativeLock(t *testing.T) {
	manager, err := NewManager(t.TempDir())
	if err != nil {
		t.Fatalf("NewManager: %v", err)
	}
	identity := testIdentity(`C:\work\book.xlsm`, helperLockID)
	lease, err := manager.Acquire(context.Background(), testRequest(identity, false))
	if err != nil {
		t.Fatalf("Acquire: %v", err)
	}
	if err := os.WriteFile(manager.ownerPath(identity), []byte("not-json"), 0o600); err != nil {
		t.Fatalf("corrupt owner metadata: %v", err)
	}

	result, err := manager.Probe(context.Background(), identity)
	if err != nil {
		t.Fatalf("Probe while held: %v", err)
	}
	if !result.Busy || result.Owner != nil {
		t.Fatalf("Probe = %#v, want busy without owner", result)
	}
	if err := lease.Release(); err != nil {
		t.Fatalf("Release: %v", err)
	}
	result, err = manager.Probe(context.Background(), identity)
	if err != nil || result.Busy {
		t.Fatalf("Probe after release = %#v, %v", result, err)
	}
	if _, err := os.Stat(manager.ownerPath(identity)); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("Probe did not remove stale metadata: %v", err)
	}
}

func TestPartialMetadataIsIgnoredWhileAuthoritativeLockRemainsBusy(t *testing.T) {
	manager, err := NewManager(t.TempDir())
	if err != nil {
		t.Fatalf("NewManager: %v", err)
	}
	identity := testIdentity(`C:\work\book.xlsm`, helperLockID)
	lease, err := manager.Acquire(context.Background(), testRequest(identity, false))
	if err != nil {
		t.Fatalf("Acquire: %v", err)
	}
	defer func() { _ = lease.Release() }()
	partial := `{"schema_version":1,"generation":"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"}`
	if err := os.WriteFile(manager.ownerPath(identity), []byte(partial), 0o600); err != nil {
		t.Fatalf("write partial owner metadata: %v", err)
	}
	result, err := manager.Probe(context.Background(), identity)
	if err != nil {
		t.Fatalf("Probe: %v", err)
	}
	if !result.Busy || result.Owner != nil {
		t.Fatalf("Probe = %#v, want authoritative busy with no owner", result)
	}
}

func TestAcquireAtomicallyReplacesStaleMetadata(t *testing.T) {
	manager, err := NewManager(t.TempDir())
	if err != nil {
		t.Fatalf("NewManager: %v", err)
	}
	identity := testIdentity(`C:\work\book.xlsm`, helperLockID)
	stale := OwnerMetadata{SchemaVersion: ownerSchemaV1, Generation: "stale", Workbook: identity.CanonicalPath}
	data, err := json.Marshal(stale)
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(manager.ownerPath(identity), data, 0o600); err != nil {
		t.Fatalf("write stale metadata: %v", err)
	}

	lease, err := manager.Acquire(context.Background(), testRequest(identity, false))
	if err != nil {
		t.Fatalf("Acquire: %v", err)
	}
	defer func() { _ = lease.Release() }()
	owner := manager.readOwnerBestEffort(identity)
	if owner == nil || owner.Generation == "stale" {
		t.Fatalf("owner metadata was not replaced: %#v", owner)
	}
}

func TestOwnerCleanupRequiresMatchingGeneration(t *testing.T) {
	manager, err := NewManager(t.TempDir())
	if err != nil {
		t.Fatalf("NewManager: %v", err)
	}
	identity := testIdentity(`C:\work\book.xlsm`, helperLockID)
	owner := OwnerMetadata{
		SchemaVersion: ownerSchemaV1,
		Generation:    "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
		Workbook:      identity.CanonicalPath,
		PID:           os.Getpid(),
		Command:       "push",
		OperationKind: OperationMutate,
		ResourceScope: ResourceWorkbook,
		StartedAt:     time.Now().UTC(),
	}
	data, err := json.Marshal(owner)
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(manager.ownerPath(identity), data, 0o600); err != nil {
		t.Fatalf("write owner metadata: %v", err)
	}
	if err := manager.removeOwnerIfMatches(identity, "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb"); err != nil {
		t.Fatalf("removeOwnerIfMatches: %v", err)
	}
	if got := manager.readOwnerBestEffort(identity); got == nil || got.Generation != owner.Generation {
		t.Fatalf("mismatched cleanup removed current owner: %#v", got)
	}
}

func TestSubprocessContentionPublishesOwnerAndCrashReleasesLock(t *testing.T) {
	stateDir := t.TempDir()
	identity := testIdentity(`C:\work\book.xlsm`, helperLockID)
	cmd := exec.Command(os.Args[0], "-test.run=^TestCoordinationLockHelper$")
	cmd.Env = append(os.Environ(),
		"XLFLOW_COORDINATION_HELPER=1",
		"XLFLOW_COORDINATION_STATE="+stateDir,
		"XLFLOW_COORDINATION_PATH="+identity.CanonicalPath,
		"XLFLOW_COORDINATION_LOCK_ID="+identity.LockID,
	)
	stdout, err := cmd.StdoutPipe()
	if err != nil {
		t.Fatalf("StdoutPipe: %v", err)
	}
	cmd.Stderr = os.Stderr
	if err := cmd.Start(); err != nil {
		t.Fatalf("start helper: %v", err)
	}
	killed := false
	defer func() {
		if !killed && cmd.Process != nil {
			_ = cmd.Process.Kill()
			_ = cmd.Wait()
		}
	}()
	ready := make(chan string, 1)
	go func() {
		scanner := bufio.NewScanner(stdout)
		if scanner.Scan() {
			ready <- scanner.Text()
			return
		}
		ready <- ""
	}()
	select {
	case line := <-ready:
		if line != "READY" {
			t.Fatalf("helper output = %q", line)
		}
	case <-time.After(10 * time.Second):
		t.Fatal("helper did not acquire lock")
	}

	manager, err := NewManager(stateDir)
	if err != nil {
		t.Fatalf("NewManager: %v", err)
	}
	_, err = manager.Acquire(context.Background(), testRequest(identity, false))
	var busy *BusyError
	if !errors.As(err, &busy) || !errors.Is(err, ErrWorkbookBusy) {
		t.Fatalf("Acquire contention error = %v", err)
	}
	if busy.Owner == nil || busy.Owner.PID != cmd.Process.Pid || busy.Owner.Command != "push" {
		t.Fatalf("busy owner = %#v, helper PID %d", busy.Owner, cmd.Process.Pid)
	}

	if err := cmd.Process.Kill(); err != nil {
		t.Fatalf("kill helper: %v", err)
	}
	if err := cmd.Wait(); err == nil {
		t.Fatal("killed helper unexpectedly exited successfully")
	}
	killed = true

	deadline := time.Now().Add(5 * time.Second)
	for {
		lease, acquireErr := manager.Acquire(context.Background(), testRequest(identity, false))
		if acquireErr == nil {
			if err := lease.Release(); err != nil {
				t.Fatalf("Release after crash: %v", err)
			}
			break
		}
		if !errors.Is(acquireErr, ErrWorkbookBusy) || time.Now().After(deadline) {
			t.Fatalf("Acquire after crash: %v", acquireErr)
		}
		time.Sleep(25 * time.Millisecond)
	}
}

func TestCoordinationLockHelper(t *testing.T) {
	if os.Getenv("XLFLOW_COORDINATION_HELPER") != "1" {
		t.Skip("subprocess helper")
	}
	manager, err := NewManager(os.Getenv("XLFLOW_COORDINATION_STATE"))
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(2)
	}
	identity := testIdentity(os.Getenv("XLFLOW_COORDINATION_PATH"), os.Getenv("XLFLOW_COORDINATION_LOCK_ID"))
	lease, err := manager.Acquire(context.Background(), testRequest(identity, false))
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(2)
	}
	defer func() { _ = lease.Release() }()
	fmt.Println("READY")
	if value := os.Getenv("XLFLOW_COORDINATION_HELPER_EXIT"); value != "" {
		code, _ := strconv.Atoi(strings.TrimSpace(value))
		os.Exit(code)
	}
	select {}
}
