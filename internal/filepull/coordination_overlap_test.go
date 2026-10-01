package filepull

import (
	"context"
	"errors"
	"path/filepath"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/coordination"
)

func TestSourceTreeCoordinationBlocksAncestorDescendantAcrossProjects(t *testing.T) {
	manager, err := coordination.NewManager(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	shared := filepath.Join(t.TempDir(), "shared")
	parentConfig := testConfig()
	parentConfig.Src.Modules = shared
	parentConfig.Src.Classes = shared
	parentConfig.Src.Workbook = shared
	childConfig := testConfig()
	childConfig.Src.Modules = filepath.Join(shared, "modules")
	childConfig.Src.Classes = filepath.Join(shared, "classes")
	childConfig.Src.Workbook = filepath.Join(shared, "workbook")

	release, err := acquireSourceTrees(t.Context(), t.TempDir(), parentConfig, PullOptions{Coordination: manager})
	if err != nil {
		t.Fatal(err)
	}
	defer release()
	_, err = acquireSourceTrees(t.Context(), t.TempDir(), childConfig, PullOptions{Coordination: manager})
	if !errors.Is(err, coordination.ErrSourceTreeBusy) {
		t.Fatalf("descendant acquisition error = %v, want source-tree busy", err)
	}
}

func TestSourceTreeCoordinationAllowsDisjointSiblings(t *testing.T) {
	manager, err := coordination.NewManager(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	shared := t.TempDir()
	firstConfig := testConfig()
	firstConfig.Src.Modules = filepath.Join(shared, "first")
	firstConfig.Src.Classes = filepath.Join(shared, "first")
	firstConfig.Src.Workbook = filepath.Join(shared, "first")
	secondConfig := testConfig()
	secondConfig.Src.Modules = filepath.Join(shared, "second")
	secondConfig.Src.Classes = filepath.Join(shared, "second")
	secondConfig.Src.Workbook = filepath.Join(shared, "second")

	releaseFirst, err := acquireSourceTrees(t.Context(), t.TempDir(), firstConfig, PullOptions{Coordination: manager})
	if err != nil {
		t.Fatal(err)
	}
	defer releaseFirst()
	releaseSecond, err := acquireSourceTrees(t.Context(), t.TempDir(), secondConfig, PullOptions{Coordination: manager})
	if err != nil {
		t.Fatalf("disjoint sibling acquisition failed: %v", err)
	}
	defer releaseSecond()
}

func TestSourceTreeWaitTimeoutStartsOnlyAfterContention(t *testing.T) {
	manager, err := coordination.NewManager(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	cfg := testConfig()
	release, err := acquireSourceTrees(t.Context(), t.TempDir(), cfg, PullOptions{
		Coordination: manager,
		Wait:         true,
		WaitTimeout:  time.Nanosecond,
	})
	if err != nil {
		t.Fatalf("uncontended acquisition consumed wait timeout: %v", err)
	}
	release()

	root := t.TempDir()
	identity, err := coordination.NewSourceTreeIdentity(root, filepath.Join(root, cfg.Src.Modules))
	if err != nil {
		t.Fatal(err)
	}
	owner, err := manager.Acquire(t.Context(), coordination.AcquireRequest{
		Identity: identity, Command: "pull", OperationKind: coordination.OperationMutate,
		ResourceScope: coordination.ResourceSourceTree,
	})
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = owner.Release() }()
	ctx, cancel := context.WithTimeout(t.Context(), time.Second)
	defer cancel()
	_, err = acquireSourceTrees(ctx, root, cfg, PullOptions{
		Coordination: manager,
		Wait:         true,
		WaitTimeout:  10 * time.Millisecond,
	})
	if !errors.Is(err, context.DeadlineExceeded) {
		t.Fatalf("contended acquisition error = %v, want deadline exceeded", err)
	}
}
