package coordination

import (
	"errors"
	"path/filepath"
	"testing"
)

func TestAcquireSourceTreePublishesTypedMetadata(t *testing.T) {
	manager, err := NewManager(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	identity, err := NewSourceTreeIdentity(t.TempDir(), "src/modules")
	if err != nil {
		t.Fatal(err)
	}
	lease, err := manager.Acquire(t.Context(), AcquireRequest{
		Identity: identity, Command: "pull", OperationKind: OperationMutate,
		ResourceScope: ResourceSourceTree,
	})
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = lease.Release() }()
	owner := manager.readOwnerBestEffort(identity)
	if owner == nil || owner.ResourceScope != ResourceSourceTree || owner.ResourcePath != identity.CanonicalPath || owner.Workbook != "" {
		t.Fatalf("source-tree owner metadata = %#v", owner)
	}
	busy := &BusyError{Identity: identity, ResourceScope: ResourceSourceTree, Owner: owner}
	if busy.Code() != SourceTreeBusyCode || !errors.Is(busy, ErrSourceTreeBusy) {
		t.Fatalf("source-tree busy classification = %q, %v", busy.Code(), busy)
	}
}

func TestSourceTreeSharedIntentConflictsWithExclusiveAncestor(t *testing.T) {
	manager, err := NewManager(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	parent, err := NewSourceTreeIdentity(t.TempDir(), "src")
	if err != nil {
		t.Fatal(err)
	}
	owner, err := manager.Acquire(t.Context(), AcquireRequest{
		Identity: parent, Command: "pull", OperationKind: OperationMutate,
		ResourceScope: ResourceSourceTree,
	})
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = owner.Release() }()

	child, err := NewSourceTreeIdentity(filepath.Dir(parent.CanonicalPath), filepath.Join(parent.CanonicalPath, "modules"))
	if err != nil {
		t.Fatal(err)
	}
	hierarchy := SourceTreeLockHierarchy(child)
	var parentIntent ResourceIdentity
	for _, identity := range hierarchy {
		if identity.LockID == parent.LockID {
			parentIntent = identity
			break
		}
	}
	if parentIntent.LockID == "" {
		t.Fatal("child hierarchy does not include parent")
	}
	_, err = manager.Acquire(t.Context(), AcquireRequest{
		Identity: parentIntent, Command: "pull", OperationKind: OperationMutate,
		ResourceScope: ResourceSourceTree, Shared: true,
	})
	if !errors.Is(err, ErrSourceTreeBusy) {
		t.Fatalf("shared child intent error = %v, want source-tree busy", err)
	}
}

func TestSourceTreeSharedIntentsAllowDisjointSiblings(t *testing.T) {
	manager, err := NewManager(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	parent, err := NewSourceTreeIdentity(t.TempDir(), "src")
	if err != nil {
		t.Fatal(err)
	}
	first, err := manager.Acquire(t.Context(), AcquireRequest{
		Identity: parent, Command: "pull", OperationKind: OperationMutate,
		ResourceScope: ResourceSourceTree, Shared: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = first.Release() }()
	second, err := manager.Acquire(t.Context(), AcquireRequest{
		Identity: parent, Command: "pull", OperationKind: OperationMutate,
		ResourceScope: ResourceSourceTree, Shared: true,
	})
	if err != nil {
		t.Fatalf("second shared intent contended: %v", err)
	}
	defer func() { _ = second.Release() }()
}
