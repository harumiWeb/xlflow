package coordination

import (
	"errors"
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
