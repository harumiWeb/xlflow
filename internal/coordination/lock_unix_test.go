//go:build unix

package coordination

import (
	"errors"
	"os"
	"os/exec"
	"path/filepath"
	"testing"
)

func TestUnixSourceTreeLockSerializesProcesses(t *testing.T) {
	stateDir := os.Getenv("XLFLOW_TEST_SOURCE_LOCK_STATE")
	resource := os.Getenv("XLFLOW_TEST_SOURCE_LOCK_RESOURCE")
	if stateDir != "" {
		manager, err := NewManager(stateDir)
		if err != nil {
			t.Fatal(err)
		}
		identity, err := NewSourceTreeIdentity(filepath.Dir(resource), resource)
		if err != nil {
			t.Fatal(err)
		}
		_, err = manager.Acquire(t.Context(), AcquireRequest{
			Identity: identity, Command: "pull", OperationKind: OperationMutate,
			ResourceScope: ResourceSourceTree,
		})
		if !errors.Is(err, ErrSourceTreeBusy) {
			t.Fatalf("child acquisition error = %v, want source-tree busy", err)
		}
		return
	}

	stateDir = t.TempDir()
	resource = filepath.Join(t.TempDir(), "src", "modules")
	manager, err := NewManager(stateDir)
	if err != nil {
		t.Fatal(err)
	}
	identity, err := NewSourceTreeIdentity(filepath.Dir(resource), resource)
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

	cmd := exec.CommandContext(t.Context(), os.Args[0], "-test.run=^TestUnixSourceTreeLockSerializesProcesses$")
	cmd.Env = append(os.Environ(), "XLFLOW_TEST_SOURCE_LOCK_STATE="+stateDir, "XLFLOW_TEST_SOURCE_LOCK_RESOURCE="+resource)
	if output, err := cmd.CombinedOutput(); err != nil {
		t.Fatalf("child lock probe failed: %v\n%s", err, output)
	}
}

func TestUnixSourceTreeLockSerializesManagersInSameProcess(t *testing.T) {
	stateDir := t.TempDir()
	resource := filepath.Join(t.TempDir(), "src", "modules")
	firstManager, err := NewManager(stateDir)
	if err != nil {
		t.Fatal(err)
	}
	secondManager, err := NewManager(stateDir)
	if err != nil {
		t.Fatal(err)
	}
	identity, err := NewSourceTreeIdentity(filepath.Dir(resource), resource)
	if err != nil {
		t.Fatal(err)
	}
	lease, err := firstManager.Acquire(t.Context(), AcquireRequest{
		Identity: identity, Command: "pull", OperationKind: OperationMutate,
		ResourceScope: ResourceSourceTree,
	})
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = lease.Release() }()

	_, err = secondManager.Acquire(t.Context(), AcquireRequest{
		Identity: identity, Command: "pull", OperationKind: OperationMutate,
		ResourceScope: ResourceSourceTree,
	})
	if !errors.Is(err, ErrSourceTreeBusy) {
		t.Fatalf("same-process acquisition error = %v, want source-tree busy", err)
	}
	result, err := secondManager.Probe(t.Context(), identity)
	if err != nil {
		t.Fatal(err)
	}
	if !result.Busy {
		t.Fatal("failed same-process acquisition cleared the live lease")
	}
}
