package coordination

import (
	"encoding/json"
	"os"
	"path/filepath"
	"runtime"
	"testing"
	"time"
)

func TestRecoveryStateRoundTripAndInvalidMetadataFailsClosed(t *testing.T) {
	manager, err := NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	identity, err := NewWorkbookIdentity(t.TempDir(), "Book.xlsm")
	if err != nil {
		t.Fatal(err)
	}
	metadata := RecoveryMetadata{
		SchemaVersion: recoverySchemaV1,
		Generation:    "0123456789abcdef0123456789abcdef",
		Workbook:      identity.CanonicalPath,
		Reason:        "vba_may_still_be_running",
		Operation:     "run",
		XlflowPID:     1234,
		RecordedAt:    time.Date(2026, 7, 16, 9, 30, 0, 0, time.UTC),
		Session:       RecoverySession{Active: true, Owner: "managed"},
		ExcelPID:      2345,
		WorkerPID:     3456,
	}
	if err := writeJSONAtomic(manager.dir, manager.recoveryPath(identity), identity.LockID+".test-", metadata); err != nil {
		t.Fatal(err)
	}

	state, err := manager.RecoveryState(identity)
	if err != nil {
		t.Fatal(err)
	}
	if !state.Required || state.Invalid || state.Metadata == nil {
		t.Fatalf("state = %+v", state)
	}
	if state.Metadata.Reason != metadata.Reason || state.Metadata.ExcelPID != metadata.ExcelPID {
		t.Fatalf("metadata = %+v", state.Metadata)
	}

	if err := os.WriteFile(manager.recoveryPath(identity), []byte(`{"schema_version":999}`), 0o600); err != nil {
		t.Fatal(err)
	}
	state, err = manager.RecoveryState(identity)
	if err != nil {
		t.Fatal(err)
	}
	if !state.Required || !state.Invalid || state.Reason() != "recovery_metadata_invalid" {
		t.Fatalf("invalid state = %+v", state)
	}
}

func TestRecoveryStateRejectsWorkbookMismatch(t *testing.T) {
	manager, err := NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	root := t.TempDir()
	identity, _ := NewWorkbookIdentity(root, "Book.xlsm")
	other, _ := NewWorkbookIdentity(root, "Other.xlsm")
	metadata := RecoveryMetadata{
		SchemaVersion: recoverySchemaV1,
		Generation:    "0123456789abcdef0123456789abcdef",
		Workbook:      other.CanonicalPath,
		Reason:        "excel_com_state_uncertain",
		Operation:     "run",
		XlflowPID:     1234,
		RecordedAt:    time.Now().UTC(),
	}
	body, _ := json.Marshal(metadata)
	if err := os.WriteFile(manager.recoveryPath(identity), body, 0o600); err != nil {
		t.Fatal(err)
	}
	state, err := manager.RecoveryState(identity)
	if err != nil {
		t.Fatal(err)
	}
	if !state.Required || !state.Invalid {
		t.Fatalf("state = %+v", state)
	}
}

func TestListRecoveriesRejectsFilenameWorkbookIdentityMismatch(t *testing.T) {
	manager, err := NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	root := t.TempDir()
	identity, _ := NewWorkbookIdentity(root, "Book.xlsm")
	other, _ := NewWorkbookIdentity(root, "Other.xlsm")
	metadata := RecoveryMetadata{
		SchemaVersion: recoverySchemaV1,
		Generation:    "0123456789abcdef0123456789abcdef",
		Workbook:      other.CanonicalPath,
		Reason:        "excel_com_state_uncertain",
		Operation:     "run",
		XlflowPID:     1234,
		RecordedAt:    time.Now().UTC(),
	}
	body, _ := json.Marshal(metadata)
	if err := os.WriteFile(manager.recoveryPath(identity), body, 0o600); err != nil {
		t.Fatal(err)
	}
	entries, err := manager.ListRecoveries()
	if err != nil {
		t.Fatal(err)
	}
	if len(entries) != 1 || !entries[0].State.Required || !entries[0].State.Invalid || entries[0].State.Metadata != nil {
		t.Fatalf("entries = %+v", entries)
	}
}

func TestRecoveryStateRejectsInvalidSessionOwner(t *testing.T) {
	manager, err := NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	identity, _ := NewWorkbookIdentity(t.TempDir(), "Book.xlsm")
	metadata := RecoveryMetadata{
		SchemaVersion: recoverySchemaV1,
		Generation:    "0123456789abcdef0123456789abcdef",
		Workbook:      identity.CanonicalPath,
		Reason:        "excel_com_state_uncertain",
		Operation:     "run",
		XlflowPID:     1234,
		RecordedAt:    time.Now().UTC(),
		Session:       RecoverySession{Active: true, Owner: "unknown"},
	}
	body, _ := json.Marshal(metadata)
	if err := os.WriteFile(manager.recoveryPath(identity), body, 0o600); err != nil {
		t.Fatal(err)
	}
	state, err := manager.RecoveryState(identity)
	if err != nil {
		t.Fatal(err)
	}
	if !state.Required || !state.Invalid {
		t.Fatalf("state = %+v", state)
	}
}

func TestRecoveryStatePreservesDetachedExternalOwner(t *testing.T) {
	manager, err := NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	identity, _ := NewWorkbookIdentity(t.TempDir(), "Book.xlsm")
	metadata := RecoveryMetadata{
		SchemaVersion: recoverySchemaV1,
		Generation:    "0123456789abcdef0123456789abcdef",
		Workbook:      identity.CanonicalPath,
		Reason:        "external_session_detached",
		Operation:     "session.stop",
		XlflowPID:     1234,
		RecordedAt:    time.Now().UTC(),
		Session:       RecoverySession{Active: false, Owner: "external"},
		ExcelPID:      24680,
	}
	body, _ := json.Marshal(metadata)
	if err := os.WriteFile(manager.recoveryPath(identity), body, 0o600); err != nil {
		t.Fatal(err)
	}
	state, err := manager.RecoveryState(identity)
	if err != nil {
		t.Fatal(err)
	}
	if !state.Required || state.Invalid || state.Metadata == nil || state.Metadata.Session.Owner != "external" {
		t.Fatalf("state = %+v", state)
	}
	actions := RecoveryActions(state)
	if len(actions) == 0 || actions[0] != "close the workbook in Excel without saving" {
		t.Fatalf("actions = %#v", actions)
	}
	details := RecoveryDetails(identity, state)
	session, ok := details["session"].(map[string]any)
	if !ok || session["active"] != false || session["owner"] != "external" {
		t.Fatalf("details = %#v", details)
	}
}

func TestClearRecoveryRequiresMatchingGeneration(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	manager, err := NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	identity, _ := NewWorkbookIdentity(t.TempDir(), "Book.xlsm")
	lease, err := manager.Acquire(t.Context(), AcquireRequest{
		Identity:      identity,
		Command:       "run",
		OperationKind: OperationExecute,
		ResourceScope: ResourceWorkbook,
	})
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = lease.Release() }()
	metadata, err := lease.PublishRecovery(RecoveryPublication{
		Reason:    "vba_may_still_be_running",
		Operation: "run",
	})
	if err != nil {
		t.Fatal(err)
	}
	cleared, err := lease.ClearRecovery("ffffffffffffffffffffffffffffffff")
	if err != nil {
		t.Fatal(err)
	}
	if cleared {
		t.Fatal("mismatched generation cleared recovery")
	}
	state, err := lease.RecoveryState()
	if err != nil {
		t.Fatal(err)
	}
	if !state.Required || state.Metadata == nil || state.Metadata.Generation != metadata.Generation {
		t.Fatalf("state = %+v", state)
	}
	cleared, err = lease.ClearRecovery(metadata.Generation)
	if err != nil || !cleared {
		t.Fatalf("matching generation clear = %v, %v", cleared, err)
	}
}

func TestListRecoveriesIsStableAndIncludesInvalidEntries(t *testing.T) {
	manager, err := NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	root := t.TempDir()
	first, _ := NewWorkbookIdentity(root, "A.xlsm")
	second, _ := NewWorkbookIdentity(root, "B.xlsm")
	for _, identity := range []WorkbookIdentity{second, first} {
		metadata := RecoveryMetadata{
			SchemaVersion: recoverySchemaV1,
			Generation:    "0123456789abcdef0123456789abcdef",
			Workbook:      identity.CanonicalPath,
			Reason:        "vba_may_still_be_running",
			Operation:     "run",
			XlflowPID:     1234,
			RecordedAt:    time.Now().UTC(),
		}
		if err := writeJSONAtomic(manager.dir, manager.recoveryPath(identity), identity.LockID+".test-", metadata); err != nil {
			t.Fatal(err)
		}
	}
	invalidLockID := workbookLockIDPrefix + "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
	if err := os.WriteFile(filepath.Join(manager.dir, invalidLockID+recoveryFileSuffix), []byte(`{`), 0o600); err != nil {
		t.Fatal(err)
	}
	entries, err := manager.ListRecoveries()
	if err != nil {
		t.Fatal(err)
	}
	if len(entries) != 3 {
		t.Fatalf("entries = %+v", entries)
	}
	for i := 1; i < len(entries); i++ {
		if entries[i-1].LockID > entries[i].LockID {
			t.Fatalf("entries are not sorted: %+v", entries)
		}
	}
	foundInvalid := false
	for _, entry := range entries {
		foundInvalid = foundInvalid || entry.State.Invalid
	}
	if !foundInvalid {
		t.Fatalf("invalid recovery entry was not reported: %+v", entries)
	}
}

func TestRecoveryMarkersDoNotAffectWorkbookOwnershipProbe(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Windows LockFileEx coordination")
	}
	manager, err := NewManager(filepath.Join(t.TempDir(), "coordination"))
	if err != nil {
		t.Fatal(err)
	}
	identity, _ := NewWorkbookIdentity(t.TempDir(), "Book.xlsm")
	metadata := RecoveryMetadata{
		SchemaVersion: recoverySchemaV1,
		Generation:    "0123456789abcdef0123456789abcdef",
		Workbook:      identity.CanonicalPath,
		Reason:        "vba_may_still_be_running",
		Operation:     "run",
		XlflowPID:     1234,
		RecordedAt:    time.Now().UTC(),
	}
	if err := writeJSONAtomic(manager.dir, manager.recoveryPath(identity), identity.LockID+".test-", metadata); err != nil {
		t.Fatal(err)
	}
	result, err := manager.Probe(t.Context(), identity)
	if err != nil {
		t.Fatal(err)
	}
	if result.Busy {
		t.Fatal("recovery marker was treated as active ownership")
	}
}
