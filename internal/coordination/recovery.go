package coordination

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"time"
)

const (
	WorkbookRecoveryRequiredCode           = "workbook_recovery_required"
	WorkbookRecoveryVerificationFailedCode = "workbook_recovery_verification_failed"
	WorkbookRecoveryPublicationFailedCode  = "workbook_recovery_publication_failed"
	RecoveryCheckFailedCode                = "coordination_recovery_check_failed"
	recoverySchemaV1                       = 1
	recoveryFileSuffix                     = ".recovery.json"
)

type RecoverySession struct {
	Active bool   `json:"active"`
	Owner  string `json:"owner,omitempty"`
}

type RecoveryMetadata struct {
	SchemaVersion int             `json:"schema_version"`
	Generation    string          `json:"generation"`
	Workbook      string          `json:"workbook"`
	Reason        string          `json:"reason"`
	Operation     CommandID       `json:"operation"`
	XlflowPID     int             `json:"xlflow_pid"`
	RecordedAt    time.Time       `json:"recorded_at"`
	Session       RecoverySession `json:"session"`
	ExcelPID      int             `json:"excel_pid,omitempty"`
	WorkerPID     int             `json:"worker_pid,omitempty"`
}

type RecoveryPublication struct {
	Reason    string
	Operation CommandID
	Session   RecoverySession
	ExcelPID  int
	WorkerPID int
}

type RecoveryState struct {
	Required bool
	Invalid  bool
	Metadata *RecoveryMetadata
}

func (s RecoveryState) Reason() string {
	if s.Invalid {
		return "recovery_metadata_invalid"
	}
	if s.Metadata == nil {
		return ""
	}
	return s.Metadata.Reason
}

func RecoveryDetails(identity WorkbookIdentity, state RecoveryState) map[string]any {
	details := map[string]any{
		"workbook":          identity.CanonicalPath,
		"reason":            state.Reason(),
		"retryable":         false,
		"wait_will_resolve": false,
		"recovery_actions":  RecoveryActions(state),
	}
	if metadata := state.Metadata; metadata != nil {
		details["operation"] = metadata.Operation
		details["recorded_at"] = metadata.RecordedAt.UTC().Format(time.RFC3339Nano)
		if metadata.ExcelPID > 0 {
			details["excel_pid"] = metadata.ExcelPID
		}
		if metadata.Session.Active || meaningfulRecoveryOwner(metadata.Session.Owner) {
			details["session"] = map[string]any{
				"active": metadata.Session.Active,
				"owner":  metadata.Session.Owner,
			}
		}
	}
	return details
}

func RecoveryActions(state RecoveryState) []string {
	actions := make([]string, 0, 3)
	if metadata := state.Metadata; metadata != nil {
		if strings.EqualFold(metadata.Session.Owner, "external") {
			actions = append(actions, "close the workbook in Excel without saving")
		} else if metadata.Session.Active {
			actions = append(actions, "xlflow session stop --discard")
		}
	}
	if metadata := state.Metadata; metadata != nil && metadata.ExcelPID > 0 {
		actions = append(actions, fmt.Sprintf("xlflow process cleanup %d", metadata.ExcelPID))
	}
	actions = append(actions, "xlflow recovery clear", "xlflow recovery clear --force")
	return actions
}

type RecoveryRequiredError struct {
	Identity WorkbookIdentity
	State    RecoveryState
}

func (e *RecoveryRequiredError) Error() string {
	return fmt.Sprintf("%s: workbook %q requires explicit recovery", WorkbookRecoveryRequiredCode, e.Identity.CanonicalPath)
}

func (e *RecoveryRequiredError) Code() string { return WorkbookRecoveryRequiredCode }

type RecoveryEntry struct {
	LockID string
	State  RecoveryState
}

type Observation struct {
	Busy     bool
	Owner    *OwnerMetadata
	Recovery RecoveryState
}

func (m *Manager) RecoveryState(identity WorkbookIdentity) (RecoveryState, error) {
	if err := validateIdentity(identity); err != nil {
		return RecoveryState{}, err
	}
	return m.readRecovery(identity)
}

func (m *Manager) ListRecoveries() ([]RecoveryEntry, error) {
	entries, err := os.ReadDir(m.dir)
	if err != nil {
		return nil, fmt.Errorf("list recovery metadata: %w", err)
	}
	result := make([]RecoveryEntry, 0)
	for _, entry := range entries {
		if entry.IsDir() || !strings.HasSuffix(entry.Name(), recoveryFileSuffix) {
			continue
		}
		lockID := strings.TrimSuffix(entry.Name(), recoveryFileSuffix)
		state := RecoveryState{Required: true, Invalid: true}
		data, readErr := os.ReadFile(filepath.Join(m.dir, entry.Name()))
		if readErr != nil {
			return nil, fmt.Errorf("read recovery metadata %q: %w", entry.Name(), readErr)
		}
		var metadata RecoveryMetadata
		if json.Unmarshal(data, &metadata) == nil && validRecoveryMetadata(lockID, metadata) {
			identity, identityErr := NewWorkbookIdentity(filepath.Dir(metadata.Workbook), metadata.Workbook)
			if identityErr == nil && identity.LockID == lockID {
				state = RecoveryState{Required: true, Metadata: &metadata}
			}
		}
		result = append(result, RecoveryEntry{LockID: lockID, State: state})
	}
	sort.Slice(result, func(i, j int) bool { return result[i].LockID < result[j].LockID })
	return result, nil
}

func (m *Manager) Observe(ctx context.Context, identity WorkbookIdentity) (Observation, error) {
	if err := validateIdentity(identity); err != nil {
		return Observation{}, err
	}
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return Observation{}, err
	}
	file, err := m.openLock(identity)
	if err != nil {
		return Observation{}, err
	}
	defer func() { _ = file.Close() }()

	if err := m.lockContext(ctx, file, publicationByte); err != nil {
		return Observation{}, err
	}
	defer func() { _ = platformUnlock(file, publicationByte) }()

	acquired, err := platformTryLock(file, operationByte)
	if err != nil {
		return Observation{}, fmt.Errorf("observe workbook lock: %w", err)
	}
	observation := Observation{}
	if acquired {
		defer func() { _ = platformUnlock(file, operationByte) }()
		_ = m.removeOwnerIfMatches(identity, "")
	} else {
		observation.Busy = true
		observation.Owner = m.readOwnerBestEffort(identity)
	}
	observation.Recovery, err = m.readRecovery(identity)
	if err != nil {
		return Observation{}, err
	}
	return observation, nil
}

func (l *Lease) RecoveryState() (RecoveryState, error) {
	if l == nil || l.manager == nil || l.file == nil {
		return RecoveryState{}, fmt.Errorf("workbook lease is required")
	}
	var state RecoveryState
	err := l.withPublicationGuard(func() error {
		var err error
		state, err = l.manager.readRecovery(l.identity)
		return err
	})
	return state, err
}

func (l *Lease) RequireRecoveryAllowed(behavior RecoveryBehavior, recoveryIntent bool) error {
	if behavior == RecoveryNotApplicable || behavior == RecoveryObserve {
		return nil
	}
	state, err := l.RecoveryState()
	if err != nil {
		return err
	}
	if !state.Required {
		return nil
	}
	if behavior == RecoveryRecover && recoveryIntent {
		return nil
	}
	return &RecoveryRequiredError{Identity: l.identity, State: state}
}

func (l *Lease) PublishRecovery(publication RecoveryPublication) (*RecoveryMetadata, error) {
	if l == nil || l.manager == nil || l.file == nil {
		return nil, fmt.Errorf("workbook lease is required")
	}
	if strings.TrimSpace(publication.Reason) == "" {
		return nil, fmt.Errorf("recovery reason is required")
	}
	if publication.Operation == "" {
		return nil, fmt.Errorf("recovery operation is required")
	}
	generation, err := newGeneration()
	if err != nil {
		return nil, err
	}
	metadata := &RecoveryMetadata{
		SchemaVersion: recoverySchemaV1,
		Generation:    generation,
		Workbook:      l.identity.CanonicalPath,
		Reason:        strings.TrimSpace(publication.Reason),
		Operation:     publication.Operation,
		XlflowPID:     os.Getpid(),
		RecordedAt:    time.Now().UTC(),
		Session:       publication.Session,
		ExcelPID:      publication.ExcelPID,
		WorkerPID:     publication.WorkerPID,
	}
	if err := l.withPublicationGuard(func() error {
		return writeJSONAtomic(l.manager.dir, l.manager.recoveryPath(l.identity), l.identity.LockID+".recovery-", metadata)
	}); err != nil {
		return nil, err
	}
	return metadata, nil
}

func (l *Lease) ClearRecovery(expectedGeneration string) (bool, error) {
	if l == nil || l.manager == nil || l.file == nil {
		return false, fmt.Errorf("workbook lease is required")
	}
	cleared := false
	err := l.withPublicationGuard(func() error {
		state, err := l.manager.readRecovery(l.identity)
		if err != nil {
			return err
		}
		if !state.Required {
			return nil
		}
		if expectedGeneration != "" {
			if state.Metadata == nil || state.Metadata.Generation != expectedGeneration {
				return nil
			}
		}
		if err := os.Remove(l.manager.recoveryPath(l.identity)); err != nil && !errors.Is(err, os.ErrNotExist) {
			return fmt.Errorf("remove recovery metadata: %w", err)
		}
		cleared = true
		return nil
	})
	return cleared, err
}

func (l *Lease) withPublicationGuard(run func() error) error {
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	if err := l.manager.lockContext(ctx, l.file, publicationByte); err != nil {
		return fmt.Errorf("acquire recovery publication guard: %w", err)
	}
	runErr := run()
	unlockErr := platformUnlock(l.file, publicationByte)
	return errors.Join(runErr, unlockErr)
}

func (m *Manager) recoveryPath(identity WorkbookIdentity) string {
	return filepath.Join(m.dir, identity.LockID+recoveryFileSuffix)
}

func (m *Manager) readRecovery(identity WorkbookIdentity) (RecoveryState, error) {
	data, err := os.ReadFile(m.recoveryPath(identity))
	if errors.Is(err, os.ErrNotExist) {
		return RecoveryState{}, nil
	}
	if err != nil {
		return RecoveryState{}, fmt.Errorf("read recovery metadata: %w", err)
	}
	var metadata RecoveryMetadata
	if json.Unmarshal(data, &metadata) != nil || !validRecoveryMetadata(identity.LockID, metadata) ||
		!SamePath(metadata.Workbook, identity.CanonicalPath) {
		return RecoveryState{Required: true, Invalid: true}, nil
	}
	return RecoveryState{Required: true, Metadata: &metadata}, nil
}

func validRecoveryMetadata(lockID string, metadata RecoveryMetadata) bool {
	validOwner := metadata.Session.Owner == "" ||
		strings.EqualFold(metadata.Session.Owner, "none") ||
		strings.EqualFold(metadata.Session.Owner, "managed") ||
		strings.EqualFold(metadata.Session.Owner, "external")
	validSession := validOwner
	if metadata.Session.Active {
		validSession = strings.EqualFold(metadata.Session.Owner, "managed") ||
			strings.EqualFold(metadata.Session.Owner, "external")
	}
	return validLockID.MatchString(lockID) &&
		metadata.SchemaVersion == recoverySchemaV1 &&
		validGeneration.MatchString(metadata.Generation) &&
		strings.TrimSpace(metadata.Workbook) != "" &&
		strings.TrimSpace(metadata.Reason) != "" &&
		metadata.Operation != "" &&
		metadata.XlflowPID > 0 &&
		!metadata.RecordedAt.IsZero() &&
		validSession &&
		metadata.ExcelPID >= 0 &&
		metadata.WorkerPID >= 0
}

func meaningfulRecoveryOwner(owner string) bool {
	return strings.TrimSpace(owner) != "" && !strings.EqualFold(owner, "none")
}
