package coordination

import (
	"context"
	"crypto/rand"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"sync"
	"time"
)

const (
	WorkbookBusyCode          = "workbook_busy"
	WorkbookBusyTimeoutCode   = "workbook_busy_timeout"
	WorkbookBusyCancelledCode = "workbook_busy_cancelled"
	ownerSchemaV1             = 1
	operationByte             = int64(0)
	publicationByte           = int64(1)
)

var (
	ErrWorkbookBusy = errors.New(WorkbookBusyCode)
	validLockID     = regexp.MustCompile(`^[a-z0-9-]+$`)
	validGeneration = regexp.MustCompile(`^[0-9a-f]{32}$`)
)

// OwnerMetadata describes the xlflow process that currently owns a workbook
// lease. It is diagnostic only: the operating-system byte-range lock remains
// the authority for ownership.
type OwnerMetadata struct {
	SchemaVersion int           `json:"schema_version"`
	Generation    string        `json:"generation"`
	Workbook      string        `json:"workbook"`
	PID           int           `json:"pid"`
	Command       CommandID     `json:"command"`
	OperationKind OperationKind `json:"operation_kind"`
	ResourceScope ResourceScope `json:"resource_scope"`
	StartedAt     time.Time     `json:"started_at"`
}

// AcquireRequest identifies the workbook operation to coordinate. Wait=false
// performs one authoritative acquisition attempt. Wait=true retries until the
// context is cancelled.
type AcquireRequest struct {
	Identity      WorkbookIdentity
	Command       CommandID
	OperationKind OperationKind
	ResourceScope ResourceScope
	Wait          bool
}

// BusyError reports authoritative lock contention. Owner may be nil when the
// owner crashed while publishing metadata or the metadata is unavailable.
type BusyError struct {
	Identity WorkbookIdentity
	Owner    *OwnerMetadata
}

func (e *BusyError) Error() string {
	return fmt.Sprintf("%s: workbook %q is in use", WorkbookBusyCode, e.Identity.CanonicalPath)
}

func (e *BusyError) Unwrap() error { return ErrWorkbookBusy }
func (e *BusyError) Code() string  { return WorkbookBusyCode }

// ProbeResult is a point-in-time view of authoritative workbook ownership.
type ProbeResult struct {
	Busy  bool
	Owner *OwnerMetadata
}

// Manager owns the filesystem namespace used for OS lock handles and
// diagnostic ownership metadata.
type Manager struct {
	dir          string
	pollInterval time.Duration
}

// NewManager creates a manager rooted at stateDir. A single shared directory
// must be used by all xlflow processes that need to coordinate.
func NewManager(stateDir string) (*Manager, error) {
	if stateDir == "" {
		return nil, fmt.Errorf("coordination state directory is required")
	}
	abs, err := filepath.Abs(stateDir)
	if err != nil {
		return nil, fmt.Errorf("resolve coordination state directory: %w", err)
	}
	if err := os.MkdirAll(abs, 0o700); err != nil {
		return nil, fmt.Errorf("create coordination state directory: %w", err)
	}
	return &Manager{dir: abs, pollInterval: 25 * time.Millisecond}, nil
}

// NewDefaultManager creates the per-user coordination manager.
func NewDefaultManager() (*Manager, error) {
	base, err := os.UserCacheDir()
	if err != nil {
		return nil, fmt.Errorf("locate user cache directory: %w", err)
	}
	return NewManager(filepath.Join(base, "xlflow", "coordination"))
}

// StateDir returns the shared state directory used by the manager.
func (m *Manager) StateDir() string { return m.dir }

// Acquire obtains the authoritative exclusive workbook lease and publishes
// ownership metadata before returning it to the caller.
func (m *Manager) Acquire(ctx context.Context, req AcquireRequest) (*Lease, error) {
	if err := validateRequest(req); err != nil {
		return nil, err
	}
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}

	file, err := m.openLock(req.Identity)
	if err != nil {
		return nil, err
	}
	keepFile := false
	defer func() {
		if !keepFile {
			_ = file.Close()
		}
	}()

	for {
		// Observe the operation byte before entering the publication handshake.
		// A free observation is released immediately; authoritative acquisition
		// is then repeated while holding the publication guard so readers can
		// never see a new owner paired with stale metadata.
		observedFree, observeErr := platformTryLock(file, operationByte)
		if observeErr != nil {
			return nil, fmt.Errorf("observe workbook lock: %w", observeErr)
		}
		if observedFree {
			if err := platformUnlock(file, operationByte); err != nil {
				return nil, fmt.Errorf("release workbook lock observation: %w", err)
			}
		}
		if req.Wait {
			if err := m.lockContext(ctx, file, publicationByte); err != nil {
				return nil, err
			}
		} else {
			guardAcquired, guardErr := platformTryLock(file, publicationByte)
			if guardErr != nil {
				return nil, fmt.Errorf("acquire metadata publication guard: %w", guardErr)
			}
			if !guardAcquired {
				return nil, &BusyError{Identity: req.Identity}
			}
		}

		acquired, lockErr := platformTryLock(file, operationByte)
		if lockErr != nil {
			_ = platformUnlock(file, publicationByte)
			return nil, fmt.Errorf("acquire workbook lock: %w", lockErr)
		}
		if acquired {
			owner, metadataErr := newOwnerMetadata(req)
			if metadataErr == nil {
				metadataErr = m.writeOwner(req.Identity, owner)
			}
			guardErr := platformUnlock(file, publicationByte)
			if metadataErr != nil || guardErr != nil {
				_ = platformUnlock(file, operationByte)
				return nil, errors.Join(metadataErr, guardErr)
			}
			keepFile = true
			return &Lease{manager: m, file: file, identity: req.Identity, generation: owner.Generation}, nil
		}

		owner := m.readOwnerBestEffort(req.Identity)
		guardErr := platformUnlock(file, publicationByte)
		if guardErr != nil {
			return nil, fmt.Errorf("release metadata publication guard: %w", guardErr)
		}
		if !req.Wait {
			return nil, &BusyError{Identity: req.Identity, Owner: owner}
		}
		if err := waitContext(ctx, m.pollInterval); err != nil {
			return nil, err
		}
	}
}

// Probe reports whether identity is currently locked. Stale metadata is
// removed only after successfully acquiring the authoritative operation byte.
func (m *Manager) Probe(ctx context.Context, identity WorkbookIdentity) (ProbeResult, error) {
	if err := validateIdentity(identity); err != nil {
		return ProbeResult{}, err
	}
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return ProbeResult{}, err
	}
	file, err := m.openLock(identity)
	if err != nil {
		return ProbeResult{}, err
	}
	defer func() { _ = file.Close() }()

	if err := m.lockContext(ctx, file, publicationByte); err != nil {
		return ProbeResult{}, err
	}
	defer func() { _ = platformUnlock(file, publicationByte) }()

	acquired, err := platformTryLock(file, operationByte)
	if err != nil {
		return ProbeResult{}, fmt.Errorf("probe workbook lock: %w", err)
	}
	if !acquired {
		return ProbeResult{Busy: true, Owner: m.readOwnerBestEffort(identity)}, nil
	}
	defer func() { _ = platformUnlock(file, operationByte) }()
	_ = m.removeOwnerIfMatches(identity, "")
	return ProbeResult{}, nil
}

// ReadCurrentOwner returns ownership metadata only when the authoritative lock
// is currently held. Metadata left behind by a crashed process is not reported.
func (m *Manager) ReadCurrentOwner(ctx context.Context, identity WorkbookIdentity) (*OwnerMetadata, error) {
	result, err := m.Probe(ctx, identity)
	if err != nil {
		return nil, err
	}
	if !result.Busy {
		return nil, nil
	}
	return result.Owner, nil
}

// Lease is an acquired workbook operation lock. Release is idempotent.
type Lease struct {
	manager    *Manager
	file       *os.File
	identity   WorkbookIdentity
	generation string
	once       sync.Once
	err        error
}

func (l *Lease) Identity() WorkbookIdentity { return l.identity }

func (l *Lease) Release() error {
	if l == nil {
		return nil
	}
	l.once.Do(func() { l.err = l.release() })
	return l.err
}

func (l *Lease) release() error {
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	guardErr := l.manager.lockContext(ctx, l.file, publicationByte)
	if guardErr != nil {
		unlockErr := platformUnlock(l.file, operationByte)
		closeErr := l.file.Close()
		return errors.Join(guardErr, unlockErr, closeErr)
	}
	// Metadata is diagnostic and may be stale after a crash. Failure to clean it
	// must not turn a successfully completed workbook operation into a failure;
	// the next acquisition atomically replaces it and Probe removes it when free.
	_ = l.manager.removeOwnerIfMatches(l.identity, l.generation)
	unlockErr := platformUnlock(l.file, operationByte)
	guardReleaseErr := platformUnlock(l.file, publicationByte)
	closeErr := l.file.Close()
	return errors.Join(unlockErr, guardReleaseErr, closeErr)
}

func validateRequest(req AcquireRequest) error {
	if err := validateIdentity(req.Identity); err != nil {
		return err
	}
	if req.Command == "" {
		return fmt.Errorf("coordination command is required")
	}
	if !req.OperationKind.Valid() {
		return fmt.Errorf("invalid operation kind %q", req.OperationKind)
	}
	if req.ResourceScope != ResourceWorkbook {
		return fmt.Errorf("workbook lock requires resource scope %q, got %q", ResourceWorkbook, req.ResourceScope)
	}
	return nil
}

func validateIdentity(identity WorkbookIdentity) error {
	if identity.CanonicalPath == "" {
		return fmt.Errorf("canonical workbook path is required")
	}
	if identity.LockID == "" || !validLockID.MatchString(identity.LockID) {
		return fmt.Errorf("invalid workbook lock ID %q", identity.LockID)
	}
	return nil
}

func (m *Manager) openLock(identity WorkbookIdentity) (*os.File, error) {
	path := filepath.Join(m.dir, identity.LockID+".lock")
	file, err := os.OpenFile(path, os.O_CREATE|os.O_RDWR, 0o600)
	if err != nil {
		return nil, fmt.Errorf("open workbook lock file: %w", err)
	}
	return file, nil
}

func (m *Manager) lockContext(ctx context.Context, file *os.File, offset int64) error {
	for {
		if err := ctx.Err(); err != nil {
			return err
		}
		acquired, err := platformTryLock(file, offset)
		if err != nil {
			return err
		}
		if acquired {
			return nil
		}
		if err := waitContext(ctx, m.pollInterval); err != nil {
			return err
		}
	}
}

func waitContext(ctx context.Context, delay time.Duration) error {
	timer := time.NewTimer(delay)
	defer timer.Stop()
	select {
	case <-ctx.Done():
		return ctx.Err()
	case <-timer.C:
		return nil
	}
}

func newOwnerMetadata(req AcquireRequest) (*OwnerMetadata, error) {
	var token [16]byte
	if _, err := io.ReadFull(rand.Reader, token[:]); err != nil {
		return nil, fmt.Errorf("generate ownership token: %w", err)
	}
	return &OwnerMetadata{
		SchemaVersion: ownerSchemaV1,
		Generation:    hex.EncodeToString(token[:]),
		Workbook:      req.Identity.CanonicalPath,
		PID:           os.Getpid(),
		Command:       req.Command,
		OperationKind: req.OperationKind,
		ResourceScope: req.ResourceScope,
		StartedAt:     time.Now().UTC(),
	}, nil
}

func (m *Manager) ownerPath(identity WorkbookIdentity) string {
	return filepath.Join(m.dir, identity.LockID+".owner.json")
}

func (m *Manager) writeOwner(identity WorkbookIdentity, owner *OwnerMetadata) error {
	tmp, err := os.CreateTemp(m.dir, identity.LockID+".owner-*.tmp")
	if err != nil {
		return fmt.Errorf("create owner metadata: %w", err)
	}
	tmpName := tmp.Name()
	defer func() { _ = os.Remove(tmpName) }()
	if err := tmp.Chmod(0o600); err != nil {
		_ = tmp.Close()
		return fmt.Errorf("secure owner metadata: %w", err)
	}
	encoder := json.NewEncoder(tmp)
	encoder.SetEscapeHTML(false)
	if err := encoder.Encode(owner); err != nil {
		_ = tmp.Close()
		return fmt.Errorf("encode owner metadata: %w", err)
	}
	if err := tmp.Sync(); err != nil {
		_ = tmp.Close()
		return fmt.Errorf("flush owner metadata: %w", err)
	}
	if err := tmp.Close(); err != nil {
		return fmt.Errorf("close owner metadata: %w", err)
	}
	if err := platformAtomicReplace(tmpName, m.ownerPath(identity)); err != nil {
		return fmt.Errorf("publish owner metadata: %w", err)
	}
	return nil
}

func (m *Manager) readOwnerBestEffort(identity WorkbookIdentity) *OwnerMetadata {
	data, err := os.ReadFile(m.ownerPath(identity))
	if err != nil {
		return nil
	}
	var owner OwnerMetadata
	if json.Unmarshal(data, &owner) != nil ||
		owner.SchemaVersion != ownerSchemaV1 ||
		!validGeneration.MatchString(owner.Generation) ||
		!SamePath(owner.Workbook, identity.CanonicalPath) ||
		owner.PID <= 0 ||
		owner.Command == "" ||
		!owner.OperationKind.Valid() ||
		owner.ResourceScope != ResourceWorkbook ||
		owner.StartedAt.IsZero() {
		return nil
	}
	return &owner
}

// removeOwnerIfMatches removes metadata for generation. An empty generation is
// used only after proving the operation byte is free and therefore removes any
// stale metadata.
func (m *Manager) removeOwnerIfMatches(identity WorkbookIdentity, generation string) error {
	if generation != "" {
		owner := m.readOwnerBestEffort(identity)
		if owner == nil || owner.Generation != generation {
			return nil
		}
	}
	err := os.Remove(m.ownerPath(identity))
	if err != nil && !errors.Is(err, os.ErrNotExist) {
		return fmt.Errorf("remove owner metadata: %w", err)
	}
	return nil
}
