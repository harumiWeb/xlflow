package filepush

import "errors"

var (
	// ErrDuplicateModule reports case-insensitive VBA component name collisions
	// across the managed source roots. The .NET Excel push fails with the
	// duplicate_module_name contract before touching Excel; the file backend
	// detects the same conflicts before reading the workbook.
	ErrDuplicateModule = errors.New("file push: duplicate VBA module names")

	// ErrLineNumberSafety reports a source file that Erl instrumentation cannot
	// transform safely (existing numeric labels or numeric GoTo/GoSub/Resume
	// targets). The workbook is never touched when this error is returned.
	ErrLineNumberSafety = errors.New("file push: VBA line number safety check failed")

	// ErrBackup reports a failure while creating the pre-push workbook backup.
	// The workbook is left unchanged when this error is returned.
	ErrBackup = errors.New("file push: backup creation failed")

	// ErrPublish reports that the generated workbook could not be staged,
	// validated, or atomically published. The existing workbook is left
	// unchanged when this error is returned.
	ErrPublish = errors.New("file push: workbook publication failed")

	// ErrWorkbookLease reports that the publish window's workbook lease could
	// not be acquired. A typed coordination.BusyError stays reachable through
	// errors.As so callers can distinguish contention from I/O failures.
	ErrWorkbookLease = errors.New("file push: workbook lease acquisition failed")

	// ErrRecoveryCheck reports that the workbook's recovery state could not be
	// confirmed safe under the lease. A typed
	// coordination.RecoveryRequiredError stays reachable through errors.As.
	ErrRecoveryCheck = errors.New("file push: workbook recovery check failed")
)
