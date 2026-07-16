// Package coordination defines command policy, canonical workbook identity,
// and cross-process workbook operation locking for xlflow commands.
package coordination

import "fmt"

// CommandID is a stable, implementation-independent identifier for an xlflow
// command. IDs are intended to remain suitable for future serialization.
type CommandID string

type ResourceScope string

const (
	ResourceNone          ResourceScope = "none"
	ResourceWorkbook      ResourceScope = "workbook"
	ResourceExcelInstance ResourceScope = "excel_instance"
)

func (v ResourceScope) Valid() bool {
	switch v {
	case ResourceNone, ResourceWorkbook, ResourceExcelInstance:
		return true
	default:
		return false
	}
}

type OperationKind string

const (
	OperationRead     OperationKind = "read"
	OperationMutate   OperationKind = "mutate"
	OperationExecute  OperationKind = "execute"
	OperationDesigner OperationKind = "designer"
)

func (v OperationKind) Valid() bool {
	switch v {
	case OperationRead, OperationMutate, OperationExecute, OperationDesigner:
		return true
	default:
		return false
	}
}

type WaitPolicy string

const (
	WaitFail WaitPolicy = "fail"
	WaitWait WaitPolicy = "wait"
)

func (v WaitPolicy) Valid() bool {
	return v == WaitFail || v == WaitWait
}

type RecoveryBehavior string

const (
	RecoveryNotApplicable RecoveryBehavior = "not_applicable"
	RecoveryBlock         RecoveryBehavior = "block"
	RecoveryObserve       RecoveryBehavior = "observe"
	RecoveryRecover       RecoveryBehavior = "recover"
)

func (v RecoveryBehavior) Valid() bool {
	switch v {
	case RecoveryNotApplicable, RecoveryBlock, RecoveryObserve, RecoveryRecover:
		return true
	default:
		return false
	}
}

// Policy describes the concurrency characteristics of one command.
type Policy struct {
	ResourceScope     ResourceScope    `json:"resource_scope"`
	OperationKind     OperationKind    `json:"operation_kind"`
	ParallelSafe      bool             `json:"parallel_safe"`
	RetryableWhenBusy bool             `json:"retryable_when_busy"`
	DefaultWaitPolicy WaitPolicy       `json:"default_wait_policy"`
	RecoveryBehavior  RecoveryBehavior `json:"recovery_behavior"`
}

func (p Policy) Validate() error {
	if !p.ResourceScope.Valid() {
		return fmt.Errorf("invalid resource scope %q", p.ResourceScope)
	}
	if !p.OperationKind.Valid() {
		return fmt.Errorf("invalid operation kind %q", p.OperationKind)
	}
	if !p.DefaultWaitPolicy.Valid() {
		return fmt.Errorf("invalid default wait policy %q", p.DefaultWaitPolicy)
	}
	if !p.RecoveryBehavior.Valid() {
		return fmt.Errorf("invalid recovery behavior %q", p.RecoveryBehavior)
	}
	if p.ResourceScope == ResourceNone && p.RetryableWhenBusy {
		return fmt.Errorf("resource_scope none cannot be retryable when busy")
	}
	if p.ParallelSafe && p.RetryableWhenBusy {
		return fmt.Errorf("parallel-safe command cannot be retryable when busy")
	}
	return nil
}

const MissingPolicyCode = "coordination_policy_missing"

// MissingPolicyError is returned instead of applying a permissive implicit
// policy. This is the safe default for every unregistered command.
type MissingPolicyError struct {
	Selector string
	Value    string
}

func (e *MissingPolicyError) Error() string {
	return fmt.Sprintf("%s: no coordination policy for %s %q", MissingPolicyCode, e.Selector, e.Value)
}

func (e *MissingPolicyError) Code() string { return MissingPolicyCode }
