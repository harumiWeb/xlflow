// Package preflight owns the project-level policy that decides whether a
// registry-owned diagnostic blocks source preflight.
package preflight

import (
	"strings"

	staticrules "github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
)

type Decision uint8

const (
	DecisionNonBlocking Decision = iota
	DecisionBlocking
	DecisionAllowed
)

type Policy struct {
	allowed map[string]struct{}
}

func NewPolicy(allowedDiagnostics []string) Policy {
	allowed := make(map[string]struct{}, len(allowedDiagnostics))
	for _, raw := range allowedDiagnostics {
		id := strings.ToUpper(strings.TrimSpace(raw))
		if id != "" {
			allowed[id] = struct{}{}
		}
	}
	return Policy{allowed: allowed}
}

func (p Policy) Decide(code string) Decision {
	id := strings.ToUpper(strings.TrimSpace(code))
	rule, ok := staticrules.Lookup(id)
	if !ok || !rule.PreflightBlocking {
		return DecisionNonBlocking
	}
	if _, ok := p.allowed[id]; ok {
		return DecisionAllowed
	}
	return DecisionBlocking
}

func Partition[T any](items []T, code func(T) string, policy Policy) (blocking []T, allowed []T) {
	for _, item := range items {
		switch policy.Decide(code(item)) {
		case DecisionBlocking:
			blocking = append(blocking, item)
		case DecisionAllowed:
			allowed = append(allowed, item)
		}
	}
	return blocking, allowed
}
