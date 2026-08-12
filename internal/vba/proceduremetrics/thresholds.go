package proceduremetrics

import (
	"fmt"
	"sort"
)

// Thresholds maps metric names to their configured maximum values.  Missing
// keys and zero values disable that metric.  Positive values report a
// violation only when the measured value is strictly greater than the
// threshold; negative values are invalid configuration.
type Thresholds map[MetricName]int

// ThresholdViolation is the protocol-neutral form of an MX001 threshold
// finding.  The CLI can project it into its diagnostic envelope without
// making the collector depend on analyzer finding types.
type ThresholdViolation struct {
	File            string     `json:"file"`
	Line            int        `json:"line"`
	Module          string     `json:"module"`
	Procedure       string     `json:"procedure"`
	ProcedureKind   string     `json:"procedure_kind"`
	Metric          MetricName `json:"metric"`
	Value           int        `json:"value"`
	Threshold       int        `json:"threshold"`
	Severity        string     `json:"severity"`
	Code            string     `json:"code"`
	Message         string     `json:"message"`
	declarationByte int        // ordering-only; intentionally not serialized
}

// ValidateThresholds validates names and values before evaluation.  A nil
// map is valid and means all threshold diagnostics are disabled.
func ValidateThresholds(thresholds Thresholds) error {
	names := make([]string, 0, len(thresholds))
	for name := range thresholds {
		names = append(names, string(name))
	}
	sort.Strings(names)
	for _, rawName := range names {
		name := MetricName(rawName)
		threshold := thresholds[name]
		if _, ok := (Metrics{}).Value(name); !ok {
			return fmt.Errorf("unknown procedure metric threshold %q", name)
		}
		if threshold < 0 {
			return fmt.Errorf("procedure metric threshold %q must be non-negative", name)
		}
	}
	return nil
}

// EvaluateThresholds returns deterministic MX001 candidates for all
// procedures.  Metrics themselves are never modified.  The function rejects
// invalid threshold maps before inspecting procedure values.
func EvaluateThresholds(procedures []ProcedureMetrics, thresholds Thresholds) ([]ThresholdViolation, error) {
	if err := ValidateThresholds(thresholds); err != nil {
		return nil, err
	}
	ordered := append([]ProcedureMetrics(nil), procedures...)
	Sort(ordered)
	violations := make([]ThresholdViolation, 0)
	for _, procedure := range ordered {
		for _, metric := range canonicalMetricNames {
			threshold, enabled := thresholds[metric]
			if !enabled || threshold == 0 {
				continue
			}
			value, ok := procedure.Value(metric)
			if !ok || value <= threshold {
				continue
			}
			violations = append(violations, ThresholdViolation{
				File: procedure.File, Line: procedure.DeclarationRange.StartLine,
				Module: procedure.Module, Procedure: procedure.Name,
				ProcedureKind: string(procedure.Kind), Metric: metric,
				Value: value, Threshold: threshold, Severity: "warning", Code: "MX001",
				Message:         fmt.Sprintf("%s %s exceeds threshold %d (value %d)", procedure.Name, metric, threshold, value),
				declarationByte: procedure.DeclarationRange.StartByte,
			})
		}
	}
	// Keep a defensive sort in case callers provide procedures with duplicate
	// identities or declaration offsets.
	sort.SliceStable(violations, func(i, j int) bool {
		a, b := violations[i], violations[j]
		if a.File != b.File {
			return a.File < b.File
		}
		if a.Line != b.Line {
			return a.Line < b.Line
		}
		if a.declarationByte != b.declarationByte {
			return a.declarationByte < b.declarationByte
		}
		if a.Module != b.Module {
			return a.Module < b.Module
		}
		if a.Procedure != b.Procedure {
			return a.Procedure < b.Procedure
		}
		return metricIndex(a.Metric) < metricIndex(b.Metric)
	})
	return violations, nil
}

func metricIndex(name MetricName) int {
	for index, candidate := range canonicalMetricNames {
		if candidate == name {
			return index
		}
	}
	return len(canonicalMetricNames)
}
