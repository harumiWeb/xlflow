// Package rules owns protocol-neutral metadata for xlflow static-analysis rules.
package rules

import (
	_ "embed"
	"encoding/json"
	"errors"
	"fmt"
	"net/url"
	"regexp"
	"sort"
	"strings"
)

type RuleFamily string

const (
	FamilyLint    RuleFamily = "lint"
	FamilyAnalyze RuleFamily = "analyze"
)

type RuleCategory string

const (
	CategoryCorrectness     RuleCategory = "correctness"
	CategoryDocumentation   RuleCategory = "documentation"
	CategoryMaintainability RuleCategory = "maintainability"
	CategoryReliability     RuleCategory = "reliability"
	CategoryRuntimeSafety   RuleCategory = "runtime-safety"
	CategoryPerformance     RuleCategory = "performance"
	CategoryArchitecture    RuleCategory = "architecture"
	CategorySecurity        RuleCategory = "security"
	CategoryTypeSafety      RuleCategory = "type-safety"
)

// RuleEvidenceClass describes the strongest evidence behind a diagnostic.
// Only compile-equivalent evidence is strong enough to make a source
// preflight blocking claim. Runtime-error evidence is distinct from
// compile-equivalent evidence: it proves that execution will fail, but does
// not claim that the VBA source itself is rejected by the compiler.
type RuleEvidenceClass string

const (
	EvidenceCompileEquivalent RuleEvidenceClass = "compile-equivalent"
	EvidenceInference         RuleEvidenceClass = "inference"
	EvidenceRuntimeError      RuleEvidenceClass = "runtime-error"
	EvidenceRuntimeSafety     RuleEvidenceClass = "runtime-safety"
	EvidencePolicy            RuleEvidenceClass = "policy"
	EvidenceMaintainability   RuleEvidenceClass = "maintainability"
)

type RuleScope string

const (
	ScopeFileLocal       RuleScope = "file-local"
	ScopeProcedureLocal  RuleScope = "procedure-local"
	ScopeProjectWide     RuleScope = "project-wide"
	ScopeInterprocedural RuleScope = "interprocedural"
)

type RulePrecision string

const (
	PrecisionHigh   RulePrecision = "high"
	PrecisionMedium RulePrecision = "medium"
	PrecisionLow    RulePrecision = "low"
)

type RuleSeverity string

const (
	SeverityError       RuleSeverity = "error"
	SeverityWarning     RuleSeverity = "warning"
	SeverityInformation RuleSeverity = "information"
)

// RuleSurface identifies a public analysis surface that can emit a rule.
type RuleSurface string

const (
	SurfaceLint    RuleSurface = "lint"
	SurfaceAnalyze RuleSurface = "analyze"
	SurfaceLSP     RuleSurface = "lsp"
)

type RuleMetadata struct {
	ID                  string            `json:"id"`
	Title               string            `json:"title"`
	Description         string            `json:"description"`
	Family              RuleFamily        `json:"family"`
	Category            RuleCategory      `json:"category"`
	EvidenceClass       RuleEvidenceClass `json:"evidence_class"`
	CompileEquivalent   bool              `json:"compile_equivalent"`
	DefaultSeverity     RuleSeverity      `json:"default_severity"`
	Surfaces            []RuleSurface     `json:"surfaces"`
	SupportedSeverities []RuleSeverity    `json:"supported_severities"`
	DefaultEnabled      bool              `json:"default_enabled"`
	Scope               RuleScope         `json:"scope"`
	Realtime            bool              `json:"realtime"`
	Precision           RulePrecision     `json:"precision"`
	FixAvailable        bool              `json:"fix_available"`
	DocumentationURL    string            `json:"documentation_url"`
	Configurable        bool              `json:"configurable"`
	ConfigurationKey    string            `json:"configuration_key"`
	InlineSuppressible  bool              `json:"inline_suppressible"`
	PreflightBlocking   bool              `json:"preflight_blocking"`
}

const SchemaVersion = 2

type Catalog struct {
	SchemaVersion int            `json:"schema_version"`
	Items         []RuleMetadata `json:"items"`
}

//go:embed registry.json
var registryJSON []byte

var (
	registry                []RuleMetadata
	byID                    map[string]RuleMetadata
	idPattern               = regexp.MustCompile(`^(VB|VBA)[0-9]{3}$`)
	configurationKeyPattern = regexp.MustCompile(`^[a-z][a-z0-9_]*$`)
)

func init() {
	var catalog Catalog
	if err := json.Unmarshal(registryJSON, &catalog); err != nil {
		panic(fmt.Sprintf("decode static-analysis rule registry: %v", err))
	}
	if catalog.SchemaVersion != SchemaVersion {
		panic(fmt.Sprintf("static-analysis rule registry schema version is %d, want %d", catalog.SchemaVersion, SchemaVersion))
	}
	if err := Validate(catalog.Items); err != nil {
		panic(fmt.Sprintf("validate static-analysis rule registry: %v", err))
	}
	registry = cloneRules(catalog.Items)
	sort.Slice(registry, func(i, j int) bool { return registry[i].ID < registry[j].ID })
	byID = make(map[string]RuleMetadata, len(registry))
	for _, rule := range registry {
		byID[rule.ID] = cloneRule(rule)
	}
}

// All returns all rules in deterministic diagnostic-ID order.
func All() []RuleMetadata { return cloneRules(registry) }

// Lookup performs a case-insensitive lookup after trimming whitespace.
func Lookup(id string) (RuleMetadata, bool) {
	rule, ok := byID[strings.ToUpper(strings.TrimSpace(id))]
	return cloneRule(rule), ok
}

// ByFamily returns a defensive, ID-sorted snapshot for family.
func ByFamily(family RuleFamily) []RuleMetadata {
	out := make([]RuleMetadata, 0)
	for _, rule := range registry {
		if rule.Family == family {
			out = append(out, cloneRule(rule))
		}
	}
	return out
}

// CatalogSnapshot returns the public schema-versioned catalog.
func CatalogSnapshot() Catalog {
	return Catalog{SchemaVersion: SchemaVersion, Items: All()}
}

// PublicCatalog is an alias retained for callers that prefer an explicit public-API name.
func PublicCatalog() Catalog { return CatalogSnapshot() }

// Validate checks registry invariants independently of the embedded catalog.
func Validate(items []RuleMetadata) error {
	if len(items) == 0 {
		return errors.New("registry contains no rules")
	}
	seen := make(map[string]bool, len(items))
	configurationKeys := make(map[string]string)
	for i, rule := range items {
		prefix := fmt.Sprintf("rule %d", i)
		if !idPattern.MatchString(rule.ID) || rule.ID == "VBA000" {
			return fmt.Errorf("%s has invalid diagnostic ID %q", prefix, rule.ID)
		}
		if seen[rule.ID] {
			return fmt.Errorf("duplicate diagnostic ID %s", rule.ID)
		}
		seen[rule.ID] = true
		if strings.TrimSpace(rule.Title) == "" || strings.TrimSpace(rule.Description) == "" {
			return fmt.Errorf("%s is missing title or description", rule.ID)
		}
		if !validFamily(rule.Family) || !validCategory(rule.Category) || !validEvidenceClass(rule.EvidenceClass) || !validSeverity(rule.DefaultSeverity) || !validScope(rule.Scope) || !validPrecision(rule.Precision) {
			return fmt.Errorf("%s has invalid enum metadata", rule.ID)
		}
		if rule.CompileEquivalent != (rule.EvidenceClass == EvidenceCompileEquivalent) {
			return fmt.Errorf("%s compile_equivalent and evidence_class are inconsistent", rule.ID)
		}
		if err := validateSurfaces(rule); err != nil {
			return err
		}
		if err := validateSupportedSeverities(rule); err != nil {
			return err
		}
		if (rule.Family == FamilyLint && (!strings.HasPrefix(rule.ID, "VB") || strings.HasPrefix(rule.ID, "VBA"))) ||
			(rule.Family == FamilyAnalyze && !strings.HasPrefix(rule.ID, "VBA")) {
			return fmt.Errorf("%s does not match family %q", rule.ID, rule.Family)
		}
		if rule.Configurable != (strings.TrimSpace(rule.ConfigurationKey) != "") {
			return fmt.Errorf("%s configurable and configuration_key are inconsistent", rule.ID)
		}
		if !rule.Configurable && !rule.DefaultEnabled {
			return fmt.Errorf("%s is disabled by default but has no configuration binding", rule.ID)
		}
		if rule.Configurable {
			if !configurationKeyPattern.MatchString(rule.ConfigurationKey) {
				return fmt.Errorf("%s has invalid configuration_key %q", rule.ID, rule.ConfigurationKey)
			}
			qualifiedKey := string(rule.Family) + "." + rule.ConfigurationKey
			if prior, duplicate := configurationKeys[qualifiedKey]; duplicate {
				return fmt.Errorf("%s and %s share configuration_key %q", prior, rule.ID, qualifiedKey)
			}
			configurationKeys[qualifiedKey] = rule.ID
		}
		if rule.PreflightBlocking && rule.InlineSuppressible {
			return fmt.Errorf("%s cannot be both preflight-blocking and inline-suppressible", rule.ID)
		}
		if rule.PreflightBlocking && rule.DefaultSeverity != SeverityError {
			return fmt.Errorf("%s is preflight-blocking but does not have error severity", rule.ID)
		}
		if rule.CompileEquivalent {
			if rule.DefaultSeverity != SeverityError || len(rule.SupportedSeverities) != 1 || rule.SupportedSeverities[0] != SeverityError {
				return fmt.Errorf("%s compile-equivalent rules must support error severity only", rule.ID)
			}
			if rule.InlineSuppressible || !rule.PreflightBlocking {
				return fmt.Errorf("%s compile-equivalent rules must be unsuppressible and preflight-blocking", rule.ID)
			}
		} else if rule.EvidenceClass == EvidenceRuntimeError {
			if rule.DefaultSeverity != SeverityError || len(rule.SupportedSeverities) != 1 || rule.SupportedSeverities[0] != SeverityError {
				return fmt.Errorf("%s runtime-error rules must support error severity only", rule.ID)
			}
			if !rule.InlineSuppressible || rule.PreflightBlocking {
				return fmt.Errorf("%s runtime-error rules must be suppressible and non-preflight-blocking", rule.ID)
			}
		} else {
			if rule.DefaultSeverity == SeverityError || rule.PreflightBlocking {
				return fmt.Errorf("%s non-compile-equivalent rules cannot be error or preflight-blocking", rule.ID)
			}
			for _, severity := range rule.SupportedSeverities {
				if severity == SeverityError {
					return fmt.Errorf("%s non-compile-equivalent rules cannot support error severity", rule.ID)
				}
			}
		}
		wantURL := "https://harumiweb.github.io/xlflow/reference/diagnostics#" + strings.ToLower(rule.ID)
		parsed, err := url.ParseRequestURI(rule.DocumentationURL)
		if err != nil || parsed.Scheme != "https" || rule.DocumentationURL != wantURL {
			return fmt.Errorf("%s has invalid documentation_url %q", rule.ID, rule.DocumentationURL)
		}
	}
	return nil
}

func validFamily(v RuleFamily) bool { return v == FamilyLint || v == FamilyAnalyze }
func validCategory(v RuleCategory) bool {
	return v == CategoryCorrectness || v == CategoryDocumentation || v == CategoryMaintainability ||
		v == CategoryReliability || v == CategoryRuntimeSafety || v == CategoryPerformance ||
		v == CategoryArchitecture || v == CategorySecurity || v == CategoryTypeSafety
}
func validEvidenceClass(v RuleEvidenceClass) bool {
	return v == EvidenceCompileEquivalent || v == EvidenceInference || v == EvidenceRuntimeError || v == EvidenceRuntimeSafety || v == EvidencePolicy || v == EvidenceMaintainability
}
func validSeverity(v RuleSeverity) bool {
	return v == SeverityError || v == SeverityWarning || v == SeverityInformation
}
func validSurface(v RuleSurface) bool {
	return v == SurfaceLint || v == SurfaceAnalyze || v == SurfaceLSP
}
func validScope(v RuleScope) bool {
	return v == ScopeFileLocal || v == ScopeProcedureLocal || v == ScopeProjectWide || v == ScopeInterprocedural
}
func validPrecision(v RulePrecision) bool {
	return v == PrecisionHigh || v == PrecisionMedium || v == PrecisionLow
}

func validateSurfaces(rule RuleMetadata) error {
	wantBatch := RuleSurface(rule.Family)
	if len(rule.Surfaces) == 0 || rule.Surfaces[0] != wantBatch {
		return fmt.Errorf("%s must declare %q as its first surface", rule.ID, wantBatch)
	}
	wantLen := 1
	if rule.Realtime {
		wantLen = 2
	}
	// Compile-equivalent lint diagnostics are also projected by batch
	// analysis so source preflight can report them alongside analyzer
	// findings. Keep the public lint/LSP order first while allowing that
	// additional internal batch surface.
	batchAnalysisProjection := rule.Family == FamilyLint && rule.CompileEquivalent && len(rule.Surfaces) == wantLen+1 && rule.Surfaces[wantLen] == SurfaceAnalyze
	if len(rule.Surfaces) != wantLen && !batchAnalysisProjection {
		return fmt.Errorf("%s surfaces do not match realtime=%t", rule.ID, rule.Realtime)
	}
	for i, surface := range rule.Surfaces {
		if !validSurface(surface) {
			return fmt.Errorf("%s has invalid surface %q", rule.ID, surface)
		}
		if i > 0 && surface == rule.Surfaces[0] {
			return fmt.Errorf("%s has duplicate surface %q", rule.ID, surface)
		}
	}
	if rule.Realtime && rule.Surfaces[1] != SurfaceLSP {
		return fmt.Errorf("%s must declare %q when realtime", rule.ID, SurfaceLSP)
	}
	return nil
}

func validateSupportedSeverities(rule RuleMetadata) error {
	if len(rule.SupportedSeverities) == 0 {
		return fmt.Errorf("%s has no supported severities", rule.ID)
	}
	if rule.SupportedSeverities[0] != rule.DefaultSeverity {
		return fmt.Errorf("%s must declare default severity %q first", rule.ID, rule.DefaultSeverity)
	}
	seen := make(map[RuleSeverity]bool, len(rule.SupportedSeverities))
	for _, severity := range rule.SupportedSeverities {
		if !validSeverity(severity) {
			return fmt.Errorf("%s has invalid supported severity %q", rule.ID, severity)
		}
		if seen[severity] {
			return fmt.Errorf("%s has duplicate supported severity %q", rule.ID, severity)
		}
		seen[severity] = true
	}
	return nil
}

func cloneRule(rule RuleMetadata) RuleMetadata {
	rule.Surfaces = append([]RuleSurface(nil), rule.Surfaces...)
	rule.SupportedSeverities = append([]RuleSeverity(nil), rule.SupportedSeverities...)
	return rule
}

func cloneRules(items []RuleMetadata) []RuleMetadata {
	cloned := make([]RuleMetadata, len(items))
	for i, rule := range items {
		cloned[i] = cloneRule(rule)
	}
	return cloned
}
