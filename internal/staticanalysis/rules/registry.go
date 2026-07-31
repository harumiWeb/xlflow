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
	SeverityError   RuleSeverity = "error"
	SeverityWarning RuleSeverity = "warning"
)

type RuleMetadata struct {
	ID                 string        `json:"id"`
	Title              string        `json:"title"`
	Description        string        `json:"description"`
	Family             RuleFamily    `json:"family"`
	Category           RuleCategory  `json:"category"`
	DefaultSeverity    RuleSeverity  `json:"default_severity"`
	DefaultEnabled     bool          `json:"default_enabled"`
	Scope              RuleScope     `json:"scope"`
	Realtime           bool          `json:"realtime"`
	Precision          RulePrecision `json:"precision"`
	FixAvailable       bool          `json:"fix_available"`
	DocumentationURL   string        `json:"documentation_url"`
	Configurable       bool          `json:"configurable"`
	ConfigurationKey   string        `json:"configuration_key"`
	InlineSuppressible bool          `json:"inline_suppressible"`
	PreflightBlocking  bool          `json:"preflight_blocking"`
}

const SchemaVersion = 1

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
	registry = append([]RuleMetadata(nil), catalog.Items...)
	sort.Slice(registry, func(i, j int) bool { return registry[i].ID < registry[j].ID })
	byID = make(map[string]RuleMetadata, len(registry))
	for _, rule := range registry {
		byID[rule.ID] = rule
	}
}

// All returns all rules in deterministic diagnostic-ID order.
func All() []RuleMetadata { return append([]RuleMetadata(nil), registry...) }

// Lookup performs a case-insensitive lookup after trimming whitespace.
func Lookup(id string) (RuleMetadata, bool) {
	rule, ok := byID[strings.ToUpper(strings.TrimSpace(id))]
	return rule, ok
}

// ByFamily returns a defensive, ID-sorted snapshot for family.
func ByFamily(family RuleFamily) []RuleMetadata {
	out := make([]RuleMetadata, 0)
	for _, rule := range registry {
		if rule.Family == family {
			out = append(out, rule)
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
		if !validFamily(rule.Family) || !validCategory(rule.Category) || !validSeverity(rule.DefaultSeverity) || !validScope(rule.Scope) || !validPrecision(rule.Precision) {
			return fmt.Errorf("%s has invalid enum metadata", rule.ID)
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
func validSeverity(v RuleSeverity) bool { return v == SeverityError || v == SeverityWarning }
func validScope(v RuleScope) bool {
	return v == ScopeFileLocal || v == ScopeProcedureLocal || v == ScopeProjectWide || v == ScopeInterprocedural
}
func validPrecision(v RulePrecision) bool {
	return v == PrecisionHigh || v == PrecisionMedium || v == PrecisionLow
}
