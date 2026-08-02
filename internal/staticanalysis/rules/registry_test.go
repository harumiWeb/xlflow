package rules

import (
	"reflect"
	"sort"
	"strings"
	"testing"
)

func TestRegistryContainsEveryProductionDiagnostic(t *testing.T) {
	want := strings.Fields(`
VB001 VB002 VB003 VB004 VB005 VB006 VB007 VB008 VB009 VB010 VB011 VB012 VB013 VB014 VB015
VB018 VB019 VB020 VB021 VB022 VB023 VB026 VB027 VB028 VB029 VB030 VB031 VB032 VB033 VB034
VB035 VB036 VB037 VB038 VB039 VB040 VB041 VB042 VB043 VB044
VBA101 VBA102 VBA103 VBA104 VBA105 VBA106 VBA201 VBA202 VBA203 VBA204 VBA205 VBA206 VBA207
VBA208 VBA209 VBA210 VBA211 VBA212 VBA213 VBA214 VBA215 VBA216 VBA217 VBA218 VBA219 VBA220`)
	gotRules := All()
	got := make([]string, len(gotRules))
	for i, rule := range gotRules {
		got[i] = rule.ID
	}
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("registry IDs differ\n got: %v\nwant: %v", got, want)
	}
	if _, ok := Lookup("VBA000"); ok {
		t.Fatal("synthetic VBA000 must not be registered")
	}
}

func TestRegistrySnapshotsAreSortedAndDefensive(t *testing.T) {
	all := All()
	if !sort.SliceIsSorted(all, func(i, j int) bool { return all[i].ID < all[j].ID }) {
		t.Fatal("All is not sorted by ID")
	}
	original := all[0]
	all[0].ID = "VB999"
	got, ok := Lookup(original.ID)
	if !ok || got != original {
		t.Fatalf("caller mutated registry through All: %+v, %v", got, ok)
	}

	catalog := CatalogSnapshot()
	if catalog.SchemaVersion != SchemaVersion || len(catalog.Items) != len(All()) {
		t.Fatalf("unexpected catalog snapshot: %+v", catalog)
	}
	catalog.Items[0].Title = "changed"
	if got, _ := Lookup(original.ID); got.Title == "changed" {
		t.Fatal("caller mutated registry through CatalogSnapshot")
	}
}

func TestLookupAndFamilyFiltering(t *testing.T) {
	rule, ok := Lookup("  vba211 ")
	if !ok || rule.ID != "VBA211" || rule.Family != FamilyAnalyze || !rule.PreflightBlocking || rule.InlineSuppressible {
		t.Fatalf("unexpected lookup result: %+v, %v", rule, ok)
	}
	mismatch, ok := Lookup("VBA216")
	if !ok || mismatch.DefaultSeverity != SeverityError || !mismatch.PreflightBlocking || mismatch.InlineSuppressible || !mismatch.Realtime {
		t.Fatalf("unexpected VBA216 metadata: %+v, %v", mismatch, ok)
	}
	unstable, ok := Lookup("VBA217")
	if !ok || unstable.DefaultSeverity != SeverityWarning || unstable.PreflightBlocking || !unstable.InlineSuppressible || !unstable.Realtime || unstable.Precision != PrecisionMedium {
		t.Fatalf("unexpected VBA217 metadata: %+v, %v", unstable, ok)
	}
	failureContract, ok := Lookup("VBA218")
	if !ok || failureContract.DefaultSeverity != SeverityWarning || failureContract.PreflightBlocking || !failureContract.InlineSuppressible || !failureContract.Realtime || failureContract.Scope != ScopeInterprocedural || failureContract.Precision != PrecisionHigh {
		t.Fatalf("unexpected VBA218 metadata: %+v, %v", failureContract, ok)
	}
	resourceLeaks, ok := Lookup("VBA219")
	if !ok || resourceLeaks.DefaultSeverity != SeverityWarning || resourceLeaks.PreflightBlocking || !resourceLeaks.InlineSuppressible || !resourceLeaks.Realtime || resourceLeaks.Scope != ScopeProcedureLocal || resourceLeaks.Precision != PrecisionHigh {
		t.Fatalf("unexpected VBA219 metadata: %+v, %v", resourceLeaks, ok)
	}
	eventReentry, ok := Lookup("VBA220")
	if !ok || eventReentry.DefaultSeverity != SeverityWarning || eventReentry.PreflightBlocking || !eventReentry.InlineSuppressible || eventReentry.Realtime || eventReentry.Scope != ScopeInterprocedural || eventReentry.Precision != PrecisionMedium {
		t.Fatalf("unexpected VBA220 metadata: %+v, %v", eventReentry, ok)
	}
	for _, rule := range ByFamily(FamilyLint) {
		if rule.Family != FamilyLint {
			t.Fatalf("lint family contains %+v", rule)
		}
	}
}

func TestValidateRejectsInvalidMetadata(t *testing.T) {
	base, ok := Lookup("VB001")
	if !ok {
		t.Fatal("VB001 missing")
	}
	tests := []struct {
		name string
		edit func(*RuleMetadata)
	}{
		{"invalid ID", func(r *RuleMetadata) { r.ID = "VB1" }},
		{"synthetic ID", func(r *RuleMetadata) { r.ID = "VBA000"; r.Family = FamilyAnalyze }},
		{"missing title", func(r *RuleMetadata) { r.Title = "" }},
		{"invalid family", func(r *RuleMetadata) { r.Family = "editor" }},
		{"family mismatch", func(r *RuleMetadata) { r.Family = FamilyAnalyze }},
		{"invalid category", func(r *RuleMetadata) { r.Category = "style" }},
		{"invalid severity", func(r *RuleMetadata) { r.DefaultSeverity = "info" }},
		{"invalid scope", func(r *RuleMetadata) { r.Scope = "module" }},
		{"invalid precision", func(r *RuleMetadata) { r.Precision = "certain" }},
		{"config key mismatch", func(r *RuleMetadata) { r.ConfigurationKey = "" }},
		{"invalid config key", func(r *RuleMetadata) { r.ConfigurationKey = "Option Explicit" }},
		{"unconfigurable disabled default", func(r *RuleMetadata) { r.Configurable = false; r.ConfigurationKey = ""; r.DefaultEnabled = false }},
		{"blocking suppressible", func(r *RuleMetadata) { r.PreflightBlocking = true }},
		{"blocking warning", func(r *RuleMetadata) {
			r.PreflightBlocking = true
			r.InlineSuppressible = false
			r.DefaultSeverity = SeverityWarning
		}},
		{"invalid documentation URL", func(r *RuleMetadata) { r.DocumentationURL = "https://example.com/vb001" }},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			rule := base
			tc.edit(&rule)
			if err := Validate([]RuleMetadata{rule}); err == nil {
				t.Fatalf("Validate accepted %+v", rule)
			}
		})
	}

	if err := Validate([]RuleMetadata{base, base}); err == nil || !strings.Contains(err.Error(), "duplicate") {
		t.Fatalf("Validate duplicate error = %v", err)
	}
}
