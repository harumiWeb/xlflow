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
VB035 VB036 VB037 VB038 VB039 VB040 VB041 VB042 VB043 VB044 VB045
VBA101 VBA102 VBA103 VBA104 VBA105 VBA106 VBA201 VBA202 VBA203 VBA204 VBA205 VBA206 VBA207
VBA208 VBA209 VBA210 VBA211 VBA212 VBA213 VBA214 VBA215 VBA216 VBA217 VBA218 VBA219 VBA220 VBA221 VBA222 VBA223 VBA224 VBA225 VBA226 VBA227 VBA228 VBA229
VBA230 VBA231 VBA232 VBA233 VBA234 VBA235 VBA236 VBA237 VBA238 VBA239`)
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
	originalSeverity := original.SupportedSeverities[0]
	all[0].ID = "VB999"
	got, ok := Lookup(original.ID)
	if !ok || !reflect.DeepEqual(got, original) {
		t.Fatalf("caller mutated registry through All: %+v, %v", got, ok)
	}

	catalog := CatalogSnapshot()
	if catalog.SchemaVersion != SchemaVersion || len(catalog.Items) != len(All()) {
		t.Fatalf("unexpected catalog snapshot: %+v", catalog)
	}
	catalog.Items[0].Title = "changed"
	catalog.Items[0].Surfaces[0] = SurfaceLSP
	catalog.Items[0].SupportedSeverities[0] = SeverityError
	if got, _ := Lookup(original.ID); got.Title == "changed" {
		t.Fatal("caller mutated registry through CatalogSnapshot")
	}
	if got, _ := Lookup(original.ID); got.Surfaces[0] == SurfaceLSP || got.SupportedSeverities[0] != originalSeverity {
		t.Fatal("caller mutated registry metadata slices through CatalogSnapshot")
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
	byRef, ok := Lookup("VBA206")
	if !ok || byRef.EvidenceClass != EvidenceRuntimeSafety || byRef.CompileEquivalent || byRef.DefaultSeverity != SeverityWarning || byRef.PreflightBlocking || !byRef.InlineSuppressible || !byRef.Realtime || byRef.Scope != ScopeInterprocedural || byRef.Precision != PrecisionHigh || !byRef.DefaultEnabled || byRef.Category != "runtime-safety" {
		t.Fatalf("unexpected VBA206 metadata: %+v, %v", byRef, ok)
	}
	setScalar, ok := Lookup("VB037")
	if !ok || setScalar.EvidenceClass != EvidenceCompileEquivalent || !setScalar.CompileEquivalent || setScalar.DefaultSeverity != SeverityError || !setScalar.PreflightBlocking || setScalar.InlineSuppressible || !reflect.DeepEqual(setScalar.SupportedSeverities, []RuleSeverity{SeverityError}) {
		t.Fatalf("unexpected VB037 metadata: %+v, %v", setScalar, ok)
	}
	argumentBinding, ok := Lookup("VB045")
	if !ok || argumentBinding.EvidenceClass != EvidenceCompileEquivalent || !argumentBinding.CompileEquivalent || argumentBinding.DefaultSeverity != SeverityError || !argumentBinding.PreflightBlocking || argumentBinding.InlineSuppressible {
		t.Fatalf("unexpected VB045 metadata: %+v, %v", argumentBinding, ok)
	}
	byRefMismatch, ok := Lookup("VBA228")
	if !ok || byRefMismatch.EvidenceClass != EvidenceCompileEquivalent || !byRefMismatch.CompileEquivalent || byRefMismatch.DefaultSeverity != SeverityError || !byRefMismatch.PreflightBlocking || byRefMismatch.InlineSuppressible || byRefMismatch.Configurable {
		t.Fatalf("unexpected VBA228 metadata: %+v, %v", byRefMismatch, ok)
	}
	localType, ok := Lookup("VBA229")
	if !ok || localType.Family != FamilyAnalyze || localType.EvidenceClass != EvidenceCompileEquivalent || !localType.CompileEquivalent || localType.DefaultSeverity != SeverityError || !reflect.DeepEqual(localType.SupportedSeverities, []RuleSeverity{SeverityError}) || !reflect.DeepEqual(localType.Surfaces, []RuleSurface{SurfaceAnalyze, SurfaceLSP}) || !localType.DefaultEnabled || !localType.PreflightBlocking || localType.InlineSuppressible || localType.Configurable || !localType.Realtime || localType.Scope != ScopeProcedureLocal || localType.Precision != PrecisionHigh || localType.Category != CategoryTypeSafety {
		t.Fatalf("unexpected VBA229 metadata: %+v, %v", localType, ok)
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
	stateCallEffects, ok := Lookup("VBA221")
	if !ok || stateCallEffects.DefaultSeverity != SeverityWarning || stateCallEffects.PreflightBlocking || !stateCallEffects.InlineSuppressible || stateCallEffects.Realtime || stateCallEffects.Scope != ScopeInterprocedural || stateCallEffects.Precision != PrecisionHigh {
		t.Fatalf("unexpected VBA221 metadata: %+v, %v", stateCallEffects, ok)
	}
	publicAPITypeSafety, ok := Lookup("VBA222")
	if !ok || publicAPITypeSafety.DefaultSeverity != SeverityWarning || publicAPITypeSafety.PreflightBlocking || !publicAPITypeSafety.InlineSuppressible || publicAPITypeSafety.Realtime || publicAPITypeSafety.Scope != ScopeProjectWide || publicAPITypeSafety.Precision != PrecisionMedium || !publicAPITypeSafety.DefaultEnabled || publicAPITypeSafety.Category != CategoryTypeSafety {
		t.Fatalf("unexpected VBA222 metadata: %+v, %v", publicAPITypeSafety, ok)
	}
	hardcodedSecrets, ok := Lookup("VBA223")
	if !ok || hardcodedSecrets.DefaultSeverity != SeverityWarning || hardcodedSecrets.PreflightBlocking || !hardcodedSecrets.InlineSuppressible || !hardcodedSecrets.Realtime || hardcodedSecrets.Scope != ScopeFileLocal || hardcodedSecrets.Precision != PrecisionHigh || hardcodedSecrets.Category != CategorySecurity {
		t.Fatalf("unexpected VBA223 metadata: %+v, %v", hardcodedSecrets, ok)
	}
	dataFlow, ok := Lookup("VBA224")
	if !ok || dataFlow.Family != FamilyAnalyze || dataFlow.Category != CategorySecurity || dataFlow.DefaultSeverity != SeverityWarning || !dataFlow.DefaultEnabled || dataFlow.Scope != ScopeProcedureLocal || !dataFlow.Realtime || dataFlow.Precision != PrecisionMedium || !dataFlow.Configurable || dataFlow.ConfigurationKey != "detect_untrusted_data_flow" || !dataFlow.InlineSuppressible || dataFlow.PreflightBlocking {
		t.Fatalf("unexpected VBA224 metadata: %+v, %v", dataFlow, ok)
	}
	commandConstruction, ok := Lookup("VBA236")
	if !ok || commandConstruction.Family != FamilyAnalyze || commandConstruction.Category != CategorySecurity || commandConstruction.DefaultSeverity != SeverityWarning || !commandConstruction.DefaultEnabled || commandConstruction.Scope != ScopeProcedureLocal || !commandConstruction.Realtime || commandConstruction.Precision != PrecisionMedium || !commandConstruction.Configurable || commandConstruction.ConfigurationKey != "detect_unsafe_command_construction" || !commandConstruction.InlineSuppressible || commandConstruction.PreflightBlocking {
		t.Fatalf("unexpected VBA236 metadata: %+v, %v", commandConstruction, ok)
	}
	loopAccess, ok := Lookup("VBA225")
	if !ok || loopAccess.DefaultSeverity != SeverityWarning || !loopAccess.DefaultEnabled || !loopAccess.Configurable || loopAccess.ConfigurationKey != "detect_excel_cell_access_in_loops" || loopAccess.PreflightBlocking || !loopAccess.InlineSuppressible || !loopAccess.Realtime || loopAccess.Scope != ScopeInterprocedural || loopAccess.Precision != PrecisionMedium || loopAccess.Category != CategoryPerformance {
		t.Fatalf("unexpected VBA225 metadata: %+v, %v", loopAccess, ok)
	}
	valueShape, ok := Lookup("VBA226")
	if !ok || valueShape.DefaultSeverity != SeverityWarning || !valueShape.DefaultEnabled || !valueShape.Configurable || valueShape.ConfigurationKey != "detect_range_value_array_shape" || valueShape.PreflightBlocking || !valueShape.InlineSuppressible || !valueShape.Realtime || valueShape.Scope != ScopeProcedureLocal || valueShape.Precision != PrecisionMedium || valueShape.Category != CategoryRuntimeSafety {
		t.Fatalf("unexpected VBA226 metadata: %+v, %v", valueShape, ok)
	}
	arrayLifecycle, ok := Lookup("VBA227")
	if !ok || arrayLifecycle.Family != FamilyAnalyze || arrayLifecycle.Category != CategoryRuntimeSafety || arrayLifecycle.DefaultSeverity != SeverityWarning || !reflect.DeepEqual(arrayLifecycle.SupportedSeverities, []RuleSeverity{SeverityWarning}) || !reflect.DeepEqual(arrayLifecycle.Surfaces, []RuleSurface{SurfaceAnalyze, SurfaceLSP}) || !arrayLifecycle.DefaultEnabled || arrayLifecycle.Scope != ScopeInterprocedural || !arrayLifecycle.Realtime || arrayLifecycle.Precision != PrecisionMedium || !arrayLifecycle.Configurable || arrayLifecycle.ConfigurationKey != "detect_array_lifecycle_safety" || !arrayLifecycle.InlineSuppressible || arrayLifecycle.PreflightBlocking {
		t.Fatalf("unexpected VBA227 metadata: %+v, %v", arrayLifecycle, ok)
	}
	guard, ok := Lookup("VBA207")
	if !ok || !reflect.DeepEqual(guard.SupportedSeverities, []RuleSeverity{SeverityWarning, SeverityInformation}) {
		t.Fatalf("unexpected VBA207 severities: %+v, %v", guard, ok)
	}
	for _, id := range []string{"VBA230", "VBA231", "VBA232", "VBA233", "VBA234", "VBA235"} {
		rule, found := Lookup(id)
		if !found || rule.DefaultSeverity != SeverityWarning || !rule.DefaultEnabled || !rule.Configurable || rule.PreflightBlocking || !rule.InlineSuppressible || !rule.Realtime || rule.Precision != PrecisionHigh {
			t.Errorf("unexpected %s metadata: %+v, %v", id, rule, found)
		}
	}
	errorSuppression, ok := Lookup("VBA237")
	if !ok || errorSuppression.Family != FamilyAnalyze || errorSuppression.Category != CategoryRuntimeSafety || errorSuppression.DefaultSeverity != SeverityWarning || !errorSuppression.DefaultEnabled || !errorSuppression.Configurable || errorSuppression.ConfigurationKey != "detect_error_suppression_propagation" || errorSuppression.PreflightBlocking || !errorSuppression.InlineSuppressible || !errorSuppression.Realtime || errorSuppression.Scope != ScopeInterprocedural || errorSuppression.Precision != PrecisionHigh {
		t.Fatalf("unexpected VBA237 metadata: %+v, %v", errorSuppression, ok)
	}
	loopInvariant, ok := Lookup("VBA238")
	if !ok || loopInvariant.Family != FamilyAnalyze || loopInvariant.Category != CategoryPerformance || loopInvariant.EvidenceClass != EvidenceMaintainability || loopInvariant.DefaultSeverity != SeverityWarning || !loopInvariant.DefaultEnabled || !loopInvariant.Configurable || loopInvariant.ConfigurationKey != "detect_loop_invariant_excel_object_resolution" || loopInvariant.PreflightBlocking || !loopInvariant.InlineSuppressible || !loopInvariant.Realtime || loopInvariant.Scope != ScopeProcedureLocal || loopInvariant.Precision != PrecisionMedium {
		t.Fatalf("unexpected VBA238 metadata: %+v, %v", loopInvariant, ok)
	}
	unsafeSQL, ok := Lookup("VBA239")
	if !ok || unsafeSQL.Family != FamilyAnalyze || unsafeSQL.Category != CategorySecurity || unsafeSQL.DefaultSeverity != SeverityWarning || !unsafeSQL.DefaultEnabled || !unsafeSQL.Configurable || unsafeSQL.ConfigurationKey != "detect_unsafe_sql_construction" || unsafeSQL.PreflightBlocking || !unsafeSQL.InlineSuppressible || !unsafeSQL.Realtime || unsafeSQL.Scope != ScopeProcedureLocal || unsafeSQL.Precision != PrecisionMedium {
		t.Fatalf("unexpected VBA239 metadata: %+v, %v", unsafeSQL, ok)
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
		{"invalid evidence class", func(r *RuleMetadata) { r.EvidenceClass = "vbe" }},
		{"inconsistent compile equivalent flag", func(r *RuleMetadata) { r.CompileEquivalent = true }},
		{"invalid severity", func(r *RuleMetadata) { r.DefaultSeverity = "info" }},
		{"invalid scope", func(r *RuleMetadata) { r.Scope = "module" }},
		{"invalid precision", func(r *RuleMetadata) { r.Precision = "certain" }},
		{"invalid surface", func(r *RuleMetadata) { r.Surfaces = []RuleSurface{"editor"} }},
		{"surface realtime mismatch", func(r *RuleMetadata) { r.Realtime = false }},
		{"duplicate surface", func(r *RuleMetadata) { r.Surfaces = []RuleSurface{SurfaceLint, SurfaceLint, SurfaceLSP} }},
		{"empty supported severities", func(r *RuleMetadata) { r.SupportedSeverities = nil }},
		{"invalid supported severity", func(r *RuleMetadata) { r.SupportedSeverities = []RuleSeverity{"info"} }},
		{"duplicate supported severity", func(r *RuleMetadata) { r.SupportedSeverities = []RuleSeverity{SeverityError, SeverityError} }},
		{"default severity not first", func(r *RuleMetadata) { r.SupportedSeverities = []RuleSeverity{SeverityWarning, SeverityError} }},
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

	compileEquivalent, ok := Lookup("VB008")
	if !ok {
		t.Fatal("VB008 missing")
	}
	for _, tc := range []struct {
		name string
		edit func(*RuleMetadata)
	}{
		{"compile-equivalent warning", func(r *RuleMetadata) {
			r.DefaultSeverity = SeverityWarning
			r.SupportedSeverities = []RuleSeverity{SeverityWarning}
		}},
		{"compile-equivalent suppressible", func(r *RuleMetadata) { r.InlineSuppressible = true }},
		{"compile-equivalent nonblocking", func(r *RuleMetadata) { r.PreflightBlocking = false }},
	} {
		t.Run(tc.name, func(t *testing.T) {
			rule := compileEquivalent
			tc.edit(&rule)
			if err := Validate([]RuleMetadata{rule}); err == nil {
				t.Fatalf("Validate accepted %+v", rule)
			}
		})
	}
	for _, tc := range []struct {
		name string
		edit func(*RuleMetadata)
	}{
		{"inference error", func(r *RuleMetadata) {
			r.DefaultSeverity = SeverityError
			r.SupportedSeverities = []RuleSeverity{SeverityError}
		}},
		{"inference blocking", func(r *RuleMetadata) { r.PreflightBlocking = true }},
		{"inference supports error", func(r *RuleMetadata) { r.SupportedSeverities = []RuleSeverity{SeverityWarning, SeverityError} }},
	} {
		t.Run(tc.name, func(t *testing.T) {
			rule := base
			tc.edit(&rule)
			if err := Validate([]RuleMetadata{rule}); err == nil {
				t.Fatalf("Validate accepted %+v", rule)
			}
		})
	}
}

func TestRegistrySurfaceAndSeverityMetadata(t *testing.T) {
	for _, rule := range All() {
		wantSurface := []RuleSurface{RuleSurface(rule.Family)}
		if rule.Realtime {
			wantSurface = append(wantSurface, SurfaceLSP)
		}
		if !reflect.DeepEqual(rule.Surfaces, wantSurface) {
			t.Errorf("%s surfaces = %v, want %v", rule.ID, rule.Surfaces, wantSurface)
		}
		if len(rule.SupportedSeverities) == 0 || rule.SupportedSeverities[0] != rule.DefaultSeverity {
			t.Errorf("%s supported severities = %v, want default %q first", rule.ID, rule.SupportedSeverities, rule.DefaultSeverity)
		}
	}
	for _, id := range []string{"VBA214", "VBA225", "VBA238"} {
		rule, ok := Lookup(id)
		if !ok || !reflect.DeepEqual(rule.SupportedSeverities, []RuleSeverity{SeverityWarning}) {
			t.Errorf("%s supported severities = %v, want warning", id, rule.SupportedSeverities)
		}
	}
}

func TestRegistryMetadataSlicesAreDeepCopied(t *testing.T) {
	all := All()
	originalSeverity := all[0].SupportedSeverities[0]
	originalID := all[0].ID
	all[0].Surfaces[0] = SurfaceLSP
	all[0].SupportedSeverities[0] = SeverityError
	lookup, ok := Lookup(originalID)
	if !ok || lookup.Surfaces[0] == SurfaceLSP || lookup.SupportedSeverities[0] != originalSeverity {
		t.Fatalf("Lookup shared mutable metadata slices: %+v", lookup)
	}
	byFamily := ByFamily(lookup.Family)
	byFamily[0].Surfaces[0] = SurfaceLSP
	byFamily[0].SupportedSeverities[0] = SeverityError
	lookup, _ = Lookup(byFamily[0].ID)
	if lookup.Surfaces[0] == SurfaceLSP || lookup.SupportedSeverities[0] != originalSeverity {
		t.Fatalf("ByFamily shared mutable metadata slices: %+v", lookup)
	}
}
