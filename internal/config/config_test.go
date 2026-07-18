package config

import (
	"errors"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestLoadDefaultsAndConfiguredValues(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
name = "sales"
entry = "Main.Run"

[excel]
path = "build/Sales.xlsm"
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Project.Name != "sales" || cfg.Project.Entry != "Main.Run" {
		t.Fatalf("project config = %+v", cfg.Project)
	}
	if cfg.Src.Modules != "src/modules" {
		t.Fatalf("default modules dir = %q", cfg.Src.Modules)
	}
	if !cfg.VBA.Folders || cfg.VBA.FolderAnnotation != "update" || !cfg.VBA.DefaultComponentFolders {
		t.Fatalf("unexpected vba defaults: %+v", cfg.VBA)
	}
	if cfg.VBA.LineNumbers.Enabled {
		t.Fatal("vba.line_numbers.enabled must default to false")
	}
	if cfg.UserForm.CodeSource != "sidecar" {
		t.Fatalf("unexpected userform defaults: %+v", cfg.UserForm)
	}
	if cfg.Backup.Retention.Enabled ||
		cfg.Backup.Retention.MaxCount != 20 ||
		cfg.Backup.Retention.MaxAgeDays != 30 ||
		cfg.Backup.Retention.MinKeep != 5 ||
		cfg.Backup.Retention.MaxTotalSizeMB != 2048 {
		t.Fatalf("unexpected backup retention defaults: %+v", cfg.Backup.Retention)
	}
	if !cfg.Fmt.OperatorSpacing {
		t.Fatalf("expected fmt.operator_spacing default to be enabled")
	}
	if !cfg.Fmt.DeclarationSpacing {
		t.Fatalf("expected fmt.declaration_spacing default to be enabled")
	}
	if !cfg.Fmt.KeywordCasing {
		t.Fatalf("expected fmt.keyword_casing default to be enabled")
	}
	if !cfg.Fmt.BuiltinCasing {
		t.Fatalf("expected fmt.builtin_casing default to be enabled")
	}
	if !cfg.Lint.RequireOptionExplicit {
		t.Fatalf("lint defaults were not applied")
	}
	if cfg.Excel.Bridge != "auto" {
		t.Fatalf("unexpected excel bridge default: %q", cfg.Excel.Bridge)
	}
	if !cfg.Lint.ForbidInteractiveInput {
		t.Fatalf("interactive input lint default was not applied")
	}
	if !cfg.Lint.DetectMultipleDeclaratorClarity ||
		!cfg.Lint.DetectUnusedLocalVariables ||
		!cfg.Lint.DetectConfusingCallSyntax || !cfg.Lint.DetectForEachControlType ||
		!cfg.Lint.DetectDangerousResume {
		t.Fatalf("expected high-signal AST lint defaults to be enabled: %+v", cfg.Lint)
	}
	if cfg.Lint.DetectScopeShadowing ||
		cfg.Lint.DetectUnusedPrivateProcedures ||
		cfg.Lint.DetectNestedWithAmbiguity {
		t.Fatalf("expected false-positive-prone AST lint defaults to be opt-in: %+v", cfg.Lint)
	}
	if cfg.Lint.ProcedureNameConstant.Enabled || cfg.Lint.ProcedureNameConstant.ConstantName != "" {
		t.Fatalf("procedure-name constant lint rule must default to disabled: %+v", cfg.Lint.ProcedureNameConstant)
	}
	if !cfg.Analyze.DetectRangeFindNothingCheck || !cfg.Analyze.DetectObjectUseBeforeSet ||
		!cfg.Analyze.DetectApplicationStateRestore || !cfg.Analyze.DetectErrorHandlerFallthrough ||
		!cfg.Analyze.ForbidUnqualifiedExcelObjects || !cfg.Analyze.DetectRedimPreserveDimension ||
		!cfg.Analyze.DetectObjectArrayComparison || !cfg.Analyze.DetectExcelObjectMemberMismatch ||
		!cfg.Analyze.DetectNonShortCircuitObjectGuard {
		t.Fatalf("expected high-signal analyze defaults to be enabled: %+v", cfg.Analyze)
	}
	if cfg.Analyze.DetectByRefArgumentMismatch || cfg.Analyze.DetectDictionaryCollectionGuard ||
		cfg.Analyze.DetectFunctionReturnPath {
		t.Fatalf("expected false-positive-prone analyze defaults to be opt-in: %+v", cfg.Analyze)
	}
}

func TestLoadParsesVBALineNumbers(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[vba.line_numbers]
enabled = true
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if !cfg.VBA.LineNumbers.Enabled {
		t.Fatal("vba.line_numbers.enabled was not loaded")
	}
}

func TestLoadParsesBackupRetention(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[backup.retention]
enabled = true
max_count = 10
max_age_days = 7
min_keep = 3
max_total_size_mb = 512
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	want := BackupRetentionConfig{Enabled: true, MaxCount: 10, MaxAgeDays: 7, MinKeep: 3, MaxTotalSizeMB: 512}
	if cfg.Backup.Retention != want {
		t.Fatalf("backup.retention = %+v, want %+v", cfg.Backup.Retention, want)
	}
}

func TestLoadBackupRetentionOmittedFieldsKeepDefaultsAndLimitsCanBeDisabled(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[backup.retention]
enabled = true
max_count = 0
max_age_days = 0
max_total_size_mb = 0
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if !cfg.Backup.Retention.Enabled ||
		cfg.Backup.Retention.MaxCount != 0 ||
		cfg.Backup.Retention.MaxAgeDays != 0 ||
		cfg.Backup.Retention.MinKeep != 5 ||
		cfg.Backup.Retention.MaxTotalSizeMB != 0 {
		t.Fatalf("backup.retention = %+v", cfg.Backup.Retention)
	}
}

func TestLoadRejectsInvalidBackupRetention(t *testing.T) {
	tests := []struct {
		name string
		body string
		want string
	}{
		{"negative max_count", "max_count = -1", "backup.retention.max_count"},
		{"negative max_age_days", "max_age_days = -1", "backup.retention.max_age_days"},
		{"negative min_keep", "min_keep = -1", "backup.retention.min_keep"},
		{"negative max_total_size_mb", "max_total_size_mb = -1", "backup.retention.max_total_size_mb"},
		{"min_keep exceeds max_count", "max_count = 2\nmin_keep = 3", "backup.retention.min_keep"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			dir := t.TempDir()
			body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[backup.retention]
` + tt.body + "\n")
			if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
				t.Fatal(err)
			}
			_, err := Load(dir)
			if err == nil || !strings.Contains(err.Error(), tt.want) {
				t.Fatalf("Load error = %v, want key %q", err, tt.want)
			}
		})
	}
}

func TestLoadSupportsDisablingFmtOperatorSpacing(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[fmt]
operator_spacing = false
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Fmt.OperatorSpacing {
		t.Fatal("expected fmt.operator_spacing=false to be honored")
	}
}

func TestLoadSupportsDisablingFmtDeclarationSpacing(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[fmt]
declaration_spacing = false
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Fmt.DeclarationSpacing {
		t.Fatal("expected fmt.declaration_spacing=false to be honored")
	}
	if !cfg.Fmt.OperatorSpacing {
		t.Fatal("expected omitted fmt.operator_spacing to remain enabled")
	}
}

func TestLoadSupportsDisablingFmtCasing(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[fmt]
keyword_casing = false
builtin_casing = false
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Fmt.KeywordCasing {
		t.Fatal("expected fmt.keyword_casing=false to be honored")
	}
	if cfg.Fmt.BuiltinCasing {
		t.Fatal("expected fmt.builtin_casing=false to be honored")
	}
	if !cfg.Fmt.OperatorSpacing || !cfg.Fmt.DeclarationSpacing {
		t.Fatal("expected omitted fmt spacing settings to remain enabled")
	}
}

func TestLoadAllowsDisablingInteractiveInputLint(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[lint]
forbid_interactive_input = false

[analyze]
detect_byref_argument_mismatch = true
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Lint.ForbidInteractiveInput {
		t.Fatal("expected forbid_interactive_input=false to be honored")
	}
	if !cfg.Analyze.DetectByRefArgumentMismatch {
		t.Fatal("expected analyze override to be honored")
	}
	if !cfg.Lint.RequireOptionExplicit {
		t.Fatal("expected other lint defaults to remain enabled")
	}
}

func TestLoadSupportsDisabledLintRules(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[lint]
disabled_rules = ["VB002", "vb006", "VB002"]
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Lint.ForbidSelect {
		t.Fatal("expected VB002/forbid_select to be disabled")
	}
	if cfg.Lint.ForbidPublicModuleFields {
		t.Fatal("expected VB006/forbid_public_module_fields to be disabled")
	}
	if !cfg.Lint.ForbidActivate {
		t.Fatal("expected unrelated lint rule to remain enabled")
	}
	if got := strings.Join(cfg.Lint.DisabledRules, ","); got != "VB002,VB006" {
		t.Fatalf("disabled rules = %q, want VB002,VB006", got)
	}
}

func TestLoadSupportsDisabledAnalyzeRules(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[analyze]
disabled_rules = ["VBA201", "vba205", "VBA212", "VBA201"]
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Analyze.DetectRangeFindNothingCheck {
		t.Fatal("expected VBA201/detect_range_find_nothing_check to be disabled")
	}
	if cfg.Analyze.ForbidUnqualifiedExcelObjects {
		t.Fatal("expected VBA205/forbid_unqualified_excel_objects to be disabled")
	}
	if cfg.Analyze.DetectNonShortCircuitObjectGuard {
		t.Fatal("expected VBA212/detect_non_short_circuit_object_guard to be disabled")
	}
	if !cfg.Analyze.DetectObjectUseBeforeSet {
		t.Fatal("expected unrelated analyze rule to remain enabled")
	}
	if got := strings.Join(cfg.Analyze.DisabledRules, ","); got != "VBA201,VBA205,VBA212" {
		t.Fatalf("disabled analyze rules = %q, want VBA201,VBA205,VBA212", got)
	}
}

func TestLoadRejectsUnknownDisabledLintRule(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[lint]
disabled_rules = ["VB999"]
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	_, err := Load(dir)
	if err == nil || !strings.Contains(err.Error(), "unknown lint rule ID in [lint].disabled_rules: VB999") {
		t.Fatalf("expected unknown lint rule error, got %v", err)
	}
}

func TestLoadRejectsUnknownDisabledAnalyzeRule(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[analyze]
disabled_rules = ["VBA999"]
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	_, err := Load(dir)
	if err == nil || !strings.Contains(err.Error(), "unknown analyze rule ID in [analyze].disabled_rules: VBA999") {
		t.Fatalf("expected unknown analyze rule error, got %v", err)
	}
}

func TestLoadRejectsNonConfigurableDisabledLintRule(t *testing.T) {
	for _, ruleID := range []string{"VB013", "VB015", "VB028", "VB029", "VB031", "VB032"} {
		t.Run(ruleID, func(t *testing.T) {
			dir := t.TempDir()
			body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[lint]
disabled_rules = ["` + ruleID + `"]
`)
			if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
				t.Fatal(err)
			}
			_, err := Load(dir)
			want := "lint rule ID is not configurable in [lint].disabled_rules: " + ruleID
			if err == nil || !strings.Contains(err.Error(), want) {
				t.Fatalf("expected non-configurable lint rule error %q, got %v", want, err)
			}
		})
	}
}

func TestLoadProcedureNameConstantConfig(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[lint]
disabled_rules = ["VB044"]

[lint.procedure_name_constant]
enabled = true
constant_name = " procedure_name "
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Lint.ProcedureNameConstant.Enabled || cfg.Lint.ProcedureNameConstant.ConstantName != "procedure_name" {
		t.Fatalf("disabled_rules should override the enabled procedure-name constant rule: %+v", cfg.Lint.ProcedureNameConstant)
	}
	if !hasConfigWarning(cfg.Warnings, "conflicting_lint_rule_config", "VB044") || !hasConfigWarning(cfg.Warnings, "disabled_rules_precedence", "VB044") {
		t.Fatalf("expected VB044 conflict warnings, got %+v", cfg.Warnings)
	}
}

func TestLoadRejectsInvalidProcedureNameConstantConfig(t *testing.T) {
	for name, constantName := range map[string]string{
		"missing":            "",
		"invalid_character":  "PROCEDURE NAME",
		"leading_underscore": "_PROCEDURE_NAME",
		"combining_mark":     "PROCEDURE\u0301_NAME",
		"too_long":           strings.Repeat("A", 256),
	} {
		t.Run(name, func(t *testing.T) {
			dir := t.TempDir()
			body := []byte(`[project]
entry = "Main.Run"

[lint.procedure_name_constant]
enabled = true
constant_name = "` + constantName + `"
`)
			if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
				t.Fatal(err)
			}
			if _, err := Load(dir); err == nil || !strings.Contains(err.Error(), "lint.procedure_name_constant.constant_name") {
				t.Fatalf("expected constant-name validation error, got %v", err)
			}
		})
	}
}

func TestValidVBAIdentifier(t *testing.T) {
	for name, valid := range map[string]bool{
		"PROCEDURE_NAME":         true,
		"手続き名":                   true,
		strings.Repeat("A", 255): true,
		"_PROCEDURE_NAME":        false,
		"PROCEDURE\u0301_NAME":   false,
		strings.Repeat("A", 256): false,
	} {
		if got := validVBAIdentifier(name); got != valid {
			t.Errorf("validVBAIdentifier(%q) = %t, want %t", name, got, valid)
		}
	}
}

func TestLoadRejectsNonConfigurableDisabledAnalyzeRule(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[analyze]
disabled_rules = ["VBA104"]
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	_, err := Load(dir)
	if err == nil || !strings.Contains(err.Error(), "analyze rule ID is not configurable in [analyze].disabled_rules: VBA104") {
		t.Fatalf("expected non-configurable analyze rule error, got %v", err)
	}
}

func TestLoadWarnsForLegacyLintRuleConfig(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[lint]
forbid_select = false
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Lint.ForbidSelect {
		t.Fatal("expected legacy forbid_select=false to be honored")
	}
	if !hasConfigWarning(cfg.Warnings, "deprecated_lint_rule_config", "VB002") {
		t.Fatalf("expected deprecated config warning, got %+v", cfg.Warnings)
	}
	if !hasConfigWarningMessage(cfg.Warnings, "deprecated_lint_rule_config", "VB002", `Use [lint].disabled_rules = ["VB002"] instead`) {
		t.Fatalf("expected false legacy lint warning to suggest disabled_rules, got %+v", cfg.Warnings)
	}
}

func TestLoadWarnsForLegacyAnalyzeRuleConfig(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[analyze]
detect_byref_argument_mismatch = true
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if !cfg.Analyze.DetectByRefArgumentMismatch {
		t.Fatal("expected legacy detect_byref_argument_mismatch=true to be honored")
	}
	if !hasConfigWarning(cfg.Warnings, "deprecated_analyze_rule_config", "VBA206") {
		t.Fatalf("expected deprecated analyze config warning, got %+v", cfg.Warnings)
	}
	if !hasConfigWarningMessage(cfg.Warnings, "deprecated_analyze_rule_config", "VBA206", "compatibility opt-in") {
		t.Fatalf("expected true opt-in analyze warning to avoid disabled_rules migration, got %+v", cfg.Warnings)
	}
	if hasConfigWarningMessage(cfg.Warnings, "deprecated_analyze_rule_config", "VBA206", "disabled_rules") {
		t.Fatalf("true opt-in analyze warning must not suggest disabled_rules, got %+v", cfg.Warnings)
	}
}

func TestLoadWarnsForRedundantLegacyTrueRuleConfig(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[lint]
forbid_select = true
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if !cfg.Lint.ForbidSelect {
		t.Fatal("expected legacy forbid_select=true to keep rule enabled")
	}
	if !hasConfigWarningMessage(cfg.Warnings, "deprecated_lint_rule_config", "VB002", "Remove [lint].forbid_select") {
		t.Fatalf("expected true default-on lint warning to suggest removal, got %+v", cfg.Warnings)
	}
	if hasConfigWarningMessage(cfg.Warnings, "deprecated_lint_rule_config", "VB002", "disabled_rules") {
		t.Fatalf("true default-on lint warning must not suggest disabled_rules, got %+v", cfg.Warnings)
	}
}

func TestLoadDisabledRulesTakePrecedenceOverLegacyLintRuleConfig(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[lint]
forbid_public_module_fields = true
disabled_rules = ["VB006"]
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Lint.ForbidPublicModuleFields {
		t.Fatal("expected disabled_rules to take precedence over legacy true")
	}
	if !hasConfigWarning(cfg.Warnings, "deprecated_lint_rule_config", "VB006") ||
		!hasConfigWarning(cfg.Warnings, "conflicting_lint_rule_config", "VB006") ||
		!hasConfigWarning(cfg.Warnings, "disabled_rules_precedence", "VB006") {
		t.Fatalf("expected deprecation and conflict warnings, got %+v", cfg.Warnings)
	}
}

func TestLoadAnalyzeDisabledRulesTakePrecedenceOverLegacyRuleConfig(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[analyze]
forbid_unqualified_excel_objects = true
disabled_rules = ["VBA205"]
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Analyze.ForbidUnqualifiedExcelObjects {
		t.Fatal("expected analyze disabled_rules to take precedence over legacy true")
	}
	if !hasConfigWarning(cfg.Warnings, "deprecated_analyze_rule_config", "VBA205") ||
		!hasConfigWarning(cfg.Warnings, "conflicting_analyze_rule_config", "VBA205") ||
		!hasConfigWarning(cfg.Warnings, "analyze_disabled_rules_precedence", "VBA205") {
		t.Fatalf("expected analyze deprecation and conflict warnings, got %+v", cfg.Warnings)
	}
}

func TestLoadMissingConfig(t *testing.T) {
	_, err := Load(t.TempDir())
	if err == nil {
		t.Fatal("expected missing config error")
	}
	if !errors.Is(err, ErrConfigNotFound) {
		t.Fatalf("expected ErrConfigNotFound, got %v", err)
	}
}

func TestLoadRejectsInvalidFolderAnnotation(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[vba]
folder_annotation = "broken"
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	_, err := Load(dir)
	if err == nil {
		t.Fatal("expected invalid folder annotation error")
	}
}

func TestLoadRejectsInvalidUserFormCodeSource(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"

[userform]
code_source = "broken"
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	_, err := Load(dir)
	if err == nil {
		t.Fatal("expected invalid userform code source error")
	}
}

func TestLoadRejectsInvalidExcelBridge(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"
bridge = "broken"
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	_, err := Load(dir)
	if err == nil {
		t.Fatal("expected invalid excel bridge error")
	}
}

func TestLoadNormalizesExcelBridgeValue(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"
bridge = " DotNet "
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	cfg, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if cfg.Excel.Bridge != "dotnet" {
		t.Fatalf("excel.bridge = %q, want dotnet", cfg.Excel.Bridge)
	}
}

func TestLoadRejectsRemovedPowerShellExcelBridge(t *testing.T) {
	dir := t.TempDir()
	body := []byte(`[project]
entry = "Main.Run"

[excel]
path = "build/Book.xlsm"
bridge = " PowerShell "
`)
	if err := os.WriteFile(filepath.Join(dir, FileName), body, 0o644); err != nil {
		t.Fatal(err)
	}
	_, err := Load(dir)
	if !errors.Is(err, ErrInvalidExcelBridge) {
		t.Fatalf("Load error = %v, want %v", err, ErrInvalidExcelBridge)
	}
}

func TestWriteProducesReadableConfig(t *testing.T) {
	dir := t.TempDir()
	cfg := Default()
	cfg.Project.Name = "write-test"
	cfg.Excel.Bridge = "dotnet"
	cfg.UserForm.CodeSource = "frm"
	cfg.Lint.ForbidInteractiveInput = false
	cfg.Analyze.ForbidUnqualifiedExcelObjects = false

	p := filepath.Join(dir, FileName)
	if err := Write(p, cfg); err != nil {
		t.Fatalf("Write failed: %v", err)
	}
	body, err := os.ReadFile(p)
	if err != nil {
		t.Fatal(err)
	}
	text := string(body)
	if !strings.Contains(text, "disabled_rules = [") || !strings.Contains(text, `"VB007"`) {
		t.Fatalf("expected generated config to disable VB007 by ID:\n%s", text)
	}
	if !strings.Contains(text, `"VBA205"`) {
		t.Fatalf("expected generated config to disable VBA205 by ID:\n%s", text)
	}
	if strings.Contains(text, "forbid_interactive_input = false") || strings.Contains(text, "require_option_explicit = true") {
		t.Fatalf("generated config should prefer disabled_rules over legacy lint booleans:\n%s", text)
	}
	if strings.Contains(text, "forbid_unqualified_excel_objects = false") || strings.Contains(text, "detect_range_find_nothing_check = true") {
		t.Fatalf("generated config should prefer disabled_rules over legacy analyze booleans:\n%s", text)
	}
	if !strings.Contains(text, "[fmt]") ||
		!strings.Contains(text, "operator_spacing = true") ||
		!strings.Contains(text, "declaration_spacing = true") ||
		!strings.Contains(text, "keyword_casing = true") ||
		!strings.Contains(text, "builtin_casing = true") {
		t.Fatalf("generated config should include fmt spacing settings:\n%s", text)
	}
	for _, want := range []string{
		"# [backup.retention]",
		"# enabled = false",
		"# max_count = 20",
		"# min_keep = 5",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("generated config missing %q:\n%s", want, text)
		}
	}
	for _, want := range []string{
		"# [vba.line_numbers]",
		"# enabled = true",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("generated config missing %q:\n%s", want, text)
		}
	}
	for _, want := range []string{
		"# VB020 unused-local-variable warnings are enabled by default.",
		"# Add \"VB020\" to disabled_rules if a project intentionally keeps scratch locals.",
		"# detect_unused_private_procedures = true # VB021",
		"# [lint.procedure_name_constant]",
		"# constant_name = \"PROCEDURE_NAME\"",
	} {
		if !strings.Contains(text, want) {
			t.Fatalf("generated config missing %q:\n%s", want, text)
		}
	}

	loaded, err := Load(dir)
	if err != nil {
		t.Fatalf("Load failed after Write: %v", err)
	}
	if loaded.Project.Name != "write-test" {
		t.Fatalf("name mismatch: got %q, want write-test", loaded.Project.Name)
	}
	if loaded.UserForm.CodeSource != "frm" {
		t.Fatalf("userform.code_source mismatch: got %q, want frm", loaded.UserForm.CodeSource)
	}
	if loaded.Excel.Bridge != "dotnet" {
		t.Fatalf("excel.bridge mismatch: got %q, want dotnet", loaded.Excel.Bridge)
	}
	if !loaded.Fmt.OperatorSpacing {
		t.Fatal("expected fmt.operator_spacing=true after Write/Load")
	}
	if !loaded.Fmt.DeclarationSpacing {
		t.Fatal("expected fmt.declaration_spacing=true after Write/Load")
	}
	if !loaded.Fmt.KeywordCasing {
		t.Fatal("expected fmt.keyword_casing=true after Write/Load")
	}
	if !loaded.Fmt.BuiltinCasing {
		t.Fatal("expected fmt.builtin_casing=true after Write/Load")
	}
	if loaded.Lint.ForbidInteractiveInput {
		t.Fatal("expected forbid_interactive_input=false")
	}
	if !loaded.Lint.RequireOptionExplicit {
		t.Fatal("expected require_option_explicit=true")
	}
	if !loaded.Analyze.DetectRangeFindNothingCheck {
		t.Fatal("expected analyze defaults to be written and loaded")
	}
	if !loaded.Analyze.DetectNonShortCircuitObjectGuard {
		t.Fatal("expected detect_non_short_circuit_object_guard=true after Write/Load")
	}
	if loaded.Analyze.ForbidUnqualifiedExcelObjects {
		t.Fatal("expected forbid_unqualified_excel_objects=false")
	}
}

func TestWriteProcedureNameConstantConfig(t *testing.T) {
	dir := t.TempDir()
	cfg := Default()
	cfg.Lint.ProcedureNameConstant = ProcedureNameConstantConfig{Enabled: true, ConstantName: "PROCEDURE_NAME"}
	p := filepath.Join(dir, FileName)
	if err := Write(p, cfg); err != nil {
		t.Fatal(err)
	}
	body, err := os.ReadFile(p)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(body), "[lint.procedure_name_constant]\nenabled = true\nconstant_name = \"PROCEDURE_NAME\"") {
		t.Fatalf("generated config missing enabled procedure-name constant rule:\n%s", body)
	}
	loaded, err := Load(dir)
	if err != nil {
		t.Fatal(err)
	}
	if !loaded.Lint.ProcedureNameConstant.Enabled || loaded.Lint.ProcedureNameConstant.ConstantName != "PROCEDURE_NAME" {
		t.Fatalf("procedure-name constant config did not round-trip: %+v", loaded.Lint.ProcedureNameConstant)
	}
}

func TestWriteOmitsOptionalLintHintsForEnabledOptIns(t *testing.T) {
	dir := t.TempDir()
	cfg := Default()
	cfg.Lint.DetectUnusedPrivateProcedures = true

	p := filepath.Join(dir, FileName)
	if err := Write(p, cfg); err != nil {
		t.Fatalf("Write failed: %v", err)
	}
	body, err := os.ReadFile(p)
	if err != nil {
		t.Fatal(err)
	}
	text := string(body)
	if strings.Contains(text, "# detect_unused_private_procedures = true # VB021") {
		t.Fatalf("generated config should not include opt-in hint for enabled VB021:\n%s", text)
	}
	if !strings.Contains(text, "detect_unused_private_procedures = true") {
		t.Fatalf("generated config should include enabled VB021 setting:\n%s", text)
	}
}

func TestUpdateUserFormCodeSourcePreservesUnrelatedConfig(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, FileName)
	body := "# top\r\n[project]\r\nname = \"demo\"\r\n\r\n[userform]\r\n# keep\r\ncode_source = \"frm\"\r\n\r\n[excel]\r\npath = \"build/Book.xlsm\"\r\n"
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := UpdateUserFormCodeSource(path, "sidecar"); err != nil {
		t.Fatal(err)
	}
	got, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	text := string(got)
	if !strings.Contains(text, "# top\r\n[project]") || !strings.Contains(text, "# keep\r\ncode_source = \"sidecar\"") || !strings.Contains(text, "\r\n[excel]\r\n") {
		t.Fatalf("updated config did not preserve unrelated text/newlines:\n%s", text)
	}
	if strings.Contains(text, "code_source = \"frm\"") {
		t.Fatalf("old code_source remained:\n%s", text)
	}
}

func TestUpdateUserFormCodeSourceAddsMissingSection(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, FileName)
	if err := os.WriteFile(path, []byte("[project]\nname = \"demo\"\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := UpdateUserFormCodeSource(path, "sidecar"); err != nil {
		t.Fatal(err)
	}
	got, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	text := string(got)
	if !strings.Contains(text, "\n[userform]\ncode_source = \"sidecar\"\n") {
		t.Fatalf("missing inserted userform section:\n%s", text)
	}
}

func TestUpdateUserFormCodeSourceHandlesTableCommentsAndQuotedKeys(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, FileName)
	body := "[project]\nname = \"demo\"\n\n[userform] # settings\n\"code_source\" = \"frm\"\n"
	if err := os.WriteFile(path, []byte(body), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := UpdateUserFormCodeSource(path, "sidecar"); err != nil {
		t.Fatal(err)
	}
	got, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	text := string(got)
	if strings.Count(text, "[userform]") != 1 {
		t.Fatalf("expected one userform table, got:\n%s", text)
	}
	if !strings.Contains(text, "[userform] # settings\ncode_source = \"sidecar\"") {
		t.Fatalf("expected quoted key to be replaced in existing table:\n%s", text)
	}
}

func hasConfigWarning(warnings []map[string]any, code string, rule string) bool {
	for _, warning := range warnings {
		if warning["code"] == code && warning["rule"] == rule {
			return true
		}
	}
	return false
}

func hasConfigWarningMessage(warnings []map[string]any, code string, rule string, text string) bool {
	for _, warning := range warnings {
		if warning["code"] != code || warning["rule"] != rule {
			continue
		}
		message, _ := warning["message"].(string)
		if strings.Contains(message, text) {
			return true
		}
	}
	return false
}
