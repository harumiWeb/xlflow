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
	if cfg.UserForm.CodeSource != "sidecar" {
		t.Fatalf("unexpected userform defaults: %+v", cfg.UserForm)
	}
	if !cfg.Fmt.OperatorSpacing {
		t.Fatalf("expected fmt.operator_spacing default to be enabled")
	}
	if !cfg.Fmt.DeclarationSpacing {
		t.Fatalf("expected fmt.declaration_spacing default to be enabled")
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
		!cfg.Lint.DetectConfusingCallSyntax || !cfg.Lint.DetectForEachControlType ||
		!cfg.Lint.DetectDangerousResume {
		t.Fatalf("expected high-signal AST lint defaults to be enabled: %+v", cfg.Lint)
	}
	if cfg.Lint.DetectScopeShadowing || cfg.Lint.DetectUnusedLocalVariables ||
		cfg.Lint.DetectUnusedPrivateProcedures ||
		cfg.Lint.DetectNestedWithAmbiguity {
		t.Fatalf("expected false-positive-prone AST lint defaults to be opt-in: %+v", cfg.Lint)
	}
	if !cfg.Analyze.DetectRangeFindNothingCheck || !cfg.Analyze.DetectObjectUseBeforeSet ||
		!cfg.Analyze.DetectApplicationStateRestore || !cfg.Analyze.DetectErrorHandlerFallthrough ||
		!cfg.Analyze.ForbidUnqualifiedExcelObjects || !cfg.Analyze.DetectRedimPreserveDimension ||
		!cfg.Analyze.DetectObjectArrayComparison || !cfg.Analyze.DetectExcelObjectMemberMismatch {
		t.Fatalf("expected high-signal analyze defaults to be enabled: %+v", cfg.Analyze)
	}
	if cfg.Analyze.DetectByRefArgumentMismatch || cfg.Analyze.DetectDictionaryCollectionGuard ||
		cfg.Analyze.DetectFunctionReturnPath {
		t.Fatalf("expected false-positive-prone analyze defaults to be opt-in: %+v", cfg.Analyze)
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
disabled_rules = ["VBA201", "vba205", "VBA201"]
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
	if !cfg.Analyze.DetectObjectUseBeforeSet {
		t.Fatal("expected unrelated analyze rule to remain enabled")
	}
	if got := strings.Join(cfg.Analyze.DisabledRules, ","); got != "VBA201,VBA205" {
		t.Fatalf("disabled analyze rules = %q, want VBA201,VBA205", got)
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
	for _, ruleID := range []string{"VB013", "VB028", "VB029", "VB031", "VB032"} {
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
		!strings.Contains(text, "declaration_spacing = true") {
		t.Fatalf("generated config should include fmt spacing settings:\n%s", text)
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
	if loaded.Lint.ForbidInteractiveInput {
		t.Fatal("expected forbid_interactive_input=false")
	}
	if !loaded.Lint.RequireOptionExplicit {
		t.Fatal("expected require_option_explicit=true")
	}
	if !loaded.Analyze.DetectRangeFindNothingCheck {
		t.Fatal("expected analyze defaults to be written and loaded")
	}
	if loaded.Analyze.ForbidUnqualifiedExcelObjects {
		t.Fatal("expected forbid_unqualified_excel_objects=false")
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
