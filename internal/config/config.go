package config

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/BurntSushi/toml"
	excelbridge "github.com/harumiWeb/xlflow/internal/excel/bridge"
)

const FileName = "xlflow.toml"

var (
	ErrConfigNotFound     = errors.New("config not found")
	ErrInvalidExcelBridge = errors.New("excel.bridge must be one of auto, dotnet; powershell bridge was removed, use dotnet")
)

type Config struct {
	Project  ProjectConfig    `toml:"project"`
	Excel    ExcelConfig      `toml:"excel"`
	Src      SourceConfig     `toml:"src"`
	VBA      VBAConfig        `toml:"vba"`
	UserForm UserFormConfig   `toml:"userform"`
	Fmt      FmtConfig        `toml:"fmt"`
	Lint     LintConfig       `toml:"lint"`
	Analyze  AnalyzeConfig    `toml:"analyze"`
	Warnings []map[string]any `toml:"-"`
}

type ProjectConfig struct {
	Name  string `toml:"name"`
	Entry string `toml:"entry"`
}

type ExcelConfig struct {
	Path          string `toml:"path"`
	Visible       bool   `toml:"visible"`
	DisplayAlerts bool   `toml:"display_alerts"`
	Bridge        string `toml:"bridge"`
}

type SourceConfig struct {
	Modules  string `toml:"modules"`
	Classes  string `toml:"classes"`
	Forms    string `toml:"forms"`
	Workbook string `toml:"workbook"`
}

type VBAConfig struct {
	Folders                 bool   `toml:"folders"`
	FolderAnnotation        string `toml:"folder_annotation"`
	DefaultComponentFolders bool   `toml:"default_component_folders"`
}

type UserFormConfig struct {
	CodeSource string `toml:"code_source"`
}

type FmtConfig struct {
	OperatorSpacing    bool `toml:"operator_spacing"`
	DeclarationSpacing bool `toml:"declaration_spacing"`
}

type LintConfig struct {
	DisabledRules                   []string `toml:"disabled_rules"`
	RequireOptionExplicit           bool     `toml:"require_option_explicit"`
	ForbidSelect                    bool     `toml:"forbid_select"`
	ForbidActivate                  bool     `toml:"forbid_activate"`
	ForbidOnErrorResumeNext         bool     `toml:"forbid_on_error_resume_next"`
	DetectImplicitVariant           bool     `toml:"detect_implicit_variant"`
	ForbidPublicModuleFields        bool     `toml:"forbid_public_module_fields"`
	ForbidInteractiveInput          bool     `toml:"forbid_interactive_input"`
	DetectScopeShadowing            bool     `toml:"detect_scope_shadowing"`
	DetectMultipleDeclaratorClarity bool     `toml:"detect_multiple_declarator_clarity"`
	DetectUnusedLocalVariables      bool     `toml:"detect_unused_local_variables"`
	DetectUnusedPrivateProcedures   bool     `toml:"detect_unused_private_procedures"`
	DetectConfusingCallSyntax       bool     `toml:"detect_confusing_call_syntax"`
	DetectForEachControlType        bool     `toml:"detect_for_each_control_type"`
	DetectDangerousResume           bool     `toml:"detect_dangerous_resume"`
	DetectNestedWithAmbiguity       bool     `toml:"detect_nested_with_ambiguity"`
}

type AnalyzeConfig struct {
	DisabledRules                   []string `toml:"disabled_rules"`
	DetectRangeFindNothingCheck     bool     `toml:"detect_range_find_nothing_check"`
	DetectObjectUseBeforeSet        bool     `toml:"detect_object_use_before_set"`
	DetectApplicationStateRestore   bool     `toml:"detect_application_state_restore"`
	DetectErrorHandlerFallthrough   bool     `toml:"detect_error_handler_fallthrough"`
	ForbidUnqualifiedExcelObjects   bool     `toml:"forbid_unqualified_excel_objects"`
	DetectByRefArgumentMismatch     bool     `toml:"detect_byref_argument_mismatch"`
	DetectDictionaryCollectionGuard bool     `toml:"detect_dictionary_collection_guard"`
	DetectRedimPreserveDimension    bool     `toml:"detect_redim_preserve_dimension"`
	DetectObjectArrayComparison     bool     `toml:"detect_object_array_comparison"`
	DetectFunctionReturnPath        bool     `toml:"detect_function_return_path"`
	DetectExcelObjectMemberMismatch bool     `toml:"detect_excel_object_member_mismatch"`
}

type lintRuleConfig struct {
	ID      string
	Key     string
	Default bool
	Get     func(LintConfig) bool
	Set     func(*LintConfig, bool)
}

type analyzeRuleConfig struct {
	ID      string
	Key     string
	Default bool
	Get     func(AnalyzeConfig) bool
	Set     func(*AnalyzeConfig, bool)
}

var configurableLintRules = []lintRuleConfig{
	{ID: "VB001", Key: "require_option_explicit", Default: true, Get: func(c LintConfig) bool { return c.RequireOptionExplicit }, Set: func(c *LintConfig, v bool) { c.RequireOptionExplicit = v }},
	{ID: "VB002", Key: "forbid_select", Default: true, Get: func(c LintConfig) bool { return c.ForbidSelect }, Set: func(c *LintConfig, v bool) { c.ForbidSelect = v }},
	{ID: "VB003", Key: "forbid_activate", Default: true, Get: func(c LintConfig) bool { return c.ForbidActivate }, Set: func(c *LintConfig, v bool) { c.ForbidActivate = v }},
	{ID: "VB004", Key: "forbid_on_error_resume_next", Default: true, Get: func(c LintConfig) bool { return c.ForbidOnErrorResumeNext }, Set: func(c *LintConfig, v bool) { c.ForbidOnErrorResumeNext = v }},
	{ID: "VB005", Key: "detect_implicit_variant", Default: true, Get: func(c LintConfig) bool { return c.DetectImplicitVariant }, Set: func(c *LintConfig, v bool) { c.DetectImplicitVariant = v }},
	{ID: "VB006", Key: "forbid_public_module_fields", Default: true, Get: func(c LintConfig) bool { return c.ForbidPublicModuleFields }, Set: func(c *LintConfig, v bool) { c.ForbidPublicModuleFields = v }},
	{ID: "VB007", Key: "forbid_interactive_input", Default: true, Get: func(c LintConfig) bool { return c.ForbidInteractiveInput }, Set: func(c *LintConfig, v bool) { c.ForbidInteractiveInput = v }},
	{ID: "VB018", Key: "detect_scope_shadowing", Default: false, Get: func(c LintConfig) bool { return c.DetectScopeShadowing }, Set: func(c *LintConfig, v bool) { c.DetectScopeShadowing = v }},
	{ID: "VB019", Key: "detect_multiple_declarator_clarity", Default: true, Get: func(c LintConfig) bool { return c.DetectMultipleDeclaratorClarity }, Set: func(c *LintConfig, v bool) { c.DetectMultipleDeclaratorClarity = v }},
	{ID: "VB020", Key: "detect_unused_local_variables", Default: true, Get: func(c LintConfig) bool { return c.DetectUnusedLocalVariables }, Set: func(c *LintConfig, v bool) { c.DetectUnusedLocalVariables = v }},
	{ID: "VB021", Key: "detect_unused_private_procedures", Default: false, Get: func(c LintConfig) bool { return c.DetectUnusedPrivateProcedures }, Set: func(c *LintConfig, v bool) { c.DetectUnusedPrivateProcedures = v }},
	{ID: "VB022", Key: "detect_confusing_call_syntax", Default: true, Get: func(c LintConfig) bool { return c.DetectConfusingCallSyntax }, Set: func(c *LintConfig, v bool) { c.DetectConfusingCallSyntax = v }},
	{ID: "VB023", Key: "detect_for_each_control_type", Default: true, Get: func(c LintConfig) bool { return c.DetectForEachControlType }, Set: func(c *LintConfig, v bool) { c.DetectForEachControlType = v }},
	{ID: "VB026", Key: "detect_dangerous_resume", Default: true, Get: func(c LintConfig) bool { return c.DetectDangerousResume }, Set: func(c *LintConfig, v bool) { c.DetectDangerousResume = v }},
	{ID: "VB027", Key: "detect_nested_with_ambiguity", Default: false, Get: func(c LintConfig) bool { return c.DetectNestedWithAmbiguity }, Set: func(c *LintConfig, v bool) { c.DetectNestedWithAmbiguity = v }},
}

var configurableAnalyzeRules = []analyzeRuleConfig{
	{ID: "VBA201", Key: "detect_range_find_nothing_check", Default: true, Get: func(c AnalyzeConfig) bool { return c.DetectRangeFindNothingCheck }, Set: func(c *AnalyzeConfig, v bool) { c.DetectRangeFindNothingCheck = v }},
	{ID: "VBA202", Key: "detect_object_use_before_set", Default: true, Get: func(c AnalyzeConfig) bool { return c.DetectObjectUseBeforeSet }, Set: func(c *AnalyzeConfig, v bool) { c.DetectObjectUseBeforeSet = v }},
	{ID: "VBA203", Key: "detect_application_state_restore", Default: true, Get: func(c AnalyzeConfig) bool { return c.DetectApplicationStateRestore }, Set: func(c *AnalyzeConfig, v bool) { c.DetectApplicationStateRestore = v }},
	{ID: "VBA204", Key: "detect_error_handler_fallthrough", Default: true, Get: func(c AnalyzeConfig) bool { return c.DetectErrorHandlerFallthrough }, Set: func(c *AnalyzeConfig, v bool) { c.DetectErrorHandlerFallthrough = v }},
	{ID: "VBA205", Key: "forbid_unqualified_excel_objects", Default: true, Get: func(c AnalyzeConfig) bool { return c.ForbidUnqualifiedExcelObjects }, Set: func(c *AnalyzeConfig, v bool) { c.ForbidUnqualifiedExcelObjects = v }},
	{ID: "VBA206", Key: "detect_byref_argument_mismatch", Default: false, Get: func(c AnalyzeConfig) bool { return c.DetectByRefArgumentMismatch }, Set: func(c *AnalyzeConfig, v bool) { c.DetectByRefArgumentMismatch = v }},
	{ID: "VBA207", Key: "detect_dictionary_collection_guard", Default: false, Get: func(c AnalyzeConfig) bool { return c.DetectDictionaryCollectionGuard }, Set: func(c *AnalyzeConfig, v bool) { c.DetectDictionaryCollectionGuard = v }},
	{ID: "VBA208", Key: "detect_redim_preserve_dimension", Default: true, Get: func(c AnalyzeConfig) bool { return c.DetectRedimPreserveDimension }, Set: func(c *AnalyzeConfig, v bool) { c.DetectRedimPreserveDimension = v }},
	{ID: "VBA209", Key: "detect_object_array_comparison", Default: true, Get: func(c AnalyzeConfig) bool { return c.DetectObjectArrayComparison }, Set: func(c *AnalyzeConfig, v bool) { c.DetectObjectArrayComparison = v }},
	{ID: "VBA210", Key: "detect_function_return_path", Default: false, Get: func(c AnalyzeConfig) bool { return c.DetectFunctionReturnPath }, Set: func(c *AnalyzeConfig, v bool) { c.DetectFunctionReturnPath = v }},
	{ID: "VBA211", Key: "detect_excel_object_member_mismatch", Default: true, Get: func(c AnalyzeConfig) bool { return c.DetectExcelObjectMemberMismatch }, Set: func(c *AnalyzeConfig, v bool) { c.DetectExcelObjectMemberMismatch = v }},
}

var (
	lintRuleByID           = indexLintRulesByID()
	analyzeRuleByID        = indexAnalyzeRulesByID()
	nonConfigurableRuleIDs = map[string]bool{
		"VB008":  true,
		"VB009":  true,
		"VB010":  true,
		"VB011":  true,
		"VB012":  true,
		"VB013":  true,
		"VB014":  true,
		"VB028":  true,
		"VB029":  true,
		"VB031":  true,
		"VB032":  true,
		"VBA101": true,
		"VBA102": true,
		"VBA103": true,
		"VBA104": true,
		"VBA105": true,
		"VBA106": true,
	}
)

func KnownDiagnosticID(id string) bool {
	id = strings.ToUpper(strings.TrimSpace(id))
	if _, ok := lintRuleByID[id]; ok {
		return true
	}
	if _, ok := analyzeRuleByID[id]; ok {
		return true
	}
	return nonConfigurableRuleIDs[id]
}

func LintDiagnosticID(id string) bool {
	id = strings.ToUpper(strings.TrimSpace(id))
	if strings.HasPrefix(id, "VBA") {
		return false
	}
	if _, ok := lintRuleByID[id]; ok {
		return true
	}
	return strings.HasPrefix(id, "VB") && nonConfigurableRuleIDs[id]
}

func AnalyzeDiagnosticID(id string) bool {
	id = strings.ToUpper(strings.TrimSpace(id))
	if _, ok := analyzeRuleByID[id]; ok {
		return true
	}
	return strings.HasPrefix(id, "VBA") && nonConfigurableRuleIDs[id]
}

func InlineSuppressibleDiagnosticID(id string) bool {
	id = strings.ToUpper(strings.TrimSpace(id))
	switch id {
	case "VB008", "VB009", "VB010", "VB011", "VB012", "VB013", "VB014", "VB028", "VB029", "VB031", "VB032",
		"VBA104", "VBA105", "VBA106", "VBA211":
		return false
	default:
		return KnownDiagnosticID(id)
	}
}

func Default() Config {
	return Config{
		Project: ProjectConfig{
			Name:  "sample",
			Entry: "Main.Run",
		},
		Excel: ExcelConfig{
			Path:          filepath.ToSlash(filepath.Join("build", "Book.xlsm")),
			Visible:       false,
			DisplayAlerts: false,
			Bridge:        "auto",
		},
		Src: SourceConfig{
			Modules:  filepath.ToSlash(filepath.Join("src", "modules")),
			Classes:  filepath.ToSlash(filepath.Join("src", "classes")),
			Forms:    filepath.ToSlash(filepath.Join("src", "forms")),
			Workbook: filepath.ToSlash(filepath.Join("src", "workbook")),
		},
		VBA: VBAConfig{
			Folders:                 true,
			FolderAnnotation:        "update",
			DefaultComponentFolders: true,
		},
		UserForm: UserFormConfig{
			CodeSource: "sidecar",
		},
		Fmt: FmtConfig{
			OperatorSpacing:    true,
			DeclarationSpacing: true,
		},
		Lint: LintConfig{
			RequireOptionExplicit:           true,
			ForbidSelect:                    true,
			ForbidActivate:                  true,
			ForbidOnErrorResumeNext:         true,
			DetectImplicitVariant:           true,
			ForbidPublicModuleFields:        true,
			ForbidInteractiveInput:          true,
			DetectMultipleDeclaratorClarity: true,
			DetectUnusedLocalVariables:      true,
			DetectConfusingCallSyntax:       true,
			DetectForEachControlType:        true,
			DetectDangerousResume:           true,
		},
		Analyze: AnalyzeConfig{
			DetectRangeFindNothingCheck:     true,
			DetectObjectUseBeforeSet:        true,
			DetectApplicationStateRestore:   true,
			DetectErrorHandlerFallthrough:   true,
			ForbidUnqualifiedExcelObjects:   true,
			DetectRedimPreserveDimension:    true,
			DetectObjectArrayComparison:     true,
			DetectExcelObjectMemberMismatch: true,
		},
	}
}

func Load(cwd string) (Config, error) {
	return load(cwd, false)
}

func LoadAllowInvalidExcelBridge(cwd string) (Config, error) {
	return load(cwd, true)
}

func load(cwd string, allowInvalidExcelBridge bool) (Config, error) {
	cfg := Default()
	path := filepath.Join(cwd, FileName)
	if _, err := os.Stat(path); err != nil {
		if errors.Is(err, os.ErrNotExist) {
			return cfg, fmt.Errorf("%s not found: %w", FileName, ErrConfigNotFound)
		}
		return cfg, err
	}
	meta, err := toml.DecodeFile(path, &cfg)
	if err != nil {
		return cfg, err
	}
	applyDefaults(&cfg)
	if err := applyLintRuleConfig(&cfg, meta); err != nil {
		return cfg, err
	}
	if err := applyAnalyzeRuleConfig(&cfg, meta); err != nil {
		return cfg, err
	}
	if err := normalizeExcelBridge(&cfg, allowInvalidExcelBridge); err != nil {
		return cfg, err
	}
	return cfg, validate(cfg)
}

func applyDefaults(cfg *Config) {
	defaults := Default()
	if cfg.Project.Name == "" {
		cfg.Project.Name = defaults.Project.Name
	}
	if cfg.Excel.Path == "" {
		cfg.Excel.Path = defaults.Excel.Path
	}
	if cfg.Excel.Bridge == "" {
		cfg.Excel.Bridge = defaults.Excel.Bridge
	}
	if cfg.Src.Modules == "" {
		cfg.Src.Modules = defaults.Src.Modules
	}
	if cfg.Src.Classes == "" {
		cfg.Src.Classes = defaults.Src.Classes
	}
	if cfg.Src.Forms == "" {
		cfg.Src.Forms = defaults.Src.Forms
	}
	if cfg.Src.Workbook == "" {
		cfg.Src.Workbook = defaults.Src.Workbook
	}
	if cfg.VBA.FolderAnnotation == "" {
		cfg.VBA.FolderAnnotation = defaults.VBA.FolderAnnotation
	}
	if cfg.UserForm.CodeSource == "" {
		cfg.UserForm.CodeSource = defaults.UserForm.CodeSource
	}
}

func validate(cfg Config) error {
	if cfg.Project.Entry == "" {
		return errors.New("project.entry is required")
	}
	if cfg.Excel.Path == "" {
		return errors.New("excel.path is required")
	}
	switch cfg.VBA.FolderAnnotation {
	case "update", "preserve", "ignore":
	default:
		return fmt.Errorf("vba.folder_annotation must be one of update, preserve, ignore")
	}
	switch cfg.UserForm.CodeSource {
	case "frm", "sidecar":
	default:
		return fmt.Errorf("userform.code_source must be one of frm, sidecar")
	}
	return nil
}

func indexLintRulesByID() map[string]lintRuleConfig {
	out := make(map[string]lintRuleConfig, len(configurableLintRules))
	for _, rule := range configurableLintRules {
		out[rule.ID] = rule
	}
	return out
}

func indexAnalyzeRulesByID() map[string]analyzeRuleConfig {
	out := make(map[string]analyzeRuleConfig, len(configurableAnalyzeRules))
	for _, rule := range configurableAnalyzeRules {
		out[rule.ID] = rule
	}
	return out
}

func applyLintRuleConfig(cfg *Config, meta toml.MetaData) error {
	disabled, disabledSet, err := normalizeDisabledLintRules(cfg.Lint.DisabledRules)
	if err != nil {
		return err
	}
	cfg.Lint.DisabledRules = disabled
	warnings := make([]map[string]any, 0)
	for _, rule := range configurableLintRules {
		if !meta.IsDefined("lint", rule.Key) {
			continue
		}
		warnings = append(warnings, map[string]any{
			"code":    "deprecated_lint_rule_config",
			"message": deprecatedRuleConfigMessage("lint", rule.Key, rule.ID, rule.Get(cfg.Lint), rule.Default),
			"rule":    rule.ID,
			"key":     rule.Key,
		})
		if rule.Get(cfg.Lint) && disabledSet[rule.ID] {
			warnings = append(warnings,
				map[string]any{
					"code":    "conflicting_lint_rule_config",
					"message": fmt.Sprintf("lint rule %s is enabled by [lint].%s=true but also listed in [lint].disabled_rules.", rule.ID, rule.Key),
					"rule":    rule.ID,
					"key":     rule.Key,
				},
				map[string]any{
					"code":    "disabled_rules_precedence",
					"message": "[lint].disabled_rules takes precedence.",
					"rule":    rule.ID,
					"key":     rule.Key,
				},
			)
		}
	}
	for id := range disabledSet {
		rule := lintRuleByID[id]
		rule.Set(&cfg.Lint, false)
	}
	cfg.Warnings = append(cfg.Warnings, warnings...)
	return nil
}

func normalizeDisabledLintRules(ids []string) ([]string, map[string]bool, error) {
	seen := map[string]bool{}
	out := make([]string, 0, len(ids))
	for _, raw := range ids {
		id := strings.ToUpper(strings.TrimSpace(raw))
		if id == "" {
			continue
		}
		if _, ok := lintRuleByID[id]; !ok {
			if nonConfigurableRuleIDs[id] {
				return nil, nil, fmt.Errorf("lint rule ID is not configurable in [lint].disabled_rules: %s", id)
			}
			return nil, nil, fmt.Errorf("unknown lint rule ID in [lint].disabled_rules: %s", id)
		}
		if seen[id] {
			continue
		}
		seen[id] = true
		out = append(out, id)
	}
	return out, seen, nil
}

func applyAnalyzeRuleConfig(cfg *Config, meta toml.MetaData) error {
	disabled, disabledSet, err := normalizeDisabledAnalyzeRules(cfg.Analyze.DisabledRules)
	if err != nil {
		return err
	}
	cfg.Analyze.DisabledRules = disabled
	warnings := make([]map[string]any, 0)
	for _, rule := range configurableAnalyzeRules {
		if !meta.IsDefined("analyze", rule.Key) {
			continue
		}
		warnings = append(warnings, map[string]any{
			"code":    "deprecated_analyze_rule_config",
			"message": deprecatedRuleConfigMessage("analyze", rule.Key, rule.ID, rule.Get(cfg.Analyze), rule.Default),
			"rule":    rule.ID,
			"key":     rule.Key,
		})
		if rule.Get(cfg.Analyze) && disabledSet[rule.ID] {
			warnings = append(warnings,
				map[string]any{
					"code":    "conflicting_analyze_rule_config",
					"message": fmt.Sprintf("analyze rule %s is enabled by [analyze].%s=true but also listed in [analyze].disabled_rules.", rule.ID, rule.Key),
					"rule":    rule.ID,
					"key":     rule.Key,
				},
				map[string]any{
					"code":    "analyze_disabled_rules_precedence",
					"message": "[analyze].disabled_rules takes precedence.",
					"rule":    rule.ID,
					"key":     rule.Key,
				},
			)
		}
	}
	for id := range disabledSet {
		rule := analyzeRuleByID[id]
		rule.Set(&cfg.Analyze, false)
	}
	cfg.Warnings = append(cfg.Warnings, warnings...)
	return nil
}

func deprecatedRuleConfigMessage(section, key, id string, enabled, defaultEnabled bool) string {
	qualifiedKey := fmt.Sprintf("[%s].%s", section, key)
	switch {
	case !enabled:
		if defaultEnabled {
			return fmt.Sprintf("%s=false is deprecated. Use [%s].disabled_rules = [%q] instead.", qualifiedKey, section, id)
		}
		return fmt.Sprintf("%s=false is deprecated and redundant because %s is disabled by default. Remove %s.", qualifiedKey, id, qualifiedKey)
	case defaultEnabled:
		return fmt.Sprintf("%s=true is deprecated and redundant because %s is enabled by default. Remove %s.", qualifiedKey, id, qualifiedKey)
	default:
		return fmt.Sprintf("%s=true is deprecated but remains the compatibility opt-in for %s; keep it until an opt-in replacement is available.", qualifiedKey, id)
	}
}

func normalizeDisabledAnalyzeRules(ids []string) ([]string, map[string]bool, error) {
	seen := map[string]bool{}
	out := make([]string, 0, len(ids))
	for _, raw := range ids {
		id := strings.ToUpper(strings.TrimSpace(raw))
		if id == "" {
			continue
		}
		if _, ok := analyzeRuleByID[id]; !ok {
			if nonConfigurableRuleIDs[id] {
				return nil, nil, fmt.Errorf("analyze rule ID is not configurable in [analyze].disabled_rules: %s", id)
			}
			return nil, nil, fmt.Errorf("unknown analyze rule ID in [analyze].disabled_rules: %s", id)
		}
		if seen[id] {
			continue
		}
		seen[id] = true
		out = append(out, id)
	}
	return out, seen, nil
}

func normalizeExcelBridge(cfg *Config, allowInvalid bool) error {
	mode, err := excelbridge.ParseMode(cfg.Excel.Bridge)
	if err != nil {
		if allowInvalid {
			return nil
		}
		if errors.Is(err, excelbridge.ErrInvalidMode) {
			return ErrInvalidExcelBridge
		}
		return err
	}
	cfg.Excel.Bridge = string(mode)
	return nil
}

func renderLintConfig(cfg LintConfig) string {
	var b strings.Builder
	b.WriteString("# Disable specific lint rules by diagnostic ID.\n")
	b.WriteString("#\n")
	b.WriteString("# Example:\n")
	b.WriteString("# disabled_rules = [\n")
	b.WriteString("#   \"VB006\", # Allow public module-level fields in this legacy project.\n")
	b.WriteString("# ]\n")
	disabled := disabledLintRuleIDsForWrite(cfg)
	if len(disabled) == 0 {
		b.WriteString("disabled_rules = []\n")
	} else {
		b.WriteString("disabled_rules = [\n")
		for _, id := range disabled {
			b.WriteString("  \"")
			b.WriteString(id)
			b.WriteString("\",\n")
		}
		b.WriteString("]\n")
	}
	optIn := legacyOptInLintRulesForWrite(cfg)
	optInSet := map[string]bool{}
	for _, rule := range optIn {
		optInSet[rule.Key] = true
	}
	b.WriteString("\n")
	b.WriteString("# VB020 unused-local-variable warnings are enabled by default.\n")
	b.WriteString("# Add \"VB020\" to disabled_rules if a project intentionally keeps scratch locals.\n")
	b.WriteString("#\n")
	b.WriteString("# Optional project-wide lint rules. They are disabled by default because\n")
	b.WriteString("# they can be noisy in projects with callback-heavy or workbook-driven VBA.\n")
	b.WriteString("# Uncomment individual rules to enable them.\n")
	for _, hint := range []struct {
		Key  string
		Line string
	}{
		{"detect_scope_shadowing", "# detect_scope_shadowing = true          # VB018\n"},
		{"detect_unused_private_procedures", "# detect_unused_private_procedures = true # VB021\n"},
		{"detect_nested_with_ambiguity", "# detect_nested_with_ambiguity = true    # VB027\n"},
	} {
		if !optInSet[hint.Key] {
			b.WriteString(hint.Line)
		}
	}
	if len(optIn) > 0 {
		b.WriteString("\n")
		b.WriteString("# Enabled optional lint settings.\n")
		for _, rule := range optIn {
			b.WriteString(rule.Key)
			b.WriteString(" = true\n")
		}
	}
	return b.String()
}

func disabledLintRuleIDsForWrite(cfg LintConfig) []string {
	seen := map[string]bool{}
	for _, raw := range cfg.DisabledRules {
		id := strings.ToUpper(strings.TrimSpace(raw))
		if _, ok := lintRuleByID[id]; ok {
			seen[id] = true
		}
	}
	for _, rule := range configurableLintRules {
		if rule.Default && !rule.Get(cfg) {
			seen[rule.ID] = true
		}
	}
	out := make([]string, 0, len(seen))
	for _, rule := range configurableLintRules {
		if seen[rule.ID] {
			out = append(out, rule.ID)
		}
	}
	return out
}

func legacyOptInLintRulesForWrite(cfg LintConfig) []lintRuleConfig {
	var out []lintRuleConfig
	for _, rule := range configurableLintRules {
		if !rule.Default && rule.Get(cfg) {
			out = append(out, rule)
		}
	}
	return out
}

func renderAnalyzeConfig(cfg AnalyzeConfig) string {
	var b strings.Builder
	b.WriteString("# Disable specific analyzer rules by diagnostic ID.\n")
	b.WriteString("#\n")
	b.WriteString("# Example:\n")
	b.WriteString("# disabled_rules = [\n")
	b.WriteString("#   \"VBA205\", # Allow active worksheet dependencies in this legacy project.\n")
	b.WriteString("# ]\n")
	disabled := disabledAnalyzeRuleIDsForWrite(cfg)
	if len(disabled) == 0 {
		b.WriteString("disabled_rules = []\n")
	} else {
		b.WriteString("disabled_rules = [\n")
		for _, id := range disabled {
			b.WriteString("  \"")
			b.WriteString(id)
			b.WriteString("\",\n")
		}
		b.WriteString("]\n")
	}
	optIn := legacyOptInAnalyzeRulesForWrite(cfg)
	if len(optIn) > 0 {
		b.WriteString("\n")
		b.WriteString("# Legacy opt-in analyzer settings. Prefer disabled_rules for disabling recommended rules.\n")
		for _, rule := range optIn {
			b.WriteString(rule.Key)
			b.WriteString(" = true\n")
		}
	}
	return b.String()
}

func disabledAnalyzeRuleIDsForWrite(cfg AnalyzeConfig) []string {
	seen := map[string]bool{}
	for _, raw := range cfg.DisabledRules {
		id := strings.ToUpper(strings.TrimSpace(raw))
		if _, ok := analyzeRuleByID[id]; ok {
			seen[id] = true
		}
	}
	for _, rule := range configurableAnalyzeRules {
		if rule.Default && !rule.Get(cfg) {
			seen[rule.ID] = true
		}
	}
	out := make([]string, 0, len(seen))
	for _, rule := range configurableAnalyzeRules {
		if seen[rule.ID] {
			out = append(out, rule.ID)
		}
	}
	return out
}

func legacyOptInAnalyzeRulesForWrite(cfg AnalyzeConfig) []analyzeRuleConfig {
	var out []analyzeRuleConfig
	for _, rule := range configurableAnalyzeRules {
		if !rule.Default && rule.Get(cfg) {
			out = append(out, rule)
		}
	}
	return out
}

func Write(path string, cfg Config) (err error) {
	f, err := os.OpenFile(path, os.O_WRONLY|os.O_CREATE|os.O_EXCL, 0o644)
	if err != nil {
		return err
	}
	defer func() {
		if closeErr := f.Close(); err == nil && closeErr != nil {
			err = closeErr
		}
	}()

	lintConfigText := renderLintConfig(cfg.Lint)
	analyzeConfigText := renderAnalyzeConfig(cfg.Analyze)

	const tmpl = `# Project identity and entry point.
[project]
# Project name used in output messages. Falls back to the workbook base name.
name = %q
# Default macro invoked by xlflow run when no positional macro is given.
entry = %q

# Excel automation settings.
[excel]
# Path to the workbook, relative to the project root or absolute.
path = %q
# Make the Excel application window visible during automation.
visible = %t
# Suppress Excel alert dialogs (e.g. overwrite confirmations).
display_alerts = %t
# Excel bridge mode. Valid values: "auto", "dotnet".
bridge = %q

# Source tree directories.
[src]
# Directory for standard .bas modules.
modules = %q
# Directory for class .cls modules.
classes = %q
# Directory for UserForm .frm files.
forms = %q
# Directory for workbook document module text.
workbook = %q

# VBE component folder support (Rubberduck-style).
[vba]
# Enable @Folder("A.B") annotations and nested source paths.
folders = %t
# How xlflow handles @Folder annotations during push.
# Valid values: "update", "preserve", "ignore".
#   "update"    – rewrite from source directory layout.
#   "preserve"  – keep existing annotations as-is.
#   "ignore"    – disable folder annotation read/write.
folder_annotation = %q
# Automatically assign default folder annotations based on source paths.
default_component_folders = %t

# UserForm source mode.
[userform]
# Where UserForm code-behind lives in the source tree.
# Valid values: "frm", "sidecar".
#   "frm"     – code is kept inside the exported .frm file.
#   "sidecar" – code is split into src/forms/code/<FormName>.bas.
code_source = %q

# VBA formatter settings.
[fmt]
# Normalize spacing around safe binary operators in xlflow fmt.
operator_spacing = %t
# Normalize spacing in safe VBA declarations in xlflow fmt.
declaration_spacing = %t

# Static analysis rules.
[lint]
%s

# Runtime-risk analysis rules.
[analyze]
%s
`
	_, err = fmt.Fprintf(f, tmpl,
		cfg.Project.Name, cfg.Project.Entry,
		cfg.Excel.Path, cfg.Excel.Visible, cfg.Excel.DisplayAlerts, cfg.Excel.Bridge,
		cfg.Src.Modules, cfg.Src.Classes, cfg.Src.Forms, cfg.Src.Workbook,
		cfg.VBA.Folders, cfg.VBA.FolderAnnotation, cfg.VBA.DefaultComponentFolders,
		cfg.UserForm.CodeSource,
		cfg.Fmt.OperatorSpacing, cfg.Fmt.DeclarationSpacing,
		lintConfigText,
		analyzeConfigText,
	)
	return err
}
