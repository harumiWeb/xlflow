package config

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"unicode"

	"github.com/BurntSushi/toml"
	excelbridge "github.com/harumiWeb/xlflow/internal/excel/bridge"
	staticrules "github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
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
	Build    BuildConfig      `toml:"build"`
	Backup   BackupConfig     `toml:"backup"`
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
	Folders                 bool                 `toml:"folders"`
	FolderAnnotation        string               `toml:"folder_annotation"`
	DefaultComponentFolders bool                 `toml:"default_component_folders"`
	LineNumbers             VBALineNumbersConfig `toml:"line_numbers"`
}

// VBALineNumbersConfig controls temporary Erl-compatible instrumentation used
// only while source is imported into an Excel workbook.
type VBALineNumbersConfig struct {
	Enabled bool `toml:"enabled"`
}

type UserFormConfig struct {
	CodeSource string `toml:"code_source"`
}

// BuildConfig controls release-build source selection. It is intentionally
// separate from push, which always synchronizes the complete source tree.
type BuildConfig struct {
	Exclude []string `toml:"exclude"`
}

type BackupConfig struct {
	Retention BackupRetentionConfig `toml:"retention"`
}

type BackupRetentionConfig struct {
	Enabled        bool `toml:"enabled"`
	MaxCount       int  `toml:"max_count"`
	MaxAgeDays     int  `toml:"max_age_days"`
	MinKeep        int  `toml:"min_keep"`
	MaxTotalSizeMB int  `toml:"max_total_size_mb"`
}

type FmtConfig struct {
	OperatorSpacing    bool `toml:"operator_spacing"`
	DeclarationSpacing bool `toml:"declaration_spacing"`
	KeywordCasing      bool `toml:"keyword_casing"`
	BuiltinCasing      bool `toml:"builtin_casing"`
}

type LintConfig struct {
	DisabledRules                   []string                    `toml:"disabled_rules"`
	RequireOptionExplicit           bool                        `toml:"require_option_explicit"`
	ForbidSelect                    bool                        `toml:"forbid_select"`
	ForbidActivate                  bool                        `toml:"forbid_activate"`
	ForbidOnErrorResumeNext         bool                        `toml:"forbid_on_error_resume_next"`
	DetectImplicitVariant           bool                        `toml:"detect_implicit_variant"`
	ForbidPublicModuleFields        bool                        `toml:"forbid_public_module_fields"`
	ForbidInteractiveInput          bool                        `toml:"forbid_interactive_input"`
	DetectScopeShadowing            bool                        `toml:"detect_scope_shadowing"`
	DetectMultipleDeclaratorClarity bool                        `toml:"detect_multiple_declarator_clarity"`
	DetectUnusedLocalVariables      bool                        `toml:"detect_unused_local_variables"`
	DetectUnusedPrivateProcedures   bool                        `toml:"detect_unused_private_procedures"`
	DetectConfusingCallSyntax       bool                        `toml:"detect_confusing_call_syntax"`
	DetectForEachControlType        bool                        `toml:"detect_for_each_control_type"`
	DetectDangerousResume           bool                        `toml:"detect_dangerous_resume"`
	DetectNestedWithAmbiguity       bool                        `toml:"detect_nested_with_ambiguity"`
	ProcedureNameConstant           ProcedureNameConstantConfig `toml:"procedure_name_constant"`
}

// ProcedureNameConstantConfig controls the opt-in check that keeps a local
// procedure-name constant aligned with its enclosing VBA procedure.
type ProcedureNameConstantConfig struct {
	Enabled      bool   `toml:"enabled"`
	ConstantName string `toml:"constant_name"`
}

type AnalyzeConfig struct {
	DisabledRules                            []string `toml:"disabled_rules"`
	DetectRangeFindNothingCheck              bool     `toml:"detect_range_find_nothing_check"`
	DetectObjectUseBeforeSet                 bool     `toml:"detect_object_use_before_set"`
	DetectApplicationStateRestore            bool     `toml:"detect_application_state_restore"`
	DetectErrorHandlerFallthrough            bool     `toml:"detect_error_handler_fallthrough"`
	ForbidUnqualifiedExcelObjects            bool     `toml:"forbid_unqualified_excel_objects"`
	DetectByRefArgumentMismatch              bool     `toml:"detect_byref_argument_mismatch"`
	DetectDictionaryCollectionGuard          bool     `toml:"detect_dictionary_collection_guard"`
	DetectRedimPreserveDimension             bool     `toml:"detect_redim_preserve_dimension"`
	DetectObjectArrayComparison              bool     `toml:"detect_object_array_comparison"`
	DetectFunctionReturnPath                 bool     `toml:"detect_function_return_path"`
	DetectExcelObjectMemberMismatch          bool     `toml:"detect_excel_object_member_mismatch"`
	DetectNonShortCircuitObjectGuard         bool     `toml:"detect_non_short_circuit_object_guard"`
	DetectDictionaryIterationValueUsage      bool     `toml:"detect_dictionary_iteration_value_usage"`
	DetectLeakedOnErrorResumeNextScopes      bool     `toml:"detect_leaked_on_error_resume_next_scopes"`
	DetectStatefulExcelCallArguments         bool     `toml:"detect_stateful_excel_call_arguments"`
	DetectWorksheetRootMismatch              bool     `toml:"detect_worksheet_root_mismatch"`
	DetectUnstableLastRowPatterns            bool     `toml:"detect_unstable_last_row_patterns"`
	DetectExcelAPIFailureContracts           bool     `toml:"detect_excel_api_failure_contracts"`
	DetectResourceLeaks                      bool     `toml:"detect_resource_leaks"`
	DetectEventHandlerReentry                bool     `toml:"detect_event_handler_reentry"`
	DetectApplicationStateCallEffects        bool     `toml:"detect_application_state_call_effects"`
	DetectHardcodedSecrets                   bool     `toml:"detect_hardcoded_secrets"`
	DetectPublicAPITypeSafety                bool     `toml:"detect_public_api_type_safety"`
	DetectUntrustedDataFlow                  bool     `toml:"detect_untrusted_data_flow"`
	DetectExcelCellAccessInLoops             bool     `toml:"detect_excel_cell_access_in_loops"`
	DetectRangeValueArrayShape               bool     `toml:"detect_range_value_array_shape"`
	DetectArrayLifecycleSafety               bool     `toml:"detect_array_lifecycle_safety"`
	DetectDictionaryCompareModeOrder         bool     `toml:"detect_dictionary_compare_mode_order"`
	DetectDictionaryLoopMaterialization      bool     `toml:"detect_dictionary_loop_materialization"`
	DetectDictionaryKeyNormalization         bool     `toml:"detect_dictionary_key_normalization"`
	DetectLateBoundDictionaryConstants       bool     `toml:"detect_late_bound_dictionary_constants"`
	DetectCollectionIterationMutation        bool     `toml:"detect_collection_iteration_mutation"`
	DetectCollectionIndexOrigin              bool     `toml:"detect_collection_index_origin"`
	DetectUnsafeCommandConstruction          bool     `toml:"detect_unsafe_command_construction"`
	DetectErrorSuppressionPropagation        bool     `toml:"detect_error_suppression_propagation"`
	DetectLoopInvariantExcelObjectResolution bool     `toml:"detect_loop_invariant_excel_object_resolution"`
	DetectUnsafeSQLConstruction              bool     `toml:"detect_unsafe_sql_construction"`
	DetectRiskyModuleState                   bool     `toml:"detect_risky_module_state"`
	DetectRedimPreserveInLoops               bool     `toml:"detect_redim_preserve_in_loops"`
	DetectExpensiveFullRangeOperations       bool     `toml:"detect_expensive_full_range_operations"`
}

type lintRuleAdapter struct {
	Get func(LintConfig) bool
	Set func(*LintConfig, bool)
}

type analyzeRuleAdapter struct {
	Get func(AnalyzeConfig) bool
	Set func(*AnalyzeConfig, bool)
}

// lintRuleConfig and analyzeRuleConfig are derived views. Rule identity,
// defaults, and configuration keys are owned by the shared registry; only the
// config-type-specific accessors remain here.
type lintRuleConfig struct {
	ID      string
	Key     string
	Default bool
	lintRuleAdapter
}

type analyzeRuleConfig struct {
	ID      string
	Key     string
	Default bool
	analyzeRuleAdapter
}

var lintRuleAdapters = map[string]lintRuleAdapter{
	"VB001": {Get: func(c LintConfig) bool { return c.RequireOptionExplicit }, Set: func(c *LintConfig, v bool) { c.RequireOptionExplicit = v }},
	"VB002": {Get: func(c LintConfig) bool { return c.ForbidSelect }, Set: func(c *LintConfig, v bool) { c.ForbidSelect = v }},
	"VB003": {Get: func(c LintConfig) bool { return c.ForbidActivate }, Set: func(c *LintConfig, v bool) { c.ForbidActivate = v }},
	"VB004": {Get: func(c LintConfig) bool { return c.ForbidOnErrorResumeNext }, Set: func(c *LintConfig, v bool) { c.ForbidOnErrorResumeNext = v }},
	"VB005": {Get: func(c LintConfig) bool { return c.DetectImplicitVariant }, Set: func(c *LintConfig, v bool) { c.DetectImplicitVariant = v }},
	"VB006": {Get: func(c LintConfig) bool { return c.ForbidPublicModuleFields }, Set: func(c *LintConfig, v bool) { c.ForbidPublicModuleFields = v }},
	"VB007": {Get: func(c LintConfig) bool { return c.ForbidInteractiveInput }, Set: func(c *LintConfig, v bool) { c.ForbidInteractiveInput = v }},
	"VB018": {Get: func(c LintConfig) bool { return c.DetectScopeShadowing }, Set: func(c *LintConfig, v bool) { c.DetectScopeShadowing = v }},
	"VB019": {Get: func(c LintConfig) bool { return c.DetectMultipleDeclaratorClarity }, Set: func(c *LintConfig, v bool) { c.DetectMultipleDeclaratorClarity = v }},
	"VB020": {Get: func(c LintConfig) bool { return c.DetectUnusedLocalVariables }, Set: func(c *LintConfig, v bool) { c.DetectUnusedLocalVariables = v }},
	"VB021": {Get: func(c LintConfig) bool { return c.DetectUnusedPrivateProcedures }, Set: func(c *LintConfig, v bool) { c.DetectUnusedPrivateProcedures = v }},
	"VB022": {Get: func(c LintConfig) bool { return c.DetectConfusingCallSyntax }, Set: func(c *LintConfig, v bool) { c.DetectConfusingCallSyntax = v }},
	"VB023": {Get: func(c LintConfig) bool { return c.DetectForEachControlType }, Set: func(c *LintConfig, v bool) { c.DetectForEachControlType = v }},
	"VB026": {Get: func(c LintConfig) bool { return c.DetectDangerousResume }, Set: func(c *LintConfig, v bool) { c.DetectDangerousResume = v }},
	"VB027": {Get: func(c LintConfig) bool { return c.DetectNestedWithAmbiguity }, Set: func(c *LintConfig, v bool) { c.DetectNestedWithAmbiguity = v }},
	"VB044": {Get: func(c LintConfig) bool { return c.ProcedureNameConstant.Enabled }, Set: func(c *LintConfig, v bool) { c.ProcedureNameConstant.Enabled = v }},
}

var analyzeRuleAdapters = map[string]analyzeRuleAdapter{
	"VBA201": {Get: func(c AnalyzeConfig) bool { return c.DetectRangeFindNothingCheck }, Set: func(c *AnalyzeConfig, v bool) { c.DetectRangeFindNothingCheck = v }},
	"VBA202": {Get: func(c AnalyzeConfig) bool { return c.DetectObjectUseBeforeSet }, Set: func(c *AnalyzeConfig, v bool) { c.DetectObjectUseBeforeSet = v }},
	"VBA203": {Get: func(c AnalyzeConfig) bool { return c.DetectApplicationStateRestore }, Set: func(c *AnalyzeConfig, v bool) { c.DetectApplicationStateRestore = v }},
	"VBA204": {Get: func(c AnalyzeConfig) bool { return c.DetectErrorHandlerFallthrough }, Set: func(c *AnalyzeConfig, v bool) { c.DetectErrorHandlerFallthrough = v }},
	"VBA205": {Get: func(c AnalyzeConfig) bool { return c.ForbidUnqualifiedExcelObjects }, Set: func(c *AnalyzeConfig, v bool) { c.ForbidUnqualifiedExcelObjects = v }},
	"VBA206": {Get: func(c AnalyzeConfig) bool { return c.DetectByRefArgumentMismatch }, Set: func(c *AnalyzeConfig, v bool) { c.DetectByRefArgumentMismatch = v }},
	"VBA207": {Get: func(c AnalyzeConfig) bool { return c.DetectDictionaryCollectionGuard }, Set: func(c *AnalyzeConfig, v bool) { c.DetectDictionaryCollectionGuard = v }},
	"VBA208": {Get: func(c AnalyzeConfig) bool { return c.DetectRedimPreserveDimension }, Set: func(c *AnalyzeConfig, v bool) { c.DetectRedimPreserveDimension = v }},
	"VBA209": {Get: func(c AnalyzeConfig) bool { return c.DetectObjectArrayComparison }, Set: func(c *AnalyzeConfig, v bool) { c.DetectObjectArrayComparison = v }},
	"VBA210": {Get: func(c AnalyzeConfig) bool { return c.DetectFunctionReturnPath }, Set: func(c *AnalyzeConfig, v bool) { c.DetectFunctionReturnPath = v }},
	"VBA211": {Get: func(c AnalyzeConfig) bool { return c.DetectExcelObjectMemberMismatch }, Set: func(c *AnalyzeConfig, v bool) { c.DetectExcelObjectMemberMismatch = v }},
	"VBA212": {Get: func(c AnalyzeConfig) bool { return c.DetectNonShortCircuitObjectGuard }, Set: func(c *AnalyzeConfig, v bool) { c.DetectNonShortCircuitObjectGuard = v }},
	"VBA213": {Get: func(c AnalyzeConfig) bool { return c.DetectDictionaryIterationValueUsage }, Set: func(c *AnalyzeConfig, v bool) { c.DetectDictionaryIterationValueUsage = v }},
	"VBA214": {Get: func(c AnalyzeConfig) bool { return c.DetectLeakedOnErrorResumeNextScopes }, Set: func(c *AnalyzeConfig, v bool) { c.DetectLeakedOnErrorResumeNextScopes = v }},
	"VBA215": {Get: func(c AnalyzeConfig) bool { return c.DetectStatefulExcelCallArguments }, Set: func(c *AnalyzeConfig, v bool) { c.DetectStatefulExcelCallArguments = v }},
	"VBA216": {Get: func(c AnalyzeConfig) bool { return c.DetectWorksheetRootMismatch }, Set: func(c *AnalyzeConfig, v bool) { c.DetectWorksheetRootMismatch = v }},
	"VBA217": {Get: func(c AnalyzeConfig) bool { return c.DetectUnstableLastRowPatterns }, Set: func(c *AnalyzeConfig, v bool) { c.DetectUnstableLastRowPatterns = v }},
	"VBA218": {Get: func(c AnalyzeConfig) bool { return c.DetectExcelAPIFailureContracts }, Set: func(c *AnalyzeConfig, v bool) { c.DetectExcelAPIFailureContracts = v }},
	"VBA219": {Get: func(c AnalyzeConfig) bool { return c.DetectResourceLeaks }, Set: func(c *AnalyzeConfig, v bool) { c.DetectResourceLeaks = v }},
	"VBA220": {Get: func(c AnalyzeConfig) bool { return c.DetectEventHandlerReentry }, Set: func(c *AnalyzeConfig, v bool) { c.DetectEventHandlerReentry = v }},
	"VBA221": {Get: func(c AnalyzeConfig) bool { return c.DetectApplicationStateCallEffects }, Set: func(c *AnalyzeConfig, v bool) { c.DetectApplicationStateCallEffects = v }},
	"VBA222": {Get: func(c AnalyzeConfig) bool { return c.DetectPublicAPITypeSafety }, Set: func(c *AnalyzeConfig, v bool) { c.DetectPublicAPITypeSafety = v }},
	"VBA223": {Get: func(c AnalyzeConfig) bool { return c.DetectHardcodedSecrets }, Set: func(c *AnalyzeConfig, v bool) { c.DetectHardcodedSecrets = v }},
	"VBA224": {Get: func(c AnalyzeConfig) bool { return c.DetectUntrustedDataFlow }, Set: func(c *AnalyzeConfig, v bool) { c.DetectUntrustedDataFlow = v }},
	"VBA225": {Get: func(c AnalyzeConfig) bool { return c.DetectExcelCellAccessInLoops }, Set: func(c *AnalyzeConfig, v bool) { c.DetectExcelCellAccessInLoops = v }},
	"VBA226": {Get: func(c AnalyzeConfig) bool { return c.DetectRangeValueArrayShape }, Set: func(c *AnalyzeConfig, v bool) { c.DetectRangeValueArrayShape = v }},
	"VBA227": {Get: func(c AnalyzeConfig) bool { return c.DetectArrayLifecycleSafety }, Set: func(c *AnalyzeConfig, v bool) { c.DetectArrayLifecycleSafety = v }},
	"VBA230": {Get: func(c AnalyzeConfig) bool { return c.DetectDictionaryCompareModeOrder }, Set: func(c *AnalyzeConfig, v bool) { c.DetectDictionaryCompareModeOrder = v }},
	"VBA231": {Get: func(c AnalyzeConfig) bool { return c.DetectDictionaryLoopMaterialization }, Set: func(c *AnalyzeConfig, v bool) { c.DetectDictionaryLoopMaterialization = v }},
	"VBA232": {Get: func(c AnalyzeConfig) bool { return c.DetectDictionaryKeyNormalization }, Set: func(c *AnalyzeConfig, v bool) { c.DetectDictionaryKeyNormalization = v }},
	"VBA233": {Get: func(c AnalyzeConfig) bool { return c.DetectLateBoundDictionaryConstants }, Set: func(c *AnalyzeConfig, v bool) { c.DetectLateBoundDictionaryConstants = v }},
	"VBA234": {Get: func(c AnalyzeConfig) bool { return c.DetectCollectionIterationMutation }, Set: func(c *AnalyzeConfig, v bool) { c.DetectCollectionIterationMutation = v }},
	"VBA235": {Get: func(c AnalyzeConfig) bool { return c.DetectCollectionIndexOrigin }, Set: func(c *AnalyzeConfig, v bool) { c.DetectCollectionIndexOrigin = v }},
	"VBA236": {Get: func(c AnalyzeConfig) bool { return c.DetectUnsafeCommandConstruction }, Set: func(c *AnalyzeConfig, v bool) { c.DetectUnsafeCommandConstruction = v }},
	"VBA237": {Get: func(c AnalyzeConfig) bool { return c.DetectErrorSuppressionPropagation }, Set: func(c *AnalyzeConfig, v bool) { c.DetectErrorSuppressionPropagation = v }},
	"VBA238": {Get: func(c AnalyzeConfig) bool { return c.DetectLoopInvariantExcelObjectResolution }, Set: func(c *AnalyzeConfig, v bool) { c.DetectLoopInvariantExcelObjectResolution = v }},
	"VBA239": {Get: func(c AnalyzeConfig) bool { return c.DetectUnsafeSQLConstruction }, Set: func(c *AnalyzeConfig, v bool) { c.DetectUnsafeSQLConstruction = v }},
	"VBA240": {Get: func(c AnalyzeConfig) bool { return c.DetectRiskyModuleState }, Set: func(c *AnalyzeConfig, v bool) { c.DetectRiskyModuleState = v }},
	"VBA241": {Get: func(c AnalyzeConfig) bool { return c.DetectRedimPreserveInLoops }, Set: func(c *AnalyzeConfig, v bool) { c.DetectRedimPreserveInLoops = v }},
	"VBA242": {Get: func(c AnalyzeConfig) bool { return c.DetectExpensiveFullRangeOperations }, Set: func(c *AnalyzeConfig, v bool) { c.DetectExpensiveFullRangeOperations = v }},
}

var (
	configurableLintRules    = bindLintRules()
	configurableAnalyzeRules = bindAnalyzeRules()
	lintRuleByID             = indexLintRulesByID()
	analyzeRuleByID          = indexAnalyzeRulesByID()
)

func KnownDiagnosticID(id string) bool {
	_, ok := staticrules.Lookup(id)
	return ok
}

func LintDiagnosticID(id string) bool {
	rule, ok := staticrules.Lookup(id)
	return ok && rule.Family == staticrules.FamilyLint
}

func AnalyzeDiagnosticID(id string) bool {
	rule, ok := staticrules.Lookup(id)
	return ok && rule.Family == staticrules.FamilyAnalyze
}

func InlineSuppressibleDiagnosticID(id string) bool {
	rule, ok := staticrules.Lookup(id)
	return ok && rule.InlineSuppressible
}

func Default() Config {
	cfg := Config{
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
		Backup: BackupConfig{
			Retention: BackupRetentionConfig{
				Enabled:        false,
				MaxCount:       20,
				MaxAgeDays:     30,
				MinKeep:        5,
				MaxTotalSizeMB: 2048,
			},
		},
		Fmt: FmtConfig{
			OperatorSpacing:    true,
			DeclarationSpacing: true,
			KeywordCasing:      true,
			BuiltinCasing:      true,
		},
	}
	for _, rule := range configurableLintRules {
		rule.Set(&cfg.Lint, rule.Default)
	}
	for _, rule := range configurableAnalyzeRules {
		rule.Set(&cfg.Analyze, rule.Default)
	}
	return cfg
}

// AnalyzeRuleEnabled reads a configurable analyzer rule through the config
// adapter boundary. The registry deliberately does not depend on Config.
func AnalyzeRuleEnabled(cfg AnalyzeConfig, id string) (bool, bool) {
	rule, ok := analyzeRuleByID[strings.ToUpper(strings.TrimSpace(id))]
	if !ok {
		return false, false
	}
	return rule.Get(cfg), true
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
	cfg.Lint.ProcedureNameConstant.ConstantName = strings.TrimSpace(cfg.Lint.ProcedureNameConstant.ConstantName)
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
	if cfg.Backup.Retention.MaxCount < 0 {
		return errors.New("backup.retention.max_count must be zero or greater")
	}
	if cfg.Backup.Retention.MaxAgeDays < 0 {
		return errors.New("backup.retention.max_age_days must be zero or greater")
	}
	if cfg.Backup.Retention.MinKeep < 0 {
		return errors.New("backup.retention.min_keep must be zero or greater")
	}
	if cfg.Backup.Retention.MaxTotalSizeMB < 0 {
		return errors.New("backup.retention.max_total_size_mb must be zero or greater")
	}
	if cfg.Backup.Retention.MaxCount > 0 && cfg.Backup.Retention.MinKeep > cfg.Backup.Retention.MaxCount {
		return errors.New("backup.retention.min_keep must be less than or equal to backup.retention.max_count when max_count is enabled")
	}
	if cfg.Lint.ProcedureNameConstant.Enabled {
		name := strings.TrimSpace(cfg.Lint.ProcedureNameConstant.ConstantName)
		if name == "" {
			return errors.New("lint.procedure_name_constant.constant_name is required when enabled")
		}
		if !validVBAIdentifier(name) {
			return fmt.Errorf("lint.procedure_name_constant.constant_name must be a VBA identifier: %q", name)
		}
	}
	return nil
}

func validVBAIdentifier(name string) bool {
	runes := []rune(name)
	if len(runes) == 0 || len(runes) > 255 || !unicode.IsLetter(runes[0]) {
		return false
	}
	for _, r := range runes[1:] {
		if r != '_' && !unicode.IsLetter(r) && !unicode.IsDigit(r) {
			return false
		}
	}
	return true
}

func indexLintRulesByID() map[string]lintRuleConfig {
	out := make(map[string]lintRuleConfig, len(configurableLintRules))
	for _, rule := range configurableLintRules {
		out[rule.ID] = rule
	}
	return out
}

func bindLintRules() []lintRuleConfig {
	metadata := staticrules.ByFamily(staticrules.FamilyLint)
	out := make([]lintRuleConfig, 0, len(lintRuleAdapters))
	for _, rule := range metadata {
		if !rule.Configurable {
			continue
		}
		adapter, ok := lintRuleAdapters[rule.ID]
		if !ok {
			panic(fmt.Sprintf("missing lint config adapter for %s", rule.ID))
		}
		out = append(out, lintRuleConfig{
			ID: rule.ID, Key: rule.ConfigurationKey, Default: rule.DefaultEnabled, lintRuleAdapter: adapter,
		})
	}
	if len(out) != len(lintRuleAdapters) {
		panic("lint config adapters do not match configurable rule metadata")
	}
	return out
}

func bindAnalyzeRules() []analyzeRuleConfig {
	metadata := staticrules.ByFamily(staticrules.FamilyAnalyze)
	out := make([]analyzeRuleConfig, 0, len(analyzeRuleAdapters))
	for _, rule := range metadata {
		if !rule.Configurable {
			continue
		}
		adapter, ok := analyzeRuleAdapters[rule.ID]
		if !ok {
			panic(fmt.Sprintf("missing analyze config adapter for %s", rule.ID))
		}
		out = append(out, analyzeRuleConfig{
			ID: rule.ID, Key: rule.ConfigurationKey, Default: rule.DefaultEnabled, analyzeRuleAdapter: adapter,
		})
	}
	if len(out) != len(analyzeRuleAdapters) {
		panic("analyze config adapters do not match configurable rule metadata")
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
		if rule.ID == "VB044" {
			continue
		}
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
	if cfg.Lint.ProcedureNameConstant.Enabled && disabledSet["VB044"] {
		warnings = append(warnings,
			map[string]any{
				"code":    "conflicting_lint_rule_config",
				"message": "lint rule VB044 is enabled by [lint.procedure_name_constant] but also listed in [lint].disabled_rules.",
				"rule":    "VB044",
				"key":     "procedure_name_constant",
			},
			map[string]any{
				"code":    "disabled_rules_precedence",
				"message": "[lint].disabled_rules takes precedence.",
				"rule":    "VB044",
				"key":     "procedure_name_constant",
			},
		)
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
			if rule, known := staticrules.Lookup(id); known && rule.Family == staticrules.FamilyLint && !rule.Configurable {
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
			if rule, known := staticrules.Lookup(id); known && rule.Family == staticrules.FamilyAnalyze && !rule.Configurable {
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
	if cfg.ProcedureNameConstant.Enabled {
		b.WriteString("\n[lint.procedure_name_constant]\n")
		b.WriteString("enabled = true\n")
		b.WriteString("constant_name = ")
		b.WriteString(strconv.Quote(cfg.ProcedureNameConstant.ConstantName))
		b.WriteString("\n")
	} else {
		b.WriteString("\n# Optional local procedure-name constant check (VB044).\n")
		b.WriteString("# [lint.procedure_name_constant]\n")
		b.WriteString("# enabled = true\n")
		b.WriteString("# constant_name = \"PROCEDURE_NAME\"\n")
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
		if rule.ID == "VB044" {
			continue
		}
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
	optInSet := map[string]bool{}
	for _, rule := range optIn {
		optInSet[rule.Key] = true
	}
	if !optInSet["detect_function_return_path"] {
		b.WriteString("\n")
		b.WriteString("# Optional dataflow-sensitive analyzer rules are disabled by default.\n")
		b.WriteString("# Uncomment the following setting to check Function and Property Get return paths.\n")
		b.WriteString("# detect_function_return_path = true # VBA210\n")
	}
	if len(optIn) > 0 {
		b.WriteString("\n# Enabled optional analyzer settings.\n")
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

# Optional Erl instrumentation. When enabled, push numbers only temporary
# import copies; tracked source remains unnumbered.
# [vba.line_numbers]
# enabled = true

# UserForm source mode.
[userform]
# Where UserForm code-behind lives in the source tree.
# Valid values: "frm", "sidecar".
#   "frm"     – code is kept inside the exported .frm file.
#   "sidecar" – code is split into src/forms/code/<FormName>.bas.
code_source = %q

# Automatic backup retention is disabled by default. Uncomment to prune old
# metadata-backed backups for the configured workbook after successful backup-
# producing push and rollback operations.
# [backup.retention]
# enabled = false
# max_count = 20
# max_age_days = 30
# min_keep = 5
# max_total_size_mb = 2048

# VBA formatter settings.
[fmt]
# Normalize spacing around safe binary operators in xlflow fmt.
operator_spacing = %t
# Normalize spacing in safe VBA declarations in xlflow fmt.
declaration_spacing = %t
# Normalize VBA keyword casing in xlflow fmt.
keyword_casing = %t
# Normalize known VBA/Excel/Office built-in identifier casing in xlflow fmt.
builtin_casing = %t

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
		cfg.Fmt.OperatorSpacing, cfg.Fmt.DeclarationSpacing, cfg.Fmt.KeywordCasing, cfg.Fmt.BuiltinCasing,
		lintConfigText,
		analyzeConfigText,
	)
	return err
}

func UpdateUserFormCodeSource(path string, codeSource string) error {
	codeSource = strings.TrimSpace(codeSource)
	switch codeSource {
	case "frm", "sidecar":
	default:
		return fmt.Errorf("userform.code_source must be one of frm, sidecar")
	}
	body, err := os.ReadFile(path)
	if err != nil {
		return err
	}
	text := string(body)
	newline := "\n"
	if strings.Contains(text, "\r\n") {
		newline = "\r\n"
	}
	hadTrailingNewline := strings.HasSuffix(text, "\n") || strings.HasSuffix(text, "\r")
	normalized := strings.ReplaceAll(strings.ReplaceAll(text, "\r\n", "\n"), "\r", "\n")
	lines := strings.Split(normalized, "\n")
	if len(lines) > 0 && lines[len(lines)-1] == "" && hadTrailingNewline {
		lines = lines[:len(lines)-1]
	}

	userFormStart := -1
	userFormEnd := len(lines)
	for i, line := range lines {
		section, ok := tomlSectionName(line)
		if !ok {
			continue
		}
		if userFormStart >= 0 {
			userFormEnd = i
			break
		}
		if section == "userform" {
			userFormStart = i
		}
	}

	replacement := fmt.Sprintf("code_source = %q", codeSource)
	if userFormStart >= 0 {
		for i := userFormStart + 1; i < userFormEnd; i++ {
			if tomlKeyName(lines[i]) == "code_source" {
				lines[i] = replacement
				return os.WriteFile(path, []byte(joinConfigLines(lines, newline, hadTrailingNewline)), 0o644)
			}
		}
		insertAt := userFormEnd
		lines = append(lines, "")
		copy(lines[insertAt+1:], lines[insertAt:])
		lines[insertAt] = replacement
		return os.WriteFile(path, []byte(joinConfigLines(lines, newline, hadTrailingNewline)), 0o644)
	}

	if len(lines) > 0 && strings.TrimSpace(lines[len(lines)-1]) != "" {
		lines = append(lines, "")
	}
	lines = append(lines, "[userform]", replacement)
	return os.WriteFile(path, []byte(joinConfigLines(lines, newline, true)), 0o644)
}

func tomlSectionName(line string) (string, bool) {
	trimmed := strings.TrimSpace(line)
	if !strings.HasPrefix(trimmed, "[") {
		return "", false
	}
	end := strings.Index(trimmed, "]")
	if end < 0 {
		return "", false
	}
	rest := strings.TrimSpace(trimmed[end+1:])
	if rest != "" && !strings.HasPrefix(rest, "#") {
		return "", false
	}
	section := strings.TrimSpace(trimmed[1:end])
	if section == "" || strings.HasPrefix(section, "[") || strings.Contains(section, "]") {
		return "", false
	}
	return section, true
}

func tomlKeyName(line string) string {
	trimmed := strings.TrimSpace(line)
	if trimmed == "" || strings.HasPrefix(trimmed, "#") {
		return ""
	}
	key, _, ok := strings.Cut(trimmed, "=")
	if !ok {
		return ""
	}
	key = strings.TrimSpace(key)
	if unquoted, err := strconv.Unquote(key); err == nil {
		return unquoted
	}
	return key
}

func joinConfigLines(lines []string, newline string, trailing bool) string {
	out := strings.Join(lines, newline)
	if trailing && !strings.HasSuffix(out, newline) {
		out += newline
	}
	return out
}
