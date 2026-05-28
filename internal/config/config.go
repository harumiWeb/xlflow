package config

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"

	"github.com/BurntSushi/toml"
	excelbridge "github.com/harumiWeb/xlflow/internal/excel/bridge"
)

const FileName = "xlflow.toml"

var ErrInvalidExcelBridge = errors.New("excel.bridge must be one of auto, powershell, dotnet")

type Config struct {
	Project  ProjectConfig  `toml:"project"`
	Excel    ExcelConfig    `toml:"excel"`
	Src      SourceConfig   `toml:"src"`
	VBA      VBAConfig      `toml:"vba"`
	UserForm UserFormConfig `toml:"userform"`
	Lint     LintConfig     `toml:"lint"`
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

type LintConfig struct {
	RequireOptionExplicit    bool `toml:"require_option_explicit"`
	ForbidSelect             bool `toml:"forbid_select"`
	ForbidActivate           bool `toml:"forbid_activate"`
	ForbidOnErrorResumeNext  bool `toml:"forbid_on_error_resume_next"`
	DetectImplicitVariant    bool `toml:"detect_implicit_variant"`
	ForbidPublicModuleFields bool `toml:"forbid_public_module_fields"`
	ForbidInteractiveInput   bool `toml:"forbid_interactive_input"`
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
		Lint: LintConfig{
			RequireOptionExplicit:    true,
			ForbidSelect:             true,
			ForbidActivate:           true,
			ForbidOnErrorResumeNext:  true,
			DetectImplicitVariant:    true,
			ForbidPublicModuleFields: true,
			ForbidInteractiveInput:   true,
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
			return cfg, fmt.Errorf("%s not found", FileName)
		}
		return cfg, err
	}
	if _, err := toml.DecodeFile(path, &cfg); err != nil {
		return cfg, err
	}
	applyDefaults(&cfg)
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
# Excel bridge mode. Valid values: "auto", "powershell", "dotnet".
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

# Static analysis rules.
[lint]
# Require Option Explicit in every module.
require_option_explicit = %t
# Forbid Select / Activate patterns.
forbid_select = %t
# Forbid Activate usage.
forbid_activate = %t
# Forbid On Error Resume Next.
forbid_on_error_resume_next = %t
# Detect implicitly typed Variant variables.
detect_implicit_variant = %t
# Forbid public fields in standard modules.
forbid_public_module_fields = %t
# Forbid interactive input (MsgBox, InputBox, etc.) in headless runs.
forbid_interactive_input = %t
`
	_, err = fmt.Fprintf(f, tmpl,
		cfg.Project.Name, cfg.Project.Entry,
		cfg.Excel.Path, cfg.Excel.Visible, cfg.Excel.DisplayAlerts, cfg.Excel.Bridge,
		cfg.Src.Modules, cfg.Src.Classes, cfg.Src.Forms, cfg.Src.Workbook,
		cfg.VBA.Folders, cfg.VBA.FolderAnnotation, cfg.VBA.DefaultComponentFolders,
		cfg.UserForm.CodeSource,
		cfg.Lint.RequireOptionExplicit, cfg.Lint.ForbidSelect, cfg.Lint.ForbidActivate,
		cfg.Lint.ForbidOnErrorResumeNext, cfg.Lint.DetectImplicitVariant,
		cfg.Lint.ForbidPublicModuleFields, cfg.Lint.ForbidInteractiveInput,
	)
	return err
}
