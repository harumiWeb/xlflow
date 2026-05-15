package config

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"

	"github.com/BurntSushi/toml"
)

const FileName = "xlflow.toml"

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
	enc := toml.NewEncoder(f)
	return enc.Encode(cfg)
}
