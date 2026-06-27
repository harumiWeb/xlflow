package project

import (
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"gopkg.in/yaml.v3"
)

type ModuleRemoveResult struct {
	Operation    string   `json:"operation"`
	Module       string   `json:"module"`
	Kind         string   `json:"kind"`
	Removed      []string `json:"removed"`
	RequiresPush bool     `json:"requires_push"`
}

type ModuleRenameResult struct {
	Operation    string              `json:"operation"`
	OldName      string              `json:"old_name"`
	NewName      string              `json:"new_name"`
	Kind         string              `json:"kind"`
	Renamed      []ModuleRenameEntry `json:"renamed"`
	RequiresPush bool                `json:"requires_push"`
}

type ModuleRenameEntry struct {
	From string `json:"from"`
	To   string `json:"to"`
}

var (
	ErrModuleNotFound      = errors.New("module not found")
	ErrModuleAmbiguous     = errors.New("module name is ambiguous")
	ErrProtectedModule     = errors.New("protected module")
	ErrModuleAlreadyExists = errors.New("module already exists")
)

type sourceModule struct {
	Name      string
	Kind      string
	Path      string
	Protected bool
}

type sourceRoot struct {
	Path       string
	Kind       string
	Exts       map[string]bool
	Protected  bool
	SkipSubdir map[string]bool
}

type renameOp struct {
	OldPath     string
	NewPath     string
	OldContent  []byte
	NewContent  []byte
	UpdateBytes bool
}

func RemoveModule(cwd, name string, src config.SourceConfig) (ModuleRemoveResult, error) {
	var result ModuleRemoveResult
	module, err := resolveSourceModule(cwd, src, name)
	if err != nil {
		return result, err
	}
	if module.Protected {
		return result, fmt.Errorf("%w: document module %q cannot be removed", ErrProtectedModule, module.Name)
	}
	paths, err := moduleArtifactPaths(cwd, src, module)
	if err != nil {
		return result, err
	}
	for _, path := range paths {
		if err := os.Remove(path); err != nil {
			if errors.Is(err, os.ErrNotExist) {
				continue
			}
			return result, err
		}
		result.Removed = append(result.Removed, filepath.ToSlash(rel(cwd, path)))
	}
	result.Operation = "module.remove"
	result.Module = module.Name
	result.Kind = module.Kind
	result.RequiresPush = true
	return result, nil
}

func RenameModule(cwd, oldName, newName string, src config.SourceConfig) (ModuleRenameResult, error) {
	var result ModuleRenameResult
	module, err := resolveSourceModule(cwd, src, oldName)
	if err != nil {
		return result, err
	}
	if module.Protected {
		return result, fmt.Errorf("%w: document module %q cannot be renamed", ErrProtectedModule, module.Name)
	}
	cleanNewName, err := cleanComponentName(newName)
	if err != nil {
		return result, err
	}
	if strings.EqualFold(module.Name, cleanNewName) && module.Name == cleanNewName {
		return result, fmt.Errorf("%w: module %q already exists", ErrModuleAlreadyExists, cleanNewName)
	}
	if err := rejectRenameCollision(cwd, src, module, cleanNewName); err != nil {
		return result, err
	}
	ops, err := moduleRenameOps(cwd, src, module, cleanNewName)
	if err != nil {
		return result, err
	}
	if err := applyRenameOps(ops); err != nil {
		return result, err
	}
	result.Operation = "module.rename"
	result.OldName = module.Name
	result.NewName = cleanNewName
	result.Kind = module.Kind
	result.RequiresPush = true
	for _, op := range ops {
		result.Renamed = append(result.Renamed, ModuleRenameEntry{
			From: filepath.ToSlash(rel(cwd, op.OldPath)),
			To:   filepath.ToSlash(rel(cwd, op.NewPath)),
		})
	}
	return result, nil
}

func ResolveWorkbookRoot(cwd string, src config.SourceConfig) string {
	workbookRoot := strings.TrimSpace(src.Workbook)
	if workbookRoot == "" {
		workbookRoot = config.Default().Src.Workbook
	}
	if !filepath.IsAbs(workbookRoot) {
		workbookRoot = filepath.Join(cwd, filepath.FromSlash(workbookRoot))
	}
	return workbookRoot
}

func resolveSourceModule(cwd string, src config.SourceConfig, name string) (sourceModule, error) {
	cleanName, err := cleanComponentName(name)
	if err != nil {
		return sourceModule{}, err
	}
	var matches []sourceModule
	for _, root := range sourceRoots(cwd, src) {
		found, err := modulesUnderRoot(root)
		if err != nil {
			return sourceModule{}, err
		}
		for _, module := range found {
			if strings.EqualFold(module.Name, cleanName) {
				matches = append(matches, module)
			}
		}
	}
	if len(matches) == 0 {
		return sourceModule{}, fmt.Errorf("%w: module %q was not found", ErrModuleNotFound, cleanName)
	}
	if len(matches) > 1 {
		return sourceModule{}, fmt.Errorf("%w: module %q matched multiple source files", ErrModuleAmbiguous, cleanName)
	}
	return matches[0], nil
}

func sourceRoots(cwd string, src config.SourceConfig) []sourceRoot {
	return []sourceRoot{
		{Path: ResolveModuleRoot(cwd, src), Kind: "standard", Exts: map[string]bool{".bas": true}},
		{Path: ResolveClassRoot(cwd, src), Kind: "class", Exts: map[string]bool{".cls": true}},
		{
			Path:       ResolveFormRoot(cwd, src),
			Kind:       "form",
			Exts:       map[string]bool{".frm": true},
			SkipSubdir: map[string]bool{"code": true, "specs": true},
		},
		{
			Path:      ResolveWorkbookRoot(cwd, src),
			Kind:      "document",
			Exts:      map[string]bool{".bas": true, ".cls": true, ".frm": true},
			Protected: true,
		},
	}
}

func modulesUnderRoot(root sourceRoot) ([]sourceModule, error) {
	var modules []sourceModule
	err := filepath.WalkDir(root.Path, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			if errors.Is(err, os.ErrNotExist) {
				return nil
			}
			return err
		}
		if d.IsDir() {
			if !strings.EqualFold(path, root.Path) && root.SkipSubdir[strings.ToLower(d.Name())] {
				return filepath.SkipDir
			}
			return nil
		}
		ext := strings.ToLower(filepath.Ext(path))
		if !root.Exts[ext] {
			return nil
		}
		modules = append(modules, sourceModule{
			Name:      strings.TrimSuffix(filepath.Base(path), filepath.Ext(path)),
			Kind:      root.Kind,
			Path:      path,
			Protected: root.Protected,
		})
		return nil
	})
	if errors.Is(err, os.ErrNotExist) {
		return modules, nil
	}
	return modules, err
}

func moduleArtifactPaths(cwd string, src config.SourceConfig, module sourceModule) ([]string, error) {
	paths := []string{module.Path}
	if module.Kind != "form" {
		return paths, nil
	}
	formRoot := ResolveFormRoot(cwd, src)
	for _, path := range []string{
		filepath.Join(filepath.Dir(module.Path), module.Name+".frx"),
		filepath.Join(formRoot, "code", module.Name+".bas"),
		filepath.Join(formRoot, "specs", module.Name+".yaml"),
		filepath.Join(formRoot, "specs", module.Name+".yml"),
		filepath.Join(formRoot, "specs", module.Name+".json"),
	} {
		if _, err := os.Stat(path); err == nil {
			paths = append(paths, path)
		} else if !errors.Is(err, os.ErrNotExist) {
			return nil, err
		}
	}
	return paths, nil
}

func rejectRenameCollision(cwd string, src config.SourceConfig, module sourceModule, newName string) error {
	for _, root := range sourceRoots(cwd, src) {
		found, err := modulesUnderRoot(root)
		if err != nil {
			return err
		}
		for _, candidate := range found {
			if samePath(candidate.Path, module.Path) {
				continue
			}
			if strings.EqualFold(candidate.Name, newName) {
				return fmt.Errorf("%w: module %q already exists at %s", ErrModuleAlreadyExists, newName, candidate.Path)
			}
		}
	}
	if module.Kind != "form" {
		return nil
	}
	formRoot := ResolveFormRoot(cwd, src)
	for _, path := range []string{
		filepath.Join(formRoot, "code", newName+".bas"),
		filepath.Join(formRoot, "specs", newName+".yaml"),
		filepath.Join(formRoot, "specs", newName+".yml"),
		filepath.Join(formRoot, "specs", newName+".json"),
	} {
		if _, err := os.Stat(path); err == nil {
			return fmt.Errorf("%w: module %q already exists at %s", ErrModuleAlreadyExists, newName, path)
		} else if !errors.Is(err, os.ErrNotExist) {
			return err
		}
	}
	return nil
}

func moduleRenameOps(cwd string, src config.SourceConfig, module sourceModule, newName string) ([]renameOp, error) {
	oldArtifacts, err := moduleArtifactPaths(cwd, src, module)
	if err != nil {
		return nil, err
	}
	ops := make([]renameOp, 0, len(oldArtifacts))
	for _, oldPath := range oldArtifacts {
		newPath := renamedArtifactPath(cwd, src, module, oldPath, newName)
		op := renameOp{OldPath: oldPath, NewPath: newPath}
		if !samePath(oldPath, newPath) {
			if _, err := os.Stat(newPath); err == nil {
				return nil, fmt.Errorf("%w: module %q already exists at %s", ErrModuleAlreadyExists, newName, newPath)
			} else if !errors.Is(err, os.ErrNotExist) {
				return nil, err
			}
		}
		if shouldRewriteArtifact(module, oldPath) {
			body, err := os.ReadFile(oldPath)
			if err != nil {
				return nil, err
			}
			rewritten, err := rewriteModuleArtifact(module, oldPath, body, newName)
			if err != nil {
				return nil, err
			}
			op.OldContent = body
			op.NewContent = rewritten
			op.UpdateBytes = true
		}
		ops = append(ops, op)
	}
	return ops, nil
}

func renamedArtifactPath(cwd string, src config.SourceConfig, module sourceModule, oldPath string, newName string) string {
	ext := filepath.Ext(oldPath)
	if module.Kind == "form" {
		formRoot := ResolveFormRoot(cwd, src)
		if samePath(filepath.Dir(oldPath), filepath.Join(formRoot, "code")) {
			return filepath.Join(formRoot, "code", newName+ext)
		}
		if samePath(filepath.Dir(oldPath), filepath.Join(formRoot, "specs")) {
			return filepath.Join(formRoot, "specs", newName+ext)
		}
	}
	return filepath.Join(filepath.Dir(oldPath), newName+ext)
}

func shouldRewriteArtifact(module sourceModule, path string) bool {
	ext := strings.ToLower(filepath.Ext(path))
	switch module.Kind {
	case "standard", "class":
		return samePath(path, module.Path)
	case "form":
		return samePath(path, module.Path) || ext == ".yaml" || ext == ".yml" || ext == ".json"
	default:
		return false
	}
}

func rewriteModuleArtifact(module sourceModule, path string, body []byte, newName string) ([]byte, error) {
	ext := strings.ToLower(filepath.Ext(path))
	if module.Kind == "form" && (ext == ".yaml" || ext == ".yml" || ext == ".json") {
		return rewriteFormSpecName(path, body, newName)
	}
	rewritten := rewriteVBNameAttribute(string(body), newName)
	if module.Kind == "form" {
		rewritten = rewriteFormBeginName(rewritten, module.Name, newName)
	}
	return []byte(rewritten), nil
}

var vbNameAttributePattern = regexp.MustCompile(`(?im)^Attribute\s+VB_Name\s*=\s*"[^"]*"`)

func rewriteVBNameAttribute(source string, newName string) string {
	return vbNameAttributePattern.ReplaceAllString(source, fmt.Sprintf("Attribute VB_Name = %q", newName))
}

func rewriteFormBeginName(source string, oldName string, newName string) string {
	pattern := regexp.MustCompile(`(?im)^(\s*Begin\s+\S+\s+)` + regexp.QuoteMeta(oldName) + `(\s*)$`)
	return pattern.ReplaceAllString(source, "${1}"+newName+"${2}")
}

func rewriteFormSpecName(path string, body []byte, newName string) ([]byte, error) {
	var spec map[string]any
	ext := strings.ToLower(filepath.Ext(path))
	switch ext {
	case ".json":
		if err := json.Unmarshal(body, &spec); err != nil {
			return nil, err
		}
	case ".yaml", ".yml":
		if err := yaml.Unmarshal(body, &spec); err != nil {
			return nil, err
		}
	default:
		return body, nil
	}
	form, _ := spec["form"].(map[string]any)
	if form == nil {
		form = map[string]any{}
		spec["form"] = form
	}
	form["name"] = newName
	switch ext {
	case ".json":
		out, err := json.MarshalIndent(spec, "", "  ")
		if err != nil {
			return nil, err
		}
		return append(out, '\n'), nil
	default:
		out, err := yaml.Marshal(spec)
		if err != nil {
			return nil, err
		}
		return out, nil
	}
}

func applyRenameOps(ops []renameOp) error {
	var applied []renameOp
	for _, op := range ops {
		if err := renamePathCaseAware(op.OldPath, op.NewPath); err != nil {
			rollbackRenameOps(applied)
			return err
		}
		applied = append(applied, op)
	}
	for i, op := range applied {
		if !op.UpdateBytes {
			continue
		}
		if err := os.WriteFile(op.NewPath, op.NewContent, 0o644); err != nil {
			rollbackContentAndRename(applied[:i+1])
			return err
		}
	}
	return nil
}

func renamePathCaseAware(oldPath string, newPath string) error {
	if !samePath(oldPath, newPath) {
		return os.Rename(oldPath, newPath)
	}
	if oldPath == newPath {
		return nil
	}
	tmp := filepath.Join(filepath.Dir(oldPath), "."+filepath.Base(oldPath)+".xlflow-rename-tmp")
	if err := os.Rename(oldPath, tmp); err != nil {
		return err
	}
	if err := os.Rename(tmp, newPath); err != nil {
		_ = os.Rename(tmp, oldPath)
		return err
	}
	return nil
}

func rollbackRenameOps(ops []renameOp) {
	for i := len(ops) - 1; i >= 0; i-- {
		_ = os.Rename(ops[i].NewPath, ops[i].OldPath)
	}
}

func rollbackContentAndRename(ops []renameOp) {
	for i := len(ops) - 1; i >= 0; i-- {
		if ops[i].UpdateBytes {
			_ = os.WriteFile(ops[i].NewPath, ops[i].OldContent, 0o644)
		}
		_ = os.Rename(ops[i].NewPath, ops[i].OldPath)
	}
}

func samePath(a string, b string) bool {
	return strings.EqualFold(filepath.Clean(a), filepath.Clean(b))
}
