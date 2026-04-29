package project

import (
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
)

type InitResult struct {
	ConfigPath string   `json:"config_path"`
	Workbook   string   `json:"workbook"`
	Created    []string `json:"created"`
}

type WorkbookCreator func(path string) error

func Init(cwd, workbookPath string) (InitResult, error) {
	var result InitResult
	if workbookPath == "" {
		return result, errors.New("workbook path is required")
	}
	srcInfo, err := os.Stat(workbookPath)
	if err != nil {
		return result, fmt.Errorf("cannot read workbook: %w", err)
	}
	if srcInfo.IsDir() {
		return result, fmt.Errorf("workbook path is a directory: %s", workbookPath)
	}
	destPath := filepath.Join(cwd, "build", filepath.Base(workbookPath))
	result, err = createScaffold(cwd, destPath, projectName(workbookPath), func(path string) error {
		return copyFile(workbookPath, path)
	})
	if err != nil {
		return InitResult{}, err
	}
	return result, nil
}

func New(cwd, workbookName string, createWorkbook WorkbookCreator) (InitResult, error) {
	if createWorkbook == nil {
		return InitResult{}, errors.New("workbook creator is required")
	}
	name, err := normalizeWorkbookName(workbookName)
	if err != nil {
		return InitResult{}, err
	}
	destPath := filepath.Join(cwd, "build", name)
	return createScaffold(cwd, destPath, projectName(name), createWorkbook)
}

func createScaffold(cwd, destPath, name string, createWorkbook WorkbookCreator) (InitResult, error) {
	var result InitResult
	configPath := filepath.Join(cwd, config.FileName)
	promptPath := filepath.Join(cwd, "prompts", "agent.md")
	for _, path := range []string{destPath, configPath, promptPath} {
		if _, err := os.Stat(path); err == nil {
			return result, fmt.Errorf("refusing to overwrite existing file: %s", path)
		} else if !errors.Is(err, os.ErrNotExist) {
			return result, err
		}
	}

	dirs := []string{
		filepath.Join(cwd, "src", "modules"),
		filepath.Join(cwd, "src", "classes"),
		filepath.Join(cwd, "src", "forms"),
		filepath.Join(cwd, "src", "workbook"),
		filepath.Join(cwd, "tests"),
		filepath.Join(cwd, "build"),
		filepath.Join(cwd, "prompts"),
		filepath.Join(cwd, ".xlflow"),
	}
	for _, dir := range dirs {
		if err := os.MkdirAll(dir, 0o755); err != nil {
			return result, err
		}
		result.Created = append(result.Created, filepath.ToSlash(rel(cwd, dir)))
	}

	if err := createWorkbook(destPath); err != nil {
		return result, err
	}
	result.Workbook = filepath.ToSlash(rel(cwd, destPath))
	result.Created = append(result.Created, result.Workbook)

	cfg := config.Default()
	cfg.Project.Name = name
	cfg.Excel.Path = result.Workbook
	if err := config.Write(configPath, cfg); err != nil {
		return result, err
	}
	result.ConfigPath = config.FileName
	result.Created = append(result.Created, config.FileName)

	if err := writeExclusive(promptPath, defaultPrompt); err != nil {
		return result, err
	}
	result.Created = append(result.Created, filepath.ToSlash(rel(cwd, promptPath)))
	return result, nil
}

func copyFile(src, dest string) (err error) {
	in, err := os.Open(src)
	if err != nil {
		return err
	}
	defer func() {
		if closeErr := in.Close(); err == nil && closeErr != nil {
			err = closeErr
		}
	}()
	out, err := os.OpenFile(dest, os.O_WRONLY|os.O_CREATE|os.O_EXCL, 0o644)
	if err != nil {
		return err
	}
	defer func() {
		if closeErr := out.Close(); err == nil && closeErr != nil {
			err = closeErr
		}
	}()
	if _, err := io.Copy(out, in); err != nil {
		return err
	}
	return nil
}

func writeExclusive(path, body string) (err error) {
	f, err := os.OpenFile(path, os.O_WRONLY|os.O_CREATE|os.O_EXCL, 0o644)
	if err != nil {
		return err
	}
	defer func() {
		if closeErr := f.Close(); err == nil && closeErr != nil {
			err = closeErr
		}
	}()
	_, err = f.WriteString(body)
	return err
}

func rel(base, path string) string {
	r, err := filepath.Rel(base, path)
	if err != nil {
		return path
	}
	return r
}

func projectName(path string) string {
	name := strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
	name = strings.TrimSpace(name)
	if name == "" {
		return "sample"
	}
	return name
}

func normalizeWorkbookName(name string) (string, error) {
	name = strings.TrimSpace(name)
	if name == "" {
		return "Book.xlsm", nil
	}
	name = filepath.Base(name)
	ext := strings.ToLower(filepath.Ext(name))
	if ext == "" {
		return name + ".xlsm", nil
	}
	if ext != ".xlsm" {
		return "", fmt.Errorf("workbook name must use .xlsm extension: %s", name)
	}
	return name, nil
}

const defaultPrompt = `You are a VBA developer.

Rules:
- Never use Select/Activate
- Always use Option Explicit
- Prefer With blocks
- Avoid global state
`
