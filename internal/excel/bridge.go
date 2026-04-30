package excel

import (
	"bytes"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strconv"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/output"
)

type Runner struct {
	RootDir string
}

type WorkbookRef struct {
	Path string `json:"path"`
}

type BackupRef struct {
	Path string `json:"path"`
}

type RunArgument struct {
	Type  string `json:"type"`
	Value string `json:"value"`
}

type RunOptions struct {
	Macro        string
	WorkbookPath string
	Args         []RunArgument
	Save         bool
	SaveAs       string
}

type ScriptResult struct {
	Status      string        `json:"status"`
	Command     string        `json:"command"`
	Error       *output.Error `json:"error"`
	Logs        []string      `json:"logs"`
	Diagnostics any           `json:"diagnostics,omitempty"`
	Workbook    any           `json:"workbook,omitempty"`
	Backup      any           `json:"backup,omitempty"`
	Macro       any           `json:"macro,omitempty"`
	Tests       any           `json:"tests,omitempty"`
}

func (r Runner) Doctor(cfg config.Config) (output.Envelope, int, error) {
	return r.run("doctor", map[string]string{
		"WorkbookPath": workbookPath(r.RootDir, cfg.Excel.Path),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
	})
}

func (r Runner) New(workbook string) (output.Envelope, int, error) {
	return r.run("new", map[string]string{
		"WorkbookPath": workbookPath(r.RootDir, workbook),
	})
}

func (r Runner) Pull(cfg config.Config) (output.Envelope, int, error) {
	return r.run("pull", map[string]string{
		"WorkbookPath": workbookPath(r.RootDir, cfg.Excel.Path),
		"ModulesDir":   filepath.Join(r.RootDir, cfg.Src.Modules),
		"ClassesDir":   filepath.Join(r.RootDir, cfg.Src.Classes),
		"FormsDir":     filepath.Join(r.RootDir, cfg.Src.Forms),
		"WorkbookDir":  filepath.Join(r.RootDir, cfg.Src.Workbook),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
	})
}

func (r Runner) Push(cfg config.Config) (output.Envelope, int, error) {
	return r.run("push", map[string]string{
		"WorkbookPath": workbookPath(r.RootDir, cfg.Excel.Path),
		"ModulesDir":   filepath.Join(r.RootDir, cfg.Src.Modules),
		"ClassesDir":   filepath.Join(r.RootDir, cfg.Src.Classes),
		"FormsDir":     filepath.Join(r.RootDir, cfg.Src.Forms),
		"WorkbookDir":  filepath.Join(r.RootDir, cfg.Src.Workbook),
		"BackupRoot":   filepath.Join(r.RootDir, ".xlflow", "backups"),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
	})
}

func buildRunScriptArgs(root string, cfg config.Config, opts RunOptions) (map[string]string, error) {
	workbook := cfg.Excel.Path
	if opts.WorkbookPath != "" {
		workbook = opts.WorkbookPath
	}
	args := opts.Args
	if args == nil {
		args = []RunArgument{}
	}
	argsJSON, err := json.Marshal(args)
	if err != nil {
		return nil, err
	}
	// Base64-encode the JSON to avoid PowerShell command-line parsing issues
	argsJSON64 := base64.StdEncoding.EncodeToString(argsJSON)
	scriptArgs := map[string]string{
		"WorkbookPath":  workbookPath(root, workbook),
		"MacroName":     opts.Macro,
		"MacroArgsJSON": string(argsJSON64),
		"Visible":       strconv.FormatBool(cfg.Excel.Visible),
		"DisplayAlerts": strconv.FormatBool(cfg.Excel.DisplayAlerts),
		"SaveWorkbook":  strconv.FormatBool(opts.Save),
	}
	if opts.SaveAs != "" {
		scriptArgs["SaveAsPath"] = workbookPath(root, opts.SaveAs)
	}
	return scriptArgs, nil
}

func (r Runner) Run(cfg config.Config, opts RunOptions) (output.Envelope, int, error) {
	scriptArgs, err := buildRunScriptArgs(r.RootDir, cfg, opts)
	if err != nil {
		return output.Failure("run", output.Error{Code: "run_args_invalid", Message: err.Error(), Source: "xlflow"}), output.ExitConfig, nil
	}
	return r.run("run", scriptArgs)
}

func (r Runner) Test(cfg config.Config, filter string) (output.Envelope, int, error) {
	return r.run("test", map[string]string{
		"WorkbookPath": workbookPath(r.RootDir, cfg.Excel.Path),
		"Filter":       filter,
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
	})
}

func (r Runner) run(commandName string, args map[string]string) (output.Envelope, int, error) {
	env := output.New(commandName)
	if runtime.GOOS != "windows" {
		env = output.Failure(commandName, output.Error{Code: "environment", Message: "Excel automation is only supported on Windows in the MVP"})
		return env, output.ExitEnvironment, nil
	}

	script, err := scriptPath(r.RootDir, commandName)
	if err != nil {
		env = output.Failure(commandName, output.Error{Code: "script_not_found", Message: err.Error(), Source: "xlflow"})
		return env, output.ExitEnvironment, nil
	}
	cmdArgs := []string{"-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script}
	for k, v := range args {
		cmdArgs = append(cmdArgs, "-"+k, v)
	}
	cmd := exec.Command("powershell", cmdArgs...)
	var stdout, stderr bytes.Buffer
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr
	if err := cmd.Run(); err != nil {
		message := err.Error()
		if stderr.Len() > 0 {
			message = stderr.String()
		}
		env = output.Failure(commandName, output.Error{Code: "script_failed", Message: message, Source: "powershell"})
		return env, output.ExitEnvironment, nil
	}

	var result ScriptResult
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		env = output.Failure(commandName, output.Error{Code: "invalid_script_json", Message: fmt.Sprintf("failed to parse script JSON: %v", err), Source: "powershell"})
		env.Logs = []string{stdout.String()}
		return env, output.ExitEnvironment, nil
	}
	if result.Status == "" {
		result.Status = output.StatusOK
	}
	env.Status = result.Status
	env.Command = commandName
	env.Error = result.Error
	env.Logs = result.Logs
	if env.Logs == nil {
		env.Logs = []string{}
	}
	env.Diagnostics = result.Diagnostics
	env.Workbook = result.Workbook
	env.Backup = result.Backup
	env.Macro = result.Macro
	env.Tests = result.Tests
	if result.Status == output.StatusFailed {
		return env, exitCodeForScriptResult(result), nil
	}
	return env, output.ExitSuccess, nil
}

func exitCodeForScriptResult(result ScriptResult) int {
	if result.Error == nil {
		return output.ExitEnvironment
	}
	switch result.Error.Code {
	case "macro_failed", "test_failed", "no_tests_found", "test_not_found", "duplicate_test_name":
		return output.ExitValidation
	default:
		return output.ExitEnvironment
	}
}

func workbookPath(root, path string) string {
	if filepath.IsAbs(path) {
		return path
	}
	return filepath.Join(root, path)
}

func scriptPath(root, commandName string) (string, error) {
	name := commandName + ".ps1"
	candidates := []string{
		filepath.Join(root, "scripts", name),
	}
	if _, file, _, ok := runtime.Caller(0); ok {
		candidates = append(candidates, filepath.Join(filepath.Dir(file), "..", "..", "scripts", name))
	}
	if exe, err := os.Executable(); err == nil {
		candidates = append(candidates, filepath.Join(filepath.Dir(exe), "scripts", name))
	}
	for _, candidate := range candidates {
		clean := filepath.Clean(candidate)
		if _, err := os.Stat(clean); err == nil {
			return clean, nil
		}
	}
	return "", fmt.Errorf("script %s was not found", name)
}
