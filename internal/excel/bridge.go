package excel

import (
	"bytes"
	"context"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strconv"
	"time"

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
	Trace        bool
	Mode         string
	Timeout      time.Duration
	Keepalive    CommandOptions
}

type CommandOptions struct {
	Keepalive         bool
	KeepaliveInterval time.Duration
	Stderr            io.Writer
}

type ScriptResult struct {
	Status        string        `json:"status"`
	Command       string        `json:"command"`
	Error         *output.Error `json:"error"`
	Logs          []string      `json:"logs"`
	Diagnostics   any           `json:"diagnostics,omitempty"`
	Workbook      any           `json:"workbook,omitempty"`
	Backup        any           `json:"backup,omitempty"`
	Source        any           `json:"source,omitempty"`
	Macro         any           `json:"macro,omitempty"`
	Macros        any           `json:"macros,omitempty"`
	Tests         any           `json:"tests,omitempty"`
	Trace         any           `json:"trace,omitempty"`
	GUIBoundaries any           `json:"gui_boundaries,omitempty"`
	UI            any           `json:"ui,omitempty"`
}

type UIButtonAddOptions struct {
	Sheet       string
	Cell        string
	Text        string
	Macro       string
	ID          string
	Width       int
	Height      int
	CreateSheet bool
	VerifyMacro bool
}

type UIButtonListOptions struct {
	Sheet string
}

type UIButtonRemoveOptions struct {
	Sheet string
	ID    string
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

func (r Runner) Push(cfg config.Config, opts ...CommandOptions) (output.Envelope, int, error) {
	return r.run("push", map[string]string{
		"WorkbookPath": workbookPath(r.RootDir, cfg.Excel.Path),
		"ModulesDir":   filepath.Join(r.RootDir, cfg.Src.Modules),
		"ClassesDir":   filepath.Join(r.RootDir, cfg.Src.Classes),
		"FormsDir":     filepath.Join(r.RootDir, cfg.Src.Forms),
		"WorkbookDir":  filepath.Join(r.RootDir, cfg.Src.Workbook),
		"BackupRoot":   filepath.Join(r.RootDir, ".xlflow", "backups"),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
	}, opts...)
}

func (r Runner) TraceInject(cfg config.Config, workbook string) (output.Envelope, int, error) {
	return r.run("trace", buildTraceInjectScriptArgs(r.RootDir, cfg, workbook))
}

func buildTraceInjectScriptArgs(root string, cfg config.Config, workbook string) map[string]string {
	if workbook == "" {
		workbook = cfg.Excel.Path
	}
	args := map[string]string{
		"Action":       "inject",
		"WorkbookPath": workbookPath(root, workbook),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
	}
	if workbook == cfg.Excel.Path {
		args["ModulesDir"] = filepath.Join(root, cfg.Src.Modules)
	}
	return args
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
		"TraceEnabled":  strconv.FormatBool(opts.Trace),
	}
	if opts.Mode == "interactive" {
		scriptArgs["Visible"] = "true"
		scriptArgs["DisplayAlerts"] = "true"
	}
	if opts.Timeout > 0 {
		scriptArgs["TimeoutSeconds"] = strconv.Itoa(int(opts.Timeout.Seconds()))
	}
	if opts.SaveAs != "" {
		scriptArgs["SaveAsPath"] = workbookPath(root, opts.SaveAs)
	}
	if opts.Trace {
		scriptArgs["TraceFile"] = filepath.Join(os.TempDir(), "xlflow", fmt.Sprintf("trace-%d.log", time.Now().UnixNano()))
	}
	return scriptArgs, nil
}

func (r Runner) Run(cfg config.Config, opts RunOptions) (output.Envelope, int, error) {
	scriptArgs, err := buildRunScriptArgs(r.RootDir, cfg, opts)
	if err != nil {
		return output.Failure("run", output.Error{Code: "run_args_invalid", Message: err.Error(), Source: "xlflow"}), output.ExitConfig, nil
	}
	return r.runWithOptions("run", scriptArgs, commandRunOptions{
		Timeout:   opts.Timeout,
		Keepalive: opts.Keepalive,
	})
}

func (r Runner) Attach(cfg config.Config, active bool) (output.Envelope, int, error) {
	if !active {
		return output.Failure("attach", output.Error{Code: "attach_args_invalid", Message: "--active is required for attach in this version", Source: "xlflow"}), output.ExitConfig, nil
	}
	return r.run("attach", map[string]string{
		"WorkbookPath": workbookPath(r.RootDir, cfg.Excel.Path),
		"Active":       strconv.FormatBool(active),
	})
}

func (r Runner) Test(cfg config.Config, filter string) (output.Envelope, int, error) {
	return r.run("test", map[string]string{
		"WorkbookPath": workbookPath(r.RootDir, cfg.Excel.Path),
		"Filter":       filter,
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
	})
}

func (r Runner) Macros(cfg config.Config) (output.Envelope, int, error) {
	return r.run("macros", map[string]string{
		"WorkbookPath": workbookPath(r.RootDir, cfg.Excel.Path),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
	})
}

func (r Runner) UIButtonAdd(cfg config.Config, opts UIButtonAddOptions) (output.Envelope, int, error) {
	return r.run("ui", buildUIButtonAddScriptArgs(r.RootDir, cfg, opts))
}

func buildUIButtonAddScriptArgs(root string, cfg config.Config, opts UIButtonAddOptions) map[string]string {
	return map[string]string{
		"Action":       "add",
		"WorkbookPath": workbookPath(root, cfg.Excel.Path),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
		"Sheet":        opts.Sheet,
		"Cell":         opts.Cell,
		"Text":         opts.Text,
		"Macro":        opts.Macro,
		"Id":           opts.ID,
		"Width":        strconv.Itoa(opts.Width),
		"Height":       strconv.Itoa(opts.Height),
		"CreateSheet":  strconv.FormatBool(opts.CreateSheet),
		"VerifyMacro":  strconv.FormatBool(opts.VerifyMacro),
	}
}

func (r Runner) UIButtonList(cfg config.Config, opts UIButtonListOptions) (output.Envelope, int, error) {
	return r.run("ui", buildUIButtonListScriptArgs(r.RootDir, cfg, opts))
}

func buildUIButtonListScriptArgs(root string, cfg config.Config, opts UIButtonListOptions) map[string]string {
	return map[string]string{
		"Action":       "list",
		"WorkbookPath": workbookPath(root, cfg.Excel.Path),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
		"Sheet":        opts.Sheet,
	}
}

func (r Runner) UIButtonRemove(cfg config.Config, opts UIButtonRemoveOptions) (output.Envelope, int, error) {
	return r.run("ui", buildUIButtonRemoveScriptArgs(r.RootDir, cfg, opts))
}

func buildUIButtonRemoveScriptArgs(root string, cfg config.Config, opts UIButtonRemoveOptions) map[string]string {
	return map[string]string{
		"Action":       "remove",
		"WorkbookPath": workbookPath(root, cfg.Excel.Path),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
		"Sheet":        opts.Sheet,
		"Id":           opts.ID,
	}
}

func (r Runner) run(commandName string, args map[string]string, opts ...CommandOptions) (output.Envelope, int, error) {
	runOpts := commandRunOptions{}
	if len(opts) > 0 {
		runOpts.Keepalive = opts[0]
	}
	return r.runWithOptions(commandName, args, runOpts)
}

type commandRunOptions struct {
	Timeout   time.Duration
	Keepalive CommandOptions
}

func (r Runner) runWithOptions(commandName string, args map[string]string, opts commandRunOptions) (output.Envelope, int, error) {
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
	var ctx context.Context
	var cancel context.CancelFunc
	if opts.Timeout > 0 {
		ctx, cancel = context.WithTimeout(context.Background(), opts.Timeout)
	} else {
		ctx, cancel = context.WithCancel(context.Background())
	}
	defer cancel()
	cmd := exec.CommandContext(ctx, "powershell", cmdArgs...)
	var stdout, stderr bytes.Buffer
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr
	stopKeepalive := startKeepalive(commandName, opts.Keepalive)
	err = cmd.Start()
	if err == nil {
		err = cmd.Wait()
	}
	stopKeepalive()
	if err != nil {
		if ctx.Err() == context.DeadlineExceeded && commandName == "run" {
			env = output.Failure(commandName, output.Error{
				Code:    "macro_timeout",
				Message: fmt.Sprintf("Macro did not complete within %s. Possible causes: a file picker, MsgBox, UserForm, or long-running loop is still waiting.", opts.Timeout.String()),
				Source:  "xlflow",
				Phase:   "invoke_macro",
			})
			env.Logs = []string{
				"Excel automation timed out while running the macro.",
				"Use xlflow run --interactive when a human can complete dialogs, or refactor GUI calls behind a headless entrypoint.",
			}
			writeDoneMarker(commandName, env, opts.Keepalive)
			return env, output.ExitValidation, nil
		}
		message := err.Error()
		if stderr.Len() > 0 {
			message = stderr.String()
		}
		env = output.Failure(commandName, output.Error{Code: "script_failed", Message: message, Source: "powershell"})
		writeDoneMarker(commandName, env, opts.Keepalive)
		return env, output.ExitEnvironment, nil
	}

	var result ScriptResult
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		env = output.Failure(commandName, output.Error{Code: "invalid_script_json", Message: fmt.Sprintf("failed to parse script JSON: %v", err), Source: "powershell"})
		env.Logs = []string{stdout.String()}
		writeDoneMarker(commandName, env, opts.Keepalive)
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
	env.Source = result.Source
	env.Macro = result.Macro
	env.Macros = result.Macros
	env.Tests = result.Tests
	env.Trace = result.Trace
	env.GUIBoundaries = result.GUIBoundaries
	env.UI = result.UI
	writeDoneMarker(commandName, env, opts.Keepalive)
	if result.Status == output.StatusFailed {
		return env, exitCodeForScriptResult(result), nil
	}
	return env, output.ExitSuccess, nil
}

func startKeepalive(commandName string, opts CommandOptions) func() {
	if !opts.Keepalive {
		return func() {}
	}
	w := opts.Stderr
	if w == nil {
		w = os.Stderr
	}
	interval := opts.KeepaliveInterval
	if interval <= 0 {
		interval = 5 * time.Second
	}
	started := time.Now()
	done := make(chan struct{})
	stopped := make(chan struct{})
	_, _ = fmt.Fprintf(w, "xlflow: %s still running... elapsed=0s\n", commandName)
	go func() {
		defer close(stopped)
		ticker := time.NewTicker(interval)
		defer ticker.Stop()
		for {
			select {
			case <-done:
				return
			case <-ticker.C:
				elapsed := time.Since(started).Truncate(time.Second)
				_, _ = fmt.Fprintf(w, "xlflow: %s still running... elapsed=%s\n", commandName, elapsed)
			}
		}
	}()
	return func() {
		close(done)
		<-stopped
	}
}

func writeDoneMarker(commandName string, env output.Envelope, opts CommandOptions) {
	if !opts.Keepalive {
		return
	}
	w := opts.Stderr
	if w == nil {
		w = os.Stderr
	}
	status := "success"
	if env.Status == output.StatusFailed {
		status = "failed"
	}
	_, _ = fmt.Fprintf(w, "XLFLOW_DONE status=%s command=%s", status, commandName)
	if env.Error != nil && env.Error.Code != "" {
		_, _ = fmt.Fprintf(w, " code=%s", env.Error.Code)
	}
	_, _ = fmt.Fprintln(w)
}

func exitCodeForScriptResult(result ScriptResult) int {
	if result.Error == nil {
		return output.ExitEnvironment
	}
	switch result.Error.Code {
	case "macro_failed", "macro_not_found", "macro_timeout", "trace_not_injected", "test_failed", "no_tests_found", "test_not_found", "duplicate_test_name", "active_workbook_mismatch", "sheet_not_found", "button_not_found", "ui_button_args_invalid":
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
