package excel

import (
	"bytes"
	"context"
	"encoding/base64"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strconv"
	"strings"
	"time"
	"unicode"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/excel/forms"
	bundledscripts "github.com/harumiWeb/xlflow/internal/excel/scripts"
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

type FileDialogResponse struct {
	Kind      string   `json:"kind"`
	DialogID  string   `json:"dialog_id"`
	Values    []string `json:"values,omitempty"`
	Cancelled bool     `json:"cancelled,omitempty"`
}

type UIResponses struct {
	MsgBox     map[string]string
	Input      map[string]string
	FileDialog []FileDialogResponse
}

type UIStreamOptions struct {
	Enabled     bool
	RedactInput bool
}

type DebugStreamOptions struct {
	Enabled  bool
	PipeName string
}

type RunOptions struct {
	Macro               string
	WorkbookPath        string
	Args                []RunArgument
	UIResponses         UIResponses
	UIStream            UIStreamOptions
	DebugStream         DebugStreamOptions
	Save                bool
	SaveAs              string
	Trace               bool
	Mode                string
	RuntimeMode         string
	RuntimeSource       string
	Direct              bool
	Fast                bool
	Diagnostic          bool
	SuppressModalErrors bool
	Session             bool
	Timeout             time.Duration
	Keepalive           CommandOptions
	TraceDir            string
}

type PushOptions struct {
	BackupMode  string
	Fast        bool
	ChangedOnly bool
	Session     bool
	NoSave      bool
	Keepalive   CommandOptions
	// SourceRoot, if set, overrides the configured source directories and
	// pushes only modules under this root. Other component dirs are cleared.
	SourceRoot string
}

type ExportImageOptions struct {
	WorkbookPath string
	Sheet        string
	Range        string
	OutPath      string
	OutputDir    string
	Name         string
	Format       string
	Overwrite    bool
	Session      bool
	Keepalive    CommandOptions
}

type FormExportImageOptions struct {
	Name        string
	OutPath     string
	Initializer string
	Overwrite   bool
	Session     bool
	Keepalive   CommandOptions
}

type FormWriteOptions struct {
	Action    string
	SpecPath  string
	Spec      forms.FormSpec
	Overwrite bool
	Session   bool
	NoSave    bool
	Keepalive CommandOptions
}

type EditEventMode string

const (
	EditEventKeep EditEventMode = "keep"
	EditEventOn   EditEventMode = "on"
	EditEventOff  EditEventMode = "off"
)

type EditCellOptions struct {
	WorkbookPath string
	Sheet        string
	Cell         string
	Value        *string
	Formula      *string
	Fill         string
	Events       EditEventMode
	Session      bool
	Keepalive    CommandOptions
}

type EditRangeOptions struct {
	WorkbookPath string
	Sheet        string
	Range        string
	Fill         string
	Clear        string
	Session      bool
	Keepalive    CommandOptions
}

type EditRowsOptions struct {
	WorkbookPath string
	Sheet        string
	Rows         string
	Height       float64
	Session      bool
	Keepalive    CommandOptions
}

type EditColumnsOptions struct {
	WorkbookPath string
	Sheet        string
	Columns      string
	Width        float64
	Session      bool
	Keepalive    CommandOptions
}

type SessionCommandOptions struct {
	Session   bool
	Keepalive CommandOptions
}

type MacrosOptions struct {
	Session      bool
	Keepalive    CommandOptions
	Entry        string
	RunnableOnly bool
}

type TestOptions struct {
	Session       bool
	Keepalive     CommandOptions
	RuntimeMode   string
	RuntimeSource string
	UIResponses   UIResponses
	UIStream      UIStreamOptions
	DebugStream   DebugStreamOptions
	ModuleFilter  string
	TagFilter     string
}

const (
	RuntimeModeInteractive = "interactive"
	RuntimeModeHeadless    = "headless"
	RuntimeModeCI          = "ci"
	RuntimeModeAgent       = "agent"
	RuntimeModeTest        = "test"

	RuntimeSourceCommand     = "command"
	RuntimeSourceEnvironment = "environment"
	RuntimeSourceDefault     = "default"
)

type RuntimeInfo struct {
	Mode   string
	Source string
}

func ResolveRunRuntimeInfo(mode string) RuntimeInfo {
	trimmed := strings.ToLower(strings.TrimSpace(mode))
	if trimmed == RuntimeModeHeadless || trimmed == RuntimeModeInteractive {
		return RuntimeInfo{Mode: trimmed, Source: RuntimeSourceCommand}
	}
	if envMode, ok := resolveRuntimeModeEnv(); ok {
		return RuntimeInfo{Mode: envMode, Source: RuntimeSourceEnvironment}
	}
	return RuntimeInfo{Mode: RuntimeModeInteractive, Source: RuntimeSourceDefault}
}

func ResolveTestRuntimeInfo() RuntimeInfo {
	return RuntimeInfo{Mode: RuntimeModeTest, Source: RuntimeSourceCommand}
}

func resolveRuntimeModeEnv() (string, bool) {
	raw, ok := os.LookupEnv("XLFLOW_MODE")
	if !ok {
		return "", false
	}
	switch strings.ToLower(strings.TrimSpace(raw)) {
	case RuntimeModeInteractive, RuntimeModeHeadless, RuntimeModeCI, RuntimeModeAgent, RuntimeModeTest:
		return strings.ToLower(strings.TrimSpace(raw)), true
	default:
		return "", false
	}
}

type InspectFormOptions struct {
	Name           string
	Basis          string
	Initializer    string
	StrictDesigner bool
	Session        bool
	Keepalive      CommandOptions
}

type InspectOptions struct {
	Target       string
	Sheet        string
	Address      string
	Limits       map[string]int
	IncludeStyle bool
	Session      bool
	Keepalive    CommandOptions
}

type SessionOptions struct {
	Action string
}

type CommandOptions struct {
	Stderr   io.Writer
	Progress bool
}

type ScriptLogs []string

func (l *ScriptLogs) UnmarshalJSON(data []byte) error {
	trimmed := bytes.TrimSpace(data)
	if bytes.Equal(trimmed, []byte("null")) {
		*l = nil
		return nil
	}

	var logs []string
	if err := json.Unmarshal(trimmed, &logs); err == nil {
		*l = ScriptLogs(logs)
		return nil
	}

	var single string
	if err := json.Unmarshal(trimmed, &single); err == nil {
		*l = ScriptLogs{single}
		return nil
	}

	return fmt.Errorf("expected logs to be a string array or single string, got %s", string(trimmed))
}

type ScriptResult struct {
	Status        string        `json:"status"`
	Command       string        `json:"command"`
	Error         *output.Error `json:"error"`
	Logs          ScriptLogs    `json:"logs"`
	Diagnostics   any           `json:"diagnostics,omitempty"`
	Workbook      any           `json:"workbook,omitempty"`
	Backup        any           `json:"backup,omitempty"`
	Source        any           `json:"source,omitempty"`
	Bridge        any           `json:"bridge,omitempty"`
	Macro         any           `json:"macro,omitempty"`
	Macros        any           `json:"macros,omitempty"`
	Forms         any           `json:"forms,omitempty"`
	Tests         any           `json:"tests,omitempty"`
	Trace         any           `json:"trace,omitempty"`
	Runtime       any           `json:"runtime,omitempty"`
	GUIBoundaries any           `json:"gui_boundaries,omitempty"`
	UI            any           `json:"ui,omitempty"`
	Session       any           `json:"session,omitempty"`
	Runner        any           `json:"runner,omitempty"`
	Analysis      any           `json:"analysis,omitempty"`
	Check         any           `json:"check,omitempty"`
	RunDiagnostic any           `json:"run_diagnostic,omitempty"`
	Target        any           `json:"target,omitempty"`
	Output        any           `json:"output,omitempty"`
	Debug         any           `json:"debug,omitempty"`
	Spec          any           `json:"spec,omitempty"`
	Edit          any           `json:"edit,omitempty"`
	Warnings      any           `json:"warnings,omitempty"`
	Hints         any           `json:"hints,omitempty"`
	Inspect       any           `json:"inspect,omitempty"`
	DefaultEntry  string        `json:"default_entry,omitempty"`
	Suggestions   any           `json:"suggestions,omitempty"`
	Process       any           `json:"process,omitempty"`
}

type ProcessListOptions struct {
	Action string
}

type ProcessCleanupOptions struct {
	Action string
	PID    int
	Auto   bool
	All    bool
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
	Session     bool
}

type UIButtonListOptions struct {
	Sheet   string
	Session bool
}

type UIButtonRemoveOptions struct {
	Sheet   string
	ID      string
	Session bool
}

func (r Runner) Doctor(cfg config.Config, opts ...CommandOptions) (output.Envelope, int, error) {
	return r.run("doctor", map[string]string{
		"WorkbookPath": workbookPath(r.RootDir, cfg.Excel.Path),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
	}, opts...)
}

func (r Runner) New(workbook string, opts ...CommandOptions) (output.Envelope, int, error) {
	return r.run("new", map[string]string{
		"WorkbookPath": workbookPath(r.RootDir, workbook),
	}, opts...)
}

func (r Runner) Pull(cfg config.Config, opts ...CommandOptions) (output.Envelope, int, error) {
	cmdOpts := SessionCommandOptions{}
	if len(opts) > 0 {
		cmdOpts.Keepalive = opts[0]
	}
	return r.PullWithOptions(cfg, cmdOpts)
}

func (r Runner) PullWithOptions(cfg config.Config, opts SessionCommandOptions) (output.Envelope, int, error) {
	return r.run("pull", buildPullScriptArgs(r.RootDir, cfg, opts), opts.Keepalive)
}

func (r Runner) Push(cfg config.Config, opts ...CommandOptions) (output.Envelope, int, error) {
	pushOpts := PushOptions{BackupMode: "always"}
	if len(opts) > 0 {
		pushOpts.Keepalive = opts[0]
	}
	return r.PushWithOptions(cfg, pushOpts)
}

func (r Runner) PushWithOptions(cfg config.Config, opts PushOptions) (output.Envelope, int, error) {
	backupMode := opts.BackupMode
	if backupMode == "" {
		backupMode = "always"
	}
	changedOnly := opts.ChangedOnly
	if opts.Fast {
		backupMode = "never"
		changedOnly = true
	}
	modulesDir := filepath.Join(r.RootDir, cfg.Src.Modules)
	classesDir := filepath.Join(r.RootDir, cfg.Src.Classes)
	formsDir := filepath.Join(r.RootDir, cfg.Src.Forms)
	workbookDir := filepath.Join(r.RootDir, cfg.Src.Workbook)
	if opts.SourceRoot != "" {
		modulesDir = opts.SourceRoot
		classesDir = ""
		formsDir = ""
		workbookDir = ""
	}
	return r.run("push", map[string]string{
		"WorkbookPath":            workbookPath(r.RootDir, cfg.Excel.Path),
		"ModulesDir":              modulesDir,
		"ClassesDir":              classesDir,
		"FormsDir":                formsDir,
		"WorkbookDir":             workbookDir,
		"CodeSource":              cfg.UserForm.CodeSource,
		"BackupRoot":              filepath.Join(r.RootDir, ".xlflow", "backups"),
		"Folders":                 strconv.FormatBool(cfg.VBA.Folders),
		"FolderAnnotation":        cfg.VBA.FolderAnnotation,
		"DefaultComponentFolders": strconv.FormatBool(cfg.VBA.DefaultComponentFolders),
		"StatePath":               filepath.Join(r.RootDir, ".xlflow", "state", "push.json"),
		"Visible":                 strconv.FormatBool(cfg.Excel.Visible),
		"BackupMode":              backupMode,
		"ChangedOnly":             strconv.FormatBool(changedOnly),
		"UseSession":              strconv.FormatBool(opts.Session),
		"NoSave":                  strconv.FormatBool(opts.NoSave),
		"MetadataPath":            filepath.Join(r.RootDir, ".xlflow", "session.json"),
	}, opts.Keepalive)
}

func buildPullScriptArgs(root string, cfg config.Config, opts SessionCommandOptions) map[string]string {
	return map[string]string{
		"WorkbookPath":            workbookPath(root, cfg.Excel.Path),
		"ModulesDir":              filepath.Join(root, cfg.Src.Modules),
		"ClassesDir":              filepath.Join(root, cfg.Src.Classes),
		"FormsDir":                filepath.Join(root, cfg.Src.Forms),
		"WorkbookDir":             filepath.Join(root, cfg.Src.Workbook),
		"CodeSource":              cfg.UserForm.CodeSource,
		"Folders":                 strconv.FormatBool(cfg.VBA.Folders),
		"FolderAnnotation":        cfg.VBA.FolderAnnotation,
		"DefaultComponentFolders": strconv.FormatBool(cfg.VBA.DefaultComponentFolders),
		"Visible":                 strconv.FormatBool(cfg.Excel.Visible),
		"UseSession":              strconv.FormatBool(opts.Session),
		"MetadataPath":            filepath.Join(root, ".xlflow", "session.json"),
	}
}

func (r Runner) TraceInject(cfg config.Config, workbook string, opts ...CommandOptions) (output.Envelope, int, error) {
	return r.Trace(cfg, TraceOptions{Action: "enable", Workbook: workbook}, opts...)
}

type TraceOptions struct {
	Action   string
	Workbook string
	Force    bool
	Session  bool
}

func (r Runner) Trace(cfg config.Config, traceOpts TraceOptions, opts ...CommandOptions) (output.Envelope, int, error) {
	return r.run("trace", buildTraceScriptArgs(r.RootDir, cfg, traceOpts), opts...)
}

func buildTraceInjectScriptArgs(root string, cfg config.Config, workbook string) map[string]string {
	return buildTraceScriptArgs(root, cfg, TraceOptions{Action: "enable", Workbook: workbook})
}

func buildTraceScriptArgs(root string, cfg config.Config, traceOpts TraceOptions) map[string]string {
	action := traceOpts.Action
	if action == "" || action == "inject" {
		action = "enable"
	}
	workbook := traceOpts.Workbook
	if workbook == "" && action != "clean" {
		workbook = cfg.Excel.Path
	}
	args := map[string]string{
		"Action":       action,
		"WorkbookPath": workbookPath(root, workbook),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
		"Force":        strconv.FormatBool(traceOpts.Force),
		"UseSession":   strconv.FormatBool(traceOpts.Session),
		"MetadataPath": filepath.Join(root, ".xlflow", "session.json"),
		"TraceDir":     filepath.Join(root, ".xlflow", "traces"),
	}
	if workbook == cfg.Excel.Path && action != "clean" {
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
		"WorkbookPath":        workbookPath(root, workbook),
		"MacroName":           opts.Macro,
		"MacroArgsJSON":       string(argsJSON64),
		"Visible":             strconv.FormatBool(cfg.Excel.Visible),
		"DisplayAlerts":       strconv.FormatBool(cfg.Excel.DisplayAlerts),
		"SaveWorkbook":        strconv.FormatBool(opts.Save),
		"TraceEnabled":        strconv.FormatBool(opts.Trace),
		"Direct":              strconv.FormatBool(opts.Direct || (opts.Fast && len(args) == 0 && !opts.Trace && !opts.Diagnostic)),
		"Diagnostic":          strconv.FormatBool(opts.Diagnostic),
		"SuppressModalErrors": strconv.FormatBool(opts.SuppressModalErrors),
		"UseSession":          strconv.FormatBool(opts.Session),
		"MetadataPath":        filepath.Join(root, ".xlflow", "session.json"),
	}
	if strings.TrimSpace(opts.RuntimeMode) != "" {
		scriptArgs["RuntimeMode"] = strings.TrimSpace(opts.RuntimeMode)
	}
	if strings.TrimSpace(opts.RuntimeSource) != "" {
		scriptArgs["RuntimeSource"] = strings.TrimSpace(opts.RuntimeSource)
	}
	if len(opts.UIResponses.MsgBox) > 0 {
		msgBoxJSON, err := json.Marshal(opts.UIResponses.MsgBox)
		if err != nil {
			return nil, err
		}
		scriptArgs["MsgBoxResponsesJSON"] = base64.StdEncoding.EncodeToString(msgBoxJSON)
	}
	if len(opts.UIResponses.Input) > 0 {
		inputJSON, err := json.Marshal(opts.UIResponses.Input)
		if err != nil {
			return nil, err
		}
		scriptArgs["InputResponsesJSON"] = base64.StdEncoding.EncodeToString(inputJSON)
	}
	if len(opts.UIResponses.FileDialog) > 0 {
		fileDialogJSON, err := json.Marshal(opts.UIResponses.FileDialog)
		if err != nil {
			return nil, err
		}
		scriptArgs["FileDialogResponsesJSON"] = base64.StdEncoding.EncodeToString(fileDialogJSON)
	}
	if opts.UIStream.Enabled {
		scriptArgs["UIStreamEnabled"] = strconv.FormatBool(true)
		scriptArgs["UIStreamRedactInput"] = strconv.FormatBool(opts.UIStream.RedactInput)
	}
	if opts.DebugStream.Enabled {
		scriptArgs["DebugStreamEnabled"] = strconv.FormatBool(true)
		if strings.TrimSpace(opts.DebugStream.PipeName) != "" {
			scriptArgs["DebugStreamPipeName"] = strings.TrimSpace(opts.DebugStream.PipeName)
		}
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
		traceDir := opts.TraceDir
		if traceDir == "" {
			traceDir = filepath.Join(root, ".xlflow", "traces")
		}
		scriptArgs["TraceFile"] = filepath.Join(traceDir, fmt.Sprintf("trace-%d.log", time.Now().UnixNano()))
	}
	return scriptArgs, nil
}

func (r Runner) Session(cfg config.Config, action string, opts ...CommandOptions) (output.Envelope, int, error) {
	cmdOpts := CommandOptions{}
	if len(opts) > 0 {
		cmdOpts = opts[0]
	}
	return r.run("session", map[string]string{
		"Action":       action,
		"WorkbookPath": workbookPath(r.RootDir, cfg.Excel.Path),
		"MetadataPath": filepath.Join(r.RootDir, ".xlflow", "session.json"),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
	}, cmdOpts)
}

func (r Runner) SaveSession(cfg config.Config, opts ...SessionCommandOptions) (output.Envelope, int, error) {
	cmdOpts := SessionCommandOptions{}
	if len(opts) > 0 {
		cmdOpts = opts[0]
	}
	return r.run("session", map[string]string{
		"Action":       "save",
		"WorkbookPath": workbookPath(r.RootDir, cfg.Excel.Path),
		"MetadataPath": filepath.Join(r.RootDir, ".xlflow", "session.json"),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
		"UseSession":   strconv.FormatBool(cmdOpts.Session),
	}, cmdOpts.Keepalive)
}

func (r Runner) RunnerModule(cfg config.Config, action string, opts ...CommandOptions) (output.Envelope, int, error) {
	cmdOpts := CommandOptions{}
	if len(opts) > 0 {
		cmdOpts = opts[0]
	}
	return r.run("runner", map[string]string{
		"Action":       action,
		"WorkbookPath": workbookPath(r.RootDir, cfg.Excel.Path),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
	}, cmdOpts)
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

func (r Runner) Attach(cfg config.Config, active bool, opts ...CommandOptions) (output.Envelope, int, error) {
	if !active {
		return output.Failure("attach", output.Error{Code: "attach_args_invalid", Message: "--active is required for attach in this version", Source: "xlflow"}), output.ExitConfig, nil
	}
	return r.run("attach", map[string]string{
		"WorkbookPath": workbookPath(r.RootDir, cfg.Excel.Path),
		"Active":       strconv.FormatBool(active),
	}, opts...)
}

func (r Runner) Test(cfg config.Config, filter string, opts ...CommandOptions) (output.Envelope, int, error) {
	cmdOpts := TestOptions{}
	if len(opts) > 0 {
		cmdOpts.Keepalive = opts[0]
	}
	return r.TestWithOptions(cfg, filter, cmdOpts)
}

func (r Runner) TestWithOptions(cfg config.Config, filter string, opts TestOptions) (output.Envelope, int, error) {
	return r.run("test", buildTestScriptArgs(r.RootDir, cfg, filter, opts), opts.Keepalive)
}

func buildTestScriptArgs(root string, cfg config.Config, filter string, opts TestOptions) map[string]string {
	args := map[string]string{
		"WorkbookPath": workbookPath(root, cfg.Excel.Path),
		"Filter":       filter,
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
		"UseSession":   strconv.FormatBool(opts.Session),
		"MetadataPath": filepath.Join(root, ".xlflow", "session.json"),
	}
	if strings.TrimSpace(opts.RuntimeMode) != "" {
		args["RuntimeMode"] = strings.TrimSpace(opts.RuntimeMode)
	}
	if strings.TrimSpace(opts.RuntimeSource) != "" {
		args["RuntimeSource"] = strings.TrimSpace(opts.RuntimeSource)
	}
	if len(opts.UIResponses.MsgBox) > 0 {
		if msgBoxJSON, err := json.Marshal(opts.UIResponses.MsgBox); err == nil {
			args["MsgBoxResponsesJSON"] = base64.StdEncoding.EncodeToString(msgBoxJSON)
		}
	}
	if len(opts.UIResponses.Input) > 0 {
		if inputJSON, err := json.Marshal(opts.UIResponses.Input); err == nil {
			args["InputResponsesJSON"] = base64.StdEncoding.EncodeToString(inputJSON)
		}
	}
	if len(opts.UIResponses.FileDialog) > 0 {
		if fileDialogJSON, err := json.Marshal(opts.UIResponses.FileDialog); err == nil {
			args["FileDialogResponsesJSON"] = base64.StdEncoding.EncodeToString(fileDialogJSON)
		}
	}
	if opts.UIStream.Enabled {
		args["UIStreamEnabled"] = strconv.FormatBool(true)
		args["UIStreamRedactInput"] = strconv.FormatBool(opts.UIStream.RedactInput)
	}
	if opts.DebugStream.Enabled {
		args["DebugStreamEnabled"] = strconv.FormatBool(true)
		if strings.TrimSpace(opts.DebugStream.PipeName) != "" {
			args["DebugStreamPipeName"] = strings.TrimSpace(opts.DebugStream.PipeName)
		}
	}
	if strings.TrimSpace(opts.ModuleFilter) != "" {
		args["ModuleFilter"] = strings.TrimSpace(opts.ModuleFilter)
	}
	if strings.TrimSpace(opts.TagFilter) != "" {
		args["TagFilter"] = strings.TrimSpace(opts.TagFilter)
	}
	return args
}

func (r Runner) Macros(cfg config.Config, opts ...CommandOptions) (output.Envelope, int, error) {
	cmdOpts := CommandOptions{}
	if len(opts) > 0 {
		cmdOpts = opts[0]
	}
	return r.MacrosWithOptions(cfg, MacrosOptions{Keepalive: cmdOpts})
}

func (r Runner) MacrosWithOptions(cfg config.Config, opts MacrosOptions) (output.Envelope, int, error) {
	return r.run("macros", map[string]string{
		"WorkbookPath": workbookPath(r.RootDir, cfg.Excel.Path),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
		"UseSession":   strconv.FormatBool(opts.Session),
		"MetadataPath": filepath.Join(r.RootDir, ".xlflow", "session.json"),
		"Entry":        opts.Entry,
		"RunnableOnly": strconv.FormatBool(opts.RunnableOnly),
	}, opts.Keepalive)
}

func (r Runner) ListForms(cfg config.Config, opts SessionCommandOptions) (output.Envelope, int, error) {
	return r.run("list", buildListFormsScriptArgs(r.RootDir, cfg, opts), opts.Keepalive)
}

func (r Runner) InspectForm(cfg config.Config, opts InspectFormOptions) (output.Envelope, int, error) {
	env, code, err := r.run("inspect-form", buildInspectFormScriptArgs(r.RootDir, cfg, opts), opts.Keepalive)
	env.Command = "inspect"
	return env, code, err
}

func (r Runner) Inspect(cfg config.Config, opts InspectOptions) (output.Envelope, int, error) {
	env, code, err := r.run("inspect", buildInspectScriptArgs(r.RootDir, cfg, opts), opts.Keepalive)
	env.Command = "inspect"
	return env, code, err
}

func (r Runner) FormExportImage(cfg config.Config, opts FormExportImageOptions) (output.Envelope, int, error) {
	scriptArgs, err := buildFormExportImageScriptArgs(r.RootDir, cfg, opts)
	if err != nil {
		code, exitCode := formExportImageArgFailure(err)
		return output.Failure("form export-image", output.Error{Code: code, Message: err.Error(), Source: "xlflow"}), exitCode, nil
	}
	env, code, runErr := r.run("form-export-image", scriptArgs, opts.Keepalive)
	env.Command = "form export-image"
	return env, code, runErr
}

func (r Runner) FormWrite(cfg config.Config, opts FormWriteOptions) (output.Envelope, int, error) {
	scriptArgs, err := buildFormWriteScriptArgs(r.RootDir, cfg, opts)
	if err != nil {
		code, exitCode := formWriteArgFailure(err, opts.Action)
		return output.Failure("form "+opts.Action, output.Error{Code: code, Message: err.Error(), Source: "xlflow"}), exitCode, nil
	}
	env, code, runErr := r.run("form-write", scriptArgs, opts.Keepalive)
	env.Command = "form " + opts.Action
	return env, code, runErr
}

func buildListFormsScriptArgs(root string, cfg config.Config, opts SessionCommandOptions) map[string]string {
	return map[string]string{
		"Action":                  "forms",
		"WorkbookPath":            workbookPath(root, cfg.Excel.Path),
		"FormsDir":                filepath.Join(root, cfg.Src.Forms),
		"ModulesDir":              filepath.Join(root, cfg.Src.Modules),
		"ClassesDir":              filepath.Join(root, cfg.Src.Classes),
		"WorkbookDir":             filepath.Join(root, cfg.Src.Workbook),
		"ProjectRoot":             root,
		"Folders":                 strconv.FormatBool(cfg.VBA.Folders),
		"FolderAnnotation":        cfg.VBA.FolderAnnotation,
		"DefaultComponentFolders": strconv.FormatBool(cfg.VBA.DefaultComponentFolders),
		"Visible":                 strconv.FormatBool(cfg.Excel.Visible),
		"UseSession":              strconv.FormatBool(opts.Session),
		"MetadataPath":            filepath.Join(root, ".xlflow", "session.json"),
	}
}

func buildInspectFormScriptArgs(root string, cfg config.Config, opts InspectFormOptions) map[string]string {
	return map[string]string{
		"WorkbookPath":   workbookPath(root, cfg.Excel.Path),
		"Visible":        strconv.FormatBool(cfg.Excel.Visible),
		"UseSession":     strconv.FormatBool(opts.Session),
		"MetadataPath":   filepath.Join(root, ".xlflow", "session.json"),
		"FormName":       opts.Name,
		"Basis":          opts.Basis,
		"StrictDesigner": strconv.FormatBool(opts.StrictDesigner),
		"Initializer":    opts.Initializer,
	}
}

func buildInspectScriptArgs(root string, cfg config.Config, opts InspectOptions) map[string]string {
	args := map[string]string{
		"WorkbookPath": workbookPath(root, cfg.Excel.Path),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
		"UseSession":   strconv.FormatBool(opts.Session),
		"MetadataPath": filepath.Join(root, ".xlflow", "session.json"),
		"Target":       opts.Target,
		"Sheet":        opts.Sheet,
		"Address":      opts.Address,
		"IncludeStyle": strconv.FormatBool(opts.IncludeStyle),
		"MaxRows":      strconv.Itoa(opts.Limits["max_rows"]),
		"MaxCols":      strconv.Itoa(opts.Limits["max_cols"]),
	}
	return args
}

type formExportImageResolvedOutput struct {
	Path string
}

type formWriteArgError struct {
	code     string
	message  string
	exitCode int
}

func (e formWriteArgError) Error() string {
	return e.message
}

type formExportImageArgError struct {
	code     string
	message  string
	exitCode int
}

func (e formExportImageArgError) Error() string {
	return e.message
}

func formExportImageArgFailure(err error) (string, int) {
	var argErr formExportImageArgError
	if errors.As(err, &argErr) {
		return argErr.code, argErr.exitCode
	}
	return "form_export_image_args_invalid", output.ExitConfig
}

func formWriteArgFailure(err error, action string) (string, int) {
	var argErr formWriteArgError
	if errors.As(err, &argErr) {
		return argErr.code, argErr.exitCode
	}
	switch action {
	case "apply":
		return "form_apply_args_invalid", output.ExitConfig
	default:
		return "form_build_args_invalid", output.ExitConfig
	}
}

func buildFormExportImageScriptArgs(root string, cfg config.Config, opts FormExportImageOptions) (map[string]string, error) {
	resolvedOutput, err := resolveFormExportImageOutput(root, opts)
	if err != nil {
		return nil, err
	}
	return map[string]string{
		"WorkbookPath": workbookPath(root, cfg.Excel.Path),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
		"FormName":     opts.Name,
		"OutputPath":   resolvedOutput.Path,
		"Overwrite":    strconv.FormatBool(opts.Overwrite),
		"Initializer":  opts.Initializer,
		"UseSession":   strconv.FormatBool(opts.Session),
		"MetadataPath": filepath.Join(root, ".xlflow", "session.json"),
	}, nil
}

func buildFormWriteScriptArgs(root string, cfg config.Config, opts FormWriteOptions) (map[string]string, error) {
	action := strings.ToLower(strings.TrimSpace(opts.Action))
	if action != "build" && action != "apply" {
		return nil, formWriteArgError{
			code:     "form_build_args_invalid",
			message:  fmt.Sprintf("unsupported form action %q", opts.Action),
			exitCode: output.ExitConfig,
		}
	}
	specJSON, err := json.Marshal(opts.Spec)
	if err != nil {
		return nil, formWriteArgError{
			code:     formWriteArgsCode(action),
			message:  fmt.Sprintf("failed to serialize form spec: %v", err),
			exitCode: output.ExitConfig,
		}
	}
	return map[string]string{
		"Action":                  action,
		"WorkbookPath":            workbookPath(root, cfg.Excel.Path),
		"FormsDir":                filepath.Join(root, cfg.Src.Forms),
		"CodeSource":              cfg.UserForm.CodeSource,
		"Folders":                 strconv.FormatBool(cfg.VBA.Folders),
		"FolderAnnotation":        cfg.VBA.FolderAnnotation,
		"DefaultComponentFolders": strconv.FormatBool(cfg.VBA.DefaultComponentFolders),
		"Visible":                 strconv.FormatBool(cfg.Excel.Visible),
		"SpecPath":                opts.SpecPath,
		"SpecJson64":              base64.StdEncoding.EncodeToString(specJSON),
		"Overwrite":               strconv.FormatBool(opts.Overwrite),
		"NoSave":                  strconv.FormatBool(opts.NoSave),
		"UseSession":              strconv.FormatBool(opts.Session),
		"MetadataPath":            filepath.Join(root, ".xlflow", "session.json"),
	}, nil
}

func formWriteArgsCode(action string) string {
	if action == "apply" {
		return "form_apply_args_invalid"
	}
	return "form_build_args_invalid"
}

func resolveFormExportImageOutput(root string, opts FormExportImageOptions) (formExportImageResolvedOutput, error) {
	if strings.TrimSpace(opts.OutPath) == "" {
		return formExportImageResolvedOutput{}, formExportImageArgError{
			code:     "form_export_image_args_invalid",
			message:  "output path is required",
			exitCode: output.ExitConfig,
		}
	}
	trimmed := strings.TrimSpace(opts.OutPath)
	if filepath.Ext(trimmed) == "" {
		return formExportImageResolvedOutput{}, formExportImageArgError{
			code:     "unsupported_image_format",
			message:  `unsupported image format ""; supported formats: png`,
			exitCode: output.ExitValidation,
		}
	}
	path, format, err := normalizeExportImagePath(root, opts.OutPath, "png")
	if err != nil {
		var exportErr exportImageArgError
		if errors.As(err, &exportErr) {
			return formExportImageResolvedOutput{}, formExportImageArgError(exportErr)
		}
		return formExportImageResolvedOutput{}, err
	}
	if format != "png" {
		return formExportImageResolvedOutput{}, formExportImageArgError{
			code:     "unsupported_image_format",
			message:  fmt.Sprintf("unsupported image format %q; supported formats: png", format),
			exitCode: output.ExitValidation,
		}
	}
	return formExportImageResolvedOutput{Path: path}, nil
}

func (r Runner) UIButtonAdd(cfg config.Config, opts UIButtonAddOptions, cmdOpts ...CommandOptions) (output.Envelope, int, error) {
	return r.run("ui", buildUIButtonAddScriptArgs(r.RootDir, cfg, opts), cmdOpts...)
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
		"UseSession":   strconv.FormatBool(opts.Session),
		"MetadataPath": filepath.Join(root, ".xlflow", "session.json"),
	}
}

func (r Runner) UIButtonList(cfg config.Config, opts UIButtonListOptions, cmdOpts ...CommandOptions) (output.Envelope, int, error) {
	return r.run("ui", buildUIButtonListScriptArgs(r.RootDir, cfg, opts), cmdOpts...)
}

func buildUIButtonListScriptArgs(root string, cfg config.Config, opts UIButtonListOptions) map[string]string {
	return map[string]string{
		"Action":       "list",
		"WorkbookPath": workbookPath(root, cfg.Excel.Path),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
		"Sheet":        opts.Sheet,
		"UseSession":   strconv.FormatBool(opts.Session),
		"MetadataPath": filepath.Join(root, ".xlflow", "session.json"),
	}
}

func (r Runner) UIButtonRemove(cfg config.Config, opts UIButtonRemoveOptions, cmdOpts ...CommandOptions) (output.Envelope, int, error) {
	return r.run("ui", buildUIButtonRemoveScriptArgs(r.RootDir, cfg, opts), cmdOpts...)
}

func (r Runner) ExportImage(cfg config.Config, opts ExportImageOptions) (output.Envelope, int, error) {
	scriptArgs, err := buildExportImageScriptArgs(r.RootDir, cfg, opts)
	if err != nil {
		code, exitCode := exportImageArgFailure(err)
		return output.Failure("export-image", output.Error{Code: code, Message: err.Error(), Source: "xlflow"}), exitCode, nil
	}
	return r.run("export-image", scriptArgs, opts.Keepalive)
}

func (r Runner) ProcessList(opts ProcessListOptions) (output.Envelope, int, error) {
	env, code, err := r.run("process", map[string]string{
		"Action": opts.Action,
	})
	env.Command = "process list"
	return env, code, err
}

func (r Runner) ProcessCleanup(opts ProcessCleanupOptions) (output.Envelope, int, error) {
	args := map[string]string{
		"Action": opts.Action,
		"Auto":   strconv.FormatBool(opts.Auto),
		"All":    strconv.FormatBool(opts.All),
	}
	if opts.PID > 0 {
		args["TargetPid"] = strconv.Itoa(opts.PID)
	}
	env, code, err := r.run("process", args)
	env.Command = "process cleanup"
	return env, code, err
}

func (r Runner) EditCell(cfg config.Config, opts EditCellOptions) (output.Envelope, int, error) {
	scriptArgs, err := buildEditCellScriptArgs(r.RootDir, cfg, opts)
	if err != nil {
		code, exitCode := editArgFailure(err)
		return output.Failure("edit", output.Error{Code: code, Message: err.Error(), Source: "xlflow"}), exitCode, nil
	}
	return r.run("edit", scriptArgs, opts.Keepalive)
}

func (r Runner) EditRange(cfg config.Config, opts EditRangeOptions) (output.Envelope, int, error) {
	scriptArgs, err := buildEditRangeScriptArgs(r.RootDir, cfg, opts)
	if err != nil {
		code, exitCode := editArgFailure(err)
		return output.Failure("edit", output.Error{Code: code, Message: err.Error(), Source: "xlflow"}), exitCode, nil
	}
	return r.run("edit", scriptArgs, opts.Keepalive)
}

func (r Runner) EditRows(cfg config.Config, opts EditRowsOptions) (output.Envelope, int, error) {
	return r.run("edit", buildEditRowsScriptArgs(r.RootDir, cfg, opts), opts.Keepalive)
}

func (r Runner) EditColumns(cfg config.Config, opts EditColumnsOptions) (output.Envelope, int, error) {
	return r.run("edit", buildEditColumnsScriptArgs(r.RootDir, cfg, opts), opts.Keepalive)
}

func buildUIButtonRemoveScriptArgs(root string, cfg config.Config, opts UIButtonRemoveOptions) map[string]string {
	return map[string]string{
		"Action":       "remove",
		"WorkbookPath": workbookPath(root, cfg.Excel.Path),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
		"Sheet":        opts.Sheet,
		"Id":           opts.ID,
		"UseSession":   strconv.FormatBool(opts.Session),
		"MetadataPath": filepath.Join(root, ".xlflow", "session.json"),
	}
}

func buildEditCellScriptArgs(root string, cfg config.Config, opts EditCellOptions) (map[string]string, error) {
	mutations := 0
	if opts.Value != nil {
		mutations++
	}
	if opts.Formula != nil {
		mutations++
	}
	if opts.Fill != "" {
		mutations++
	}
	if mutations != 1 {
		return nil, editArgError{message: "exactly one of value, formula, or fill is required"}
	}
	args := map[string]string{
		"Action":       "cell",
		"WorkbookPath": workbookPath(root, chooseEditWorkbook(cfg, opts.WorkbookPath)),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
		"Sheet":        opts.Sheet,
		"Cell":         opts.Cell,
		"Events":       string(opts.Events),
		"UseSession":   strconv.FormatBool(opts.Session),
		"MetadataPath": filepath.Join(root, ".xlflow", "session.json"),
	}
	if opts.Value != nil {
		args["Value"] = *opts.Value
	}
	if opts.Formula != nil {
		args["Formula"] = *opts.Formula
	}
	if opts.Fill != "" {
		args["Fill"] = opts.Fill
	}
	return args, nil
}

func buildEditRangeScriptArgs(root string, cfg config.Config, opts EditRangeOptions) (map[string]string, error) {
	if opts.Fill != "" && opts.Clear != "" {
		return nil, editArgError{message: "fill and clear cannot be combined"}
	}
	args := map[string]string{
		"Action":       "range",
		"WorkbookPath": workbookPath(root, chooseEditWorkbook(cfg, opts.WorkbookPath)),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
		"Sheet":        opts.Sheet,
		"RangeAddress": opts.Range,
		"UseSession":   strconv.FormatBool(opts.Session),
		"MetadataPath": filepath.Join(root, ".xlflow", "session.json"),
	}
	if opts.Fill != "" {
		args["Fill"] = opts.Fill
	}
	if opts.Clear != "" {
		args["Clear"] = opts.Clear
	}
	return args, nil
}

func buildEditRowsScriptArgs(root string, cfg config.Config, opts EditRowsOptions) map[string]string {
	return map[string]string{
		"Action":       "rows",
		"WorkbookPath": workbookPath(root, chooseEditWorkbook(cfg, opts.WorkbookPath)),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
		"Sheet":        opts.Sheet,
		"Rows":         opts.Rows,
		"Height":       strconv.FormatFloat(opts.Height, 'f', -1, 64),
		"UseSession":   strconv.FormatBool(opts.Session),
		"MetadataPath": filepath.Join(root, ".xlflow", "session.json"),
	}
}

func buildEditColumnsScriptArgs(root string, cfg config.Config, opts EditColumnsOptions) map[string]string {
	return map[string]string{
		"Action":       "columns",
		"WorkbookPath": workbookPath(root, chooseEditWorkbook(cfg, opts.WorkbookPath)),
		"Visible":      strconv.FormatBool(cfg.Excel.Visible),
		"Sheet":        opts.Sheet,
		"Columns":      opts.Columns,
		"Width":        strconv.FormatFloat(opts.Width, 'f', -1, 64),
		"UseSession":   strconv.FormatBool(opts.Session),
		"MetadataPath": filepath.Join(root, ".xlflow", "session.json"),
	}
}

func chooseEditWorkbook(cfg config.Config, workbook string) string {
	if workbook != "" {
		return workbook
	}
	return cfg.Excel.Path
}

type exportImageResolvedOutput struct {
	Path    string
	Format  string
	Default bool
}

type exportImageArgError struct {
	code     string
	message  string
	exitCode int
}

func (e exportImageArgError) Error() string {
	return e.message
}

func exportImageArgFailure(err error) (string, int) {
	var argErr exportImageArgError
	if errors.As(err, &argErr) {
		return argErr.code, argErr.exitCode
	}
	return "export_image_args_invalid", output.ExitConfig
}

type editArgError struct {
	message string
}

func (e editArgError) Error() string {
	return e.message
}

func editArgFailure(err error) (string, int) {
	var argErr editArgError
	if errors.As(err, &argErr) {
		return "edit_args_invalid", output.ExitConfig
	}
	return "edit_args_invalid", output.ExitConfig
}

func buildExportImageScriptArgs(root string, cfg config.Config, opts ExportImageOptions) (map[string]string, error) {
	workbook := cfg.Excel.Path
	if opts.WorkbookPath != "" {
		workbook = opts.WorkbookPath
	}
	outputPath, err := resolveExportImageOutput(root, workbook, opts)
	if err != nil {
		return nil, err
	}
	return map[string]string{
		"WorkbookPath":    workbookPath(root, workbook),
		"Visible":         strconv.FormatBool(cfg.Excel.Visible),
		"Sheet":           opts.Sheet,
		"RangeAddress":    opts.Range,
		"OutputPath":      outputPath.Path,
		"OutputIsDefault": strconv.FormatBool(outputPath.Default),
		"ImageFormat":     outputPath.Format,
		"Overwrite":       strconv.FormatBool(opts.Overwrite),
		"UseSession":      strconv.FormatBool(opts.Session),
		"MetadataPath":    filepath.Join(root, ".xlflow", "session.json"),
	}, nil
}

func resolveExportImageOutput(root, workbook string, opts ExportImageOptions) (exportImageResolvedOutput, error) {
	format, err := normalizeExportImageFormat(opts.Format)
	if err != nil {
		return exportImageResolvedOutput{}, err
	}
	if opts.OutPath != "" {
		if opts.OutputDir != "" || opts.Name != "" {
			return exportImageResolvedOutput{}, fmt.Errorf("--out cannot be combined with --output-dir or --name")
		}
		path, format, err := normalizeExportImagePath(root, opts.OutPath, format)
		if err != nil {
			return exportImageResolvedOutput{}, err
		}
		return exportImageResolvedOutput{Path: path, Format: format, Default: false}, nil
	}

	dir := opts.OutputDir
	if dir == "" {
		dir = filepath.Join(".xlflow", "artifacts", "images", sanitizeExportImageComponent(strings.TrimSuffix(filepath.Base(workbook), filepath.Ext(workbook)), "workbook"))
	}
	filename := opts.Name
	if filename == "" {
		filename = defaultExportImageFilename(opts.Sheet, opts.Range, format)
	}
	path, format, err := normalizeExportImagePath(root, filepath.Join(dir, filename), format)
	if err != nil {
		return exportImageResolvedOutput{}, err
	}
	return exportImageResolvedOutput{
		Path:    path,
		Format:  format,
		Default: opts.OutputDir == "" && opts.Name == "",
	}, nil
}

func normalizeExportImageFormat(format string) (string, error) {
	switch strings.ToLower(strings.TrimSpace(format)) {
	case "", "png":
		return "png", nil
	default:
		return "", exportImageArgError{
			code:     "unsupported_image_format",
			message:  fmt.Sprintf("unsupported image format %q; supported formats: png", format),
			exitCode: output.ExitValidation,
		}
	}
}

func normalizeExportImagePath(root, path, format string) (string, string, error) {
	if strings.TrimSpace(path) == "" {
		return "", "", fmt.Errorf("output path is required")
	}
	resolved := path
	if !filepath.IsAbs(resolved) {
		resolved = filepath.Join(root, resolved)
	}
	ext := strings.ToLower(filepath.Ext(resolved))
	switch ext {
	case "":
		resolved += "." + format
	case ".png":
		format = "png"
	default:
		return "", "", exportImageArgError{
			code:     "unsupported_image_format",
			message:  fmt.Sprintf("unsupported image format %q; supported formats: png", strings.TrimPrefix(ext, ".")),
			exitCode: output.ExitValidation,
		}
	}
	return filepath.Clean(resolved), format, nil
}

func defaultExportImageFilename(sheet, cellRange, format string) string {
	name := sanitizeExportImageComponent(sheet, "sheet")
	rangeName := sanitizeExportImageComponent(strings.ReplaceAll(cellRange, ":", "-"), "range")
	return fmt.Sprintf("%s_%s_%s.%s", name, rangeName, time.Now().Format("20060102-150405"), format)
}

func sanitizeExportImageComponent(value, fallback string) string {
	trimmed := strings.TrimSpace(value)
	if trimmed == "" {
		return fallback
	}
	var b strings.Builder
	lastUnderscore := false
	for _, r := range trimmed {
		switch {
		case r == ':' || r == '-' || r == '_':
			if b.Len() > 0 && !lastUnderscore {
				b.WriteRune(r)
				lastUnderscore = r == '_'
			}
		case unicode.IsSpace(r):
			if b.Len() > 0 && !lastUnderscore {
				b.WriteByte('_')
				lastUnderscore = true
			}
		case strings.ContainsRune(`\/:*?"<>|`, r):
			if b.Len() > 0 && !lastUnderscore {
				b.WriteByte('_')
				lastUnderscore = true
			}
		default:
			b.WriteRune(r)
			lastUnderscore = false
		}
	}
	out := strings.Trim(b.String(), "._- ")
	if out == "" {
		return fallback
	}
	return out
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
	var err error
	if runtime.GOOS != "windows" {
		env = output.Failure(commandName, output.Error{Code: "environment", Message: "Excel automation is only supported on Windows in the MVP"})
		return env, output.ExitEnvironment, nil
	}

	var uiSession *uiStreamSession
	var debugSession *debugStreamSession
	if enabled, _ := strconv.ParseBool(strings.TrimSpace(args["UIStreamEnabled"])); enabled {
		uiSession, err = newUIStreamSession(opts.Keepalive.Stderr)
		if err != nil {
			env = output.Failure(commandName, output.Error{Code: "ui_stream_init_failed", Message: err.Error(), Source: "xlflow"})
			return env, output.ExitEnvironment, nil
		}
		args = cloneStringMap(args)
		args["UIStreamPipeName"] = uiSession.PipePath()
	}
	if enabled, _ := strconv.ParseBool(strings.TrimSpace(args["DebugStreamEnabled"])); enabled {
		debugSession, err = newDebugStreamSession(opts.Keepalive.Stderr)
		if err != nil {
			env = output.Failure(commandName, output.Error{Code: "debug_stream_init_failed", Message: err.Error(), Source: "xlflow"})
			if uiEvents, uiStreamErr := closeUIStreamSession(uiSession); len(uiEvents) > 0 || uiStreamErr != nil {
				env.UI = mergeUIResult(nil, uiEvents)
				if uiStreamErr != nil {
					env.Logs = append(env.Logs, "UI stream closed with an error: "+uiStreamErr.Error())
				}
			}
			return env, output.ExitEnvironment, nil
		}
		if uiSession == nil {
			args = cloneStringMap(args)
		}
		args["DebugStreamPipeName"] = debugSession.PipePath()
	}

	script, cleanup, err := scriptPath(r.RootDir, commandName)
	if err != nil {
		env = output.Failure(commandName, output.Error{Code: "script_not_found", Message: err.Error(), Source: "xlflow"})
		if debugResult, debugStreamErr := closeDebugStreamSession(debugSession); debugResult != nil || debugStreamErr != nil {
			env.Debug = mergeDebugResult(nil, debugResult)
			if debugStreamErr != nil {
				env.Logs = append(env.Logs, "Debug stream closed with an error: "+debugStreamErr.Error())
			}
		}
		if uiEvents, uiStreamErr := closeUIStreamSession(uiSession); len(uiEvents) > 0 || uiStreamErr != nil {
			env.UI = mergeUIResult(nil, uiEvents)
			if uiStreamErr != nil {
				env.Logs = append(env.Logs, "UI stream closed with an error: "+uiStreamErr.Error())
			}
		}
		return env, output.ExitEnvironment, nil
	}
	if cleanup != nil {
		defer cleanup()
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
	err = cmd.Start()
	if err == nil {
		err = cmd.Wait()
	}
	debugResult, debugStreamErr := closeDebugStreamSession(debugSession)
	uiEvents, uiStreamErr := closeUIStreamSession(uiSession)
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
			if debugStreamErr != nil {
				env.Logs = append(env.Logs, "Debug stream closed with an error: "+debugStreamErr.Error())
			}
			env.Debug = mergeDebugResult(nil, debugResult)
			if uiStreamErr != nil {
				env.Logs = append(env.Logs, "UI stream closed with an error: "+uiStreamErr.Error())
			}
			env.UI = mergeUIResult(nil, uiEvents)
			return env, output.ExitValidation, nil
		}
		message := err.Error()
		if stderr.Len() > 0 {
			message = stderr.String()
		}
		env = output.Failure(commandName, output.Error{Code: "script_failed", Message: message, Source: "powershell"})
		if debugStreamErr != nil {
			env.Logs = append(env.Logs, "Debug stream closed with an error: "+debugStreamErr.Error())
		}
		env.Debug = mergeDebugResult(nil, debugResult)
		if uiStreamErr != nil {
			env.Logs = append(env.Logs, "UI stream closed with an error: "+uiStreamErr.Error())
		}
		env.UI = mergeUIResult(nil, uiEvents)
		return env, output.ExitEnvironment, nil
	}

	var result ScriptResult
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		env = output.Failure(commandName, output.Error{Code: "invalid_script_json", Message: fmt.Sprintf("failed to parse script JSON: %v", err), Source: "powershell"})
		env.Logs = []string{stdout.String()}
		if uiStreamErr != nil {
			env.Logs = append(env.Logs, "UI stream closed with an error: "+uiStreamErr.Error())
		}
		env.UI = mergeUIResult(nil, uiEvents)
		return env, output.ExitEnvironment, nil
	}
	if result.Status == "" {
		result.Status = output.StatusOK
	}
	env.Status = result.Status
	env.Command = commandName
	env.Error = result.Error
	env.Logs = []string(result.Logs)
	if env.Logs == nil {
		env.Logs = []string{}
	}
	env.Diagnostics = result.Diagnostics
	env.Workbook = result.Workbook
	env.Backup = result.Backup
	env.Source = result.Source
	env.Bridge = result.Bridge
	env.Macro = result.Macro
	env.Macros = result.Macros
	env.Forms = result.Forms
	env.Tests = result.Tests
	env.Trace = result.Trace
	env.Runtime = result.Runtime
	env.GUIBoundaries = result.GUIBoundaries
	env.Debug = mergeDebugResult(result.Debug, debugResult)
	env.UI = mergeUIResult(result.UI, uiEvents)
	env.Session = result.Session
	env.Runner = result.Runner
	env.Analysis = result.Analysis
	env.Check = result.Check
	env.RunDiagnostic = result.RunDiagnostic
	env.Target = result.Target
	env.Output = result.Output
	if debugStreamErr != nil {
		env.Logs = append(env.Logs, "Debug stream closed with an error: "+debugStreamErr.Error())
	}
	env.Spec = result.Spec
	env.Edit = result.Edit
	env.Warnings = result.Warnings
	env.Hints = result.Hints
	env.Inspect = result.Inspect
	if result.DefaultEntry != "" {
		env.DefaultEntry = result.DefaultEntry
	}
	if result.Suggestions != nil {
		env.Suggestions = result.Suggestions
	}
	env.Process = result.Process
	if uiStreamErr != nil {
		env.Logs = append(env.Logs, "UI stream closed with an error: "+uiStreamErr.Error())
	}
	if result.Status == output.StatusFailed {
		return env, exitCodeForScriptResult(result), nil
	}
	return env, output.ExitSuccess, nil
}

func cloneStringMap(values map[string]string) map[string]string {
	if values == nil {
		return nil
	}
	cloned := make(map[string]string, len(values))
	for k, v := range values {
		cloned[k] = v
	}
	return cloned
}

func mergeUIResult(existing any, streamed []map[string]any) any {
	if len(streamed) == 0 {
		return existing
	}
	if existing == nil {
		return map[string]any{"events": streamed}
	}
	existingMap, ok := existing.(map[string]any)
	if !ok {
		return map[string]any{"summary": existing, "events": streamed}
	}
	merged := make(map[string]any, len(existingMap)+1)
	for k, v := range existingMap {
		merged[k] = v
	}
	if prior, ok := merged["events"].([]map[string]any); ok {
		merged["events"] = append(prior, streamed...)
		return merged
	}
	if priorAny, ok := merged["events"].([]any); ok {
		mergedEvents := make([]any, 0, len(priorAny)+len(streamed))
		mergedEvents = append(mergedEvents, priorAny...)
		for _, event := range streamed {
			mergedEvents = append(mergedEvents, event)
		}
		merged["events"] = mergedEvents
		return merged
	}
	merged["events"] = streamed
	return merged
}

func mergeDebugResult(existing any, streamed any) any {
	if streamed == nil {
		return existing
	}
	streamedMap, ok := streamed.(map[string]any)
	if !ok {
		return existing
	}
	if existing == nil {
		return streamedMap
	}
	existingMap, ok := existing.(map[string]any)
	if !ok {
		return streamedMap
	}
	merged := make(map[string]any, len(existingMap)+len(streamedMap))
	for k, v := range existingMap {
		merged[k] = v
	}
	prior := debugEventList(existingMap["events"])
	for k, v := range streamedMap {
		merged[k] = v
	}
	additional := debugEventList(streamedMap["events"])
	if len(prior) > 0 || len(additional) > 0 {
		combined := make([]map[string]any, 0, len(prior)+len(additional))
		combined = append(combined, prior...)
		combined = append(combined, additional...)
		merged["events"] = combined
	}
	count := len(debugEventList(merged["events"]))
	if n, ok := numberValueFromAny(existingMap["count"]); ok && int(n) > count {
		count = int(n)
	}
	if n, ok := numberValueFromAny(streamedMap["count"]); ok && int(n) > count {
		count = int(n)
	}
	merged["count"] = count
	if truthyValue(existingMap["truncated"]) || truthyValue(streamedMap["truncated"]) {
		merged["truncated"] = true
	}
	return merged
}

func numberValueFromAny(value any) (float64, bool) {
	switch typed := value.(type) {
	case int:
		return float64(typed), true
	case int32:
		return float64(typed), true
	case int64:
		return float64(typed), true
	case float32:
		return float64(typed), true
	case float64:
		return typed, true
	default:
		return 0, false
	}
}

func truthyValue(value any) bool {
	switch typed := value.(type) {
	case bool:
		return typed
	case string:
		return strings.EqualFold(strings.TrimSpace(typed), "true")
	default:
		return false
	}
}

func debugEventList(value any) []map[string]any {
	switch typed := value.(type) {
	case []map[string]any:
		return typed
	case []any:
		result := make([]map[string]any, 0, len(typed))
		for _, item := range typed {
			if event, ok := item.(map[string]any); ok {
				result = append(result, event)
			}
		}
		return result
	default:
		return nil
	}
}

func closeUIStreamSession(session *uiStreamSession) ([]map[string]any, error) {
	if session == nil {
		return nil, nil
	}
	if err := session.Close(); err != nil {
		return session.Events(), err
	}
	return session.Events(), nil
}

func closeDebugStreamSession(session *debugStreamSession) (any, error) {
	if session == nil {
		return nil, nil
	}
	if err := session.Close(); err != nil {
		return session.Result(), err
	}
	return session.Result(), nil
}

func exitCodeForScriptResult(result ScriptResult) int {
	if result.Error == nil {
		return output.ExitEnvironment
	}
	switch result.Error.Code {
	case "macro_failed", "macro_disabled", "macro_not_found", "macro_timeout", "vba_compile_failed", "trace_not_injected", "trace_source_modified", "trace_args_invalid", "test_failed", "no_tests_found", "test_not_found", "duplicate_test_name", "active_workbook_mismatch", "sheet_not_found", "button_not_found", "ui_button_args_invalid", "duplicate_module_name", "invalid_range", "output_file_exists", "unsupported_image_format", "session_required", "invalid_color", "invalid_cell_address", "invalid_row_selector", "invalid_column_selector", "vba_event_error", "form_not_found", "runtime_form_load_failed", "form_initializer_failed", "control_enumeration_failed", "window_not_found", "image_capture_failed":
		return output.ExitValidation
	case "form_already_exists", "unsupported_form_control", "designer_write_failed":
		return output.ExitValidation
	case "process_args_invalid", "process_not_found":
		return output.ExitConfig
	case "process_enumeration_failed", "process_termination_failed", "process_cleanup_failed":
		return output.ExitEnvironment
	case "push_args_invalid", "run_args_invalid", "session_args_invalid", "runner_args_invalid", "export_image_args_invalid", "edit_args_invalid", "list_args_invalid", "inspect_args_invalid", "inspect_form_args_invalid", "form_export_image_args_invalid", "form_build_args_invalid", "form_apply_args_invalid":
		return output.ExitConfig
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

func scriptPath(root, commandName string) (string, func(), error) {
	if path, ok := externalScriptPath(root, commandName); ok {
		return path, nil, nil
	}
	return materializeBundledScript(commandName)
}

func materializeBundledScript(commandName string) (string, func(), error) {
	name := commandName + ".ps1"
	path, cleanup, err := bundledscripts.Materialize(commandName)
	if err != nil {
		return "", nil, fmt.Errorf("script %s was not available from on-disk script locations or embedded runtime assets: %w", name, err)
	}
	return path, cleanup, nil
}

func externalScriptPath(root, commandName string) (string, bool) {
	name := commandName + ".ps1"
	candidates := []string{}
	if root != "" {
		candidates = append(candidates, filepath.Join(root, "scripts", name))
	}
	if _, file, _, ok := runtime.Caller(0); ok {
		candidates = append(candidates, filepath.Join(filepath.Dir(file), "scripts", name))
	}
	if exe, err := os.Executable(); err == nil {
		candidates = append(candidates, filepath.Join(filepath.Dir(exe), "scripts", name))
	}
	for _, candidate := range candidates {
		clean := filepath.Clean(candidate)
		if _, err := os.Stat(clean); err == nil {
			return clean, true
		}
	}
	return "", false
}
