package output

import (
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"path/filepath"
	"sort"
	"strings"

	"github.com/charmbracelet/lipgloss"
)

const (
	StatusOK     = "ok"
	StatusFailed = "failed"

	ExitSuccess     = 0
	ExitValidation  = 1
	ExitConfig      = 2
	ExitEnvironment = 3
)

type Error struct {
	Code    string `json:"code,omitempty"`
	Message string `json:"message"`
	Source  string `json:"source,omitempty"`
	Number  int    `json:"number,omitempty"`
	Line    int    `json:"line,omitempty"`
	Phase   string `json:"phase,omitempty"`
	HResult string `json:"h_result,omitempty"`
	Details any    `json:"details,omitempty"`
}

type Envelope struct {
	Status  string   `json:"status"`
	Command string   `json:"command"`
	Error   *Error   `json:"error"`
	Logs    []string `json:"logs"`

	Diagnostics    any `json:"diagnostics,omitempty"`
	Workbook       any `json:"workbook,omitempty"`
	Backup         any `json:"backup,omitempty"`
	Source         any `json:"source,omitempty"`
	Bridge         any `json:"bridge,omitempty"`
	Macro          any `json:"macro,omitempty"`
	Macros         any `json:"macros,omitempty"`
	Forms          any `json:"forms,omitempty"`
	Issues         any `json:"issues,omitempty"`
	Tests          any `json:"tests,omitempty"`
	Diff           any `json:"diff,omitempty"`
	Inspect        any `json:"inspect,omitempty"`
	Runtime        any `json:"runtime,omitempty"`
	GUIBoundaries  any `json:"gui_boundaries,omitempty"`
	Debug          any `json:"debug,omitempty"`
	UI             any `json:"ui,omitempty"`
	Session        any `json:"session,omitempty"`
	Runner         any `json:"runner,omitempty"`
	Analysis       any `json:"analysis,omitempty"`
	Check          any `json:"check,omitempty"`
	Version        any `json:"version,omitempty"`
	RunDiagnostic  any `json:"run_diagnostic,omitempty"`
	PushDiagnostic any `json:"push_diagnostic,omitempty"`
	Backups        any `json:"backups,omitempty"`
	Rollback       any `json:"rollback,omitempty"`
	Target         any `json:"target,omitempty"`
	Output         any `json:"output,omitempty"`
	Pack           any `json:"pack,omitempty"`
	Spec           any `json:"spec,omitempty"`
	Edit           any `json:"edit,omitempty"`
	Project        any `json:"project,omitempty"`
	State          any `json:"state,omitempty"`
	Warnings       any `json:"warnings,omitempty"`
	Hints          any `json:"hints,omitempty"`
	DefaultEntry   any `json:"default_entry,omitempty"`
	Suggestions    any `json:"suggestions,omitempty"`
	Process        any `json:"process,omitempty"`
}

type Options struct {
	JSON        bool
	Interactive bool
	Color       bool
	Verbose     bool
}

type ExitError struct {
	Code int
	Err  error
}

func (e *ExitError) Error() string {
	if e.Err == nil {
		return fmt.Sprintf("exit code %d", e.Code)
	}
	return e.Err.Error()
}

func (e *ExitError) Unwrap() error {
	return e.Err
}

func WithExitCode(code int, err error) error {
	if err == nil {
		return nil
	}
	return &ExitError{Code: code, Err: err}
}

func ExitCode(err error) int {
	if err == nil {
		return ExitSuccess
	}
	var exitErr *ExitError
	if errors.As(err, &exitErr) {
		return exitErr.Code
	}
	return ExitConfig
}

func New(command string) Envelope {
	return Envelope{
		Status:  StatusOK,
		Command: command,
		Error:   nil,
		Logs:    []string{},
	}
}

func Failure(command string, err Error) Envelope {
	return Envelope{
		Status:  StatusFailed,
		Command: command,
		Error:   &err,
		Logs:    []string{},
	}
}

func Write(w io.Writer, env Envelope, jsonOutput bool) error {
	if jsonOutput {
		enc := json.NewEncoder(w)
		enc.SetIndent("", "  ")
		return enc.Encode(env)
	}
	if env.Status == StatusOK {
		if len(env.Logs) == 0 {
			_, err := fmt.Fprintln(w, "ok")
			return err
		}
		for _, line := range env.Logs {
			if _, err := fmt.Fprintln(w, line); err != nil {
				return err
			}
		}
		return nil
	}
	if env.Error != nil {
		for _, line := range env.Logs {
			if _, err := fmt.Fprintln(w, line); err != nil {
				return err
			}
		}
		_, err := fmt.Fprintln(w, env.Error.Message)
		return err
	}
	_, err := fmt.Fprintln(w, "failed")
	return err
}

func WriteWithOptions(w io.Writer, env Envelope, opts Options) error {
	if opts.JSON {
		enc := json.NewEncoder(w)
		enc.SetIndent("", "  ")
		return enc.Encode(jsonPayload(env, opts))
	}
	text := renderHuman(env, opts)
	if text == "" {
		text = renderFallback(env)
	}
	_, err := fmt.Fprint(w, text)
	return err
}

func jsonPayload(env Envelope, opts Options) any {
	if env.Command == "run" {
		return runJSONPayload(env, opts.Verbose)
	}
	return env
}

func runJSONPayload(env Envelope, verbose bool) map[string]any {
	payload := map[string]any{
		"status":  env.Status,
		"command": env.Command,
	}
	if env.Error != nil {
		payload["error"] = env.Error
	}
	addPayloadField(payload, "issues", env.Issues)
	addPayloadField(payload, "analysis", env.Analysis)
	addPayloadField(payload, "gui_boundaries", env.GUIBoundaries)
	if verbose {
		if env.Logs != nil {
			payload["logs"] = env.Logs
		}
		addPayloadField(payload, "diagnostics", env.Diagnostics)
		addPayloadField(payload, "workbook", env.Workbook)
		addPayloadField(payload, "backup", env.Backup)
		addPayloadField(payload, "source", env.Source)
		addPayloadField(payload, "bridge", env.Bridge)
		addPayloadField(payload, "macros", env.Macros)
		addPayloadField(payload, "forms", env.Forms)
		addPayloadField(payload, "tests", env.Tests)
		addPayloadField(payload, "diff", env.Diff)
		addPayloadField(payload, "inspect", env.Inspect)
		addPayloadField(payload, "runtime", env.Runtime)
		addPayloadField(payload, "debug", env.Debug)
		addPayloadField(payload, "ui", env.UI)
		addPayloadField(payload, "runner", env.Runner)
		addPayloadField(payload, "check", env.Check)
		addPayloadField(payload, "version", env.Version)
		addPayloadField(payload, "push_diagnostic", env.PushDiagnostic)
		addPayloadField(payload, "backups", env.Backups)
		addPayloadField(payload, "rollback", env.Rollback)
		addPayloadField(payload, "output", env.Output)
		addPayloadField(payload, "spec", env.Spec)
		addPayloadField(payload, "edit", env.Edit)
		addPayloadField(payload, "project", env.Project)
		addPayloadField(payload, "state", env.State)
		addPayloadField(payload, "hints", env.Hints)
		addPayloadField(payload, "default_entry", env.DefaultEntry)
		addPayloadField(payload, "suggestions", env.Suggestions)
		addPayloadField(payload, "process", env.Process)
	}
	if macro := compactRunMacro(env.Macro, verbose); macro != nil {
		payload["macro"] = macro
	}
	if location := runLocation(env); location != nil {
		payload["location"] = location
	}
	if session := runSessionPayload(env.Session, verbose); session != nil {
		payload["session"] = session
	}
	addPayloadField(payload, "target", env.Target)
	addPayloadField(payload, "warnings", env.Warnings)
	if suggestion := runSuggestion(env); suggestion != "" {
		payload["suggestion"] = suggestion
	}
	if verbose {
		if diag := verboseRunDiagnostic(env.RunDiagnostic); diag != nil {
			payload["run_diagnostic"] = diag
		}
	}
	return payload
}

func addPayloadField(payload map[string]any, key string, value any) {
	if value != nil {
		payload[key] = value
	}
}

func compactRunMacro(value any, verbose bool) any {
	macro := clonedObjectMap(value)
	if len(macro) == 0 {
		return nil
	}
	if verbose {
		return macro
	}
	compact := map[string]any{}
	if name, ok := macro["name"]; ok && name != nil {
		compact["name"] = name
	}
	if duration, ok := macro["duration_ms"]; ok && duration != nil {
		compact["duration_ms"] = duration
	}
	if len(compact) == 0 {
		return nil
	}
	return compact
}

func runSessionPayload(value any, verbose bool) any {
	session := clonedObjectMap(value)
	if len(session) == 0 {
		return nil
	}
	if verbose {
		return session
	}
	compact := map[string]any{}
	for _, key := range []string{"active", "mode", "dirty", "save_required", "source_of_truth", "workbook_path"} {
		if item, ok := session[key]; ok && item != nil {
			compact[key] = item
		}
	}
	if len(compact) == 0 {
		return nil
	}
	return compact
}

func runLocation(env Envelope) any {
	diag := clonedObjectMap(env.RunDiagnostic)
	location := clonedObjectMap(diag["location"])
	if len(location) > 0 {
		return location
	}
	if env.Error == nil {
		return nil
	}
	if env.Error.Source == "" && env.Error.Line == 0 {
		return nil
	}
	location = map[string]any{}
	if env.Error.Source != "" {
		location["component"] = env.Error.Source
	}
	if env.Error.Line > 0 {
		location["line"] = env.Error.Line
		location["end_line"] = env.Error.Line
	}
	if len(location) == 0 {
		return nil
	}
	return location
}

func runSuggestion(env Envelope) string {
	diag := clonedObjectMap(env.RunDiagnostic)
	if suggestion, ok := diag["suggestion"].(string); ok {
		return strings.TrimSpace(suggestion)
	}
	return ""
}

func verboseRunDiagnostic(value any) any {
	diag := clonedObjectMap(value)
	if len(diag) == 0 {
		return nil
	}
	if dialogs, ok := diag["dialogs"]; ok && dialogs != nil {
		delete(diag, "dialog")
	}
	return diag
}

func clonedObjectMap(value any) map[string]any {
	if value == nil {
		return nil
	}
	decoded := objectMap(value)
	if len(decoded) == 0 {
		return nil
	}
	clone := make(map[string]any, len(decoded))
	for key, item := range decoded {
		clone[key] = item
	}
	return clone
}

func renderHuman(env Envelope, opts Options) string {
	r := renderer{color: opts.Color}
	if standalone := r.renderInspectStandalone(env); standalone != "" {
		return standalone
	}
	var b strings.Builder
	b.WriteString(r.title(env))
	b.WriteString("\n")
	if env.Status == StatusFailed {
		b.WriteString(r.errorBlock(env))
	}
	b.WriteString(r.renderBridge(env))
	if env.Command != "inspect" && env.Command != "status" {
		b.WriteString(r.renderTargetSession(env))
	}
	switch env.Command {
	case "version":
		b.WriteString(r.renderVersion(env))
	case "doctor":
		b.WriteString(r.renderDoctor(env))
	case "run":
		if env.Issues != nil {
			b.WriteString(r.renderLint(env))
		}
		if env.Analysis != nil {
			b.WriteString(r.renderAnalysis(env))
		}
		b.WriteString(r.renderRun(env))
	case "test":
		b.WriteString(r.renderTest(env))
	case "lint":
		b.WriteString(r.renderLint(env))
	case "analyze":
		b.WriteString(r.renderAnalysis(env))
	case "check":
		b.WriteString(r.renderCheck(env))
		if env.Issues != nil {
			b.WriteString(r.renderLint(env))
		}
		if env.Analysis != nil {
			b.WriteString(r.renderAnalysis(env))
		}
	case "inspect-gui":
		b.WriteString(r.renderGUIBoundaries(env))
	case "macros":
		b.WriteString(r.renderMacros(env))
	case "list":
		b.WriteString(r.renderList(env))
	case "backup list":
		b.WriteString(r.renderBackupList(env))
	case "rollback":
		b.WriteString(r.renderRollback(env))
	case "session":
		b.WriteString(r.renderSession(env))
	case "status":
		b.WriteString(r.renderStatus(env))
	case "save":
		b.WriteString(r.renderSave(env))
	case "export-image":
		b.WriteString(r.renderExportImage(env))
	case "pack":
		b.WriteString(r.renderExportImage(env))
	case "form export-image":
		b.WriteString(r.renderFormExportImage(env))
	case "form snapshot":
		b.WriteString(r.renderFormSnapshot(env))
	case "form build", "form apply":
		b.WriteString(r.renderFormWrite(env))
	case "edit":
		b.WriteString(r.renderEdit(env))
	case "pull", "push", "attach":
		if env.Issues != nil {
			b.WriteString(r.renderLint(env))
		}
		if env.Analysis != nil {
			b.WriteString(r.renderAnalysis(env))
		}
		if env.Issues != nil || env.Analysis != nil {
			b.WriteString(r.renderLogs(env))
		} else {
			b.WriteString(r.renderWorkbookSource(env))
		}
	case "fmt":
		b.WriteString(r.renderFmt(env))
	case "diff":
		b.WriteString(r.renderDiff(env))
	case "inspect":
		b.WriteString(r.renderInspect(env))
	case "process list":
		b.WriteString(r.renderProcessList(env))
	case "process cleanup":
		b.WriteString(r.renderProcessCleanup(env))
	case "new", "init", "skill install":
		b.WriteString(r.renderCreated(env))
	default:
		b.WriteString(r.renderLogs(env))
	}
	out := strings.TrimRight(b.String(), "\n")
	return out + "\n"
}

func (r renderer) renderGUIBoundaries(env Envelope) string {
	boundaries := listOfObjects(env.GUIBoundaries)
	if env.GUIBoundaries == nil && env.Status == StatusFailed {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	if len(boundaries) == 0 {
		b.WriteString("No GUI boundaries found.\n")
		return b.String()
	}
	b.WriteString(kv("Boundaries", fmt.Sprintf("%d", len(boundaries))))
	for _, boundary := range boundaries {
		loc := stringValue(boundary, "file")
		if n, ok := numberValue(boundary, "line"); ok && n > 0 {
			loc = fmt.Sprintf("%s:%d", loc, int(n))
		}
		fmt.Fprintf(&b, "- %s [%s] %s\n", loc, stringValue(boundary, "kind"), stringValue(boundary, "symbol"))
		if suggestion := stringValue(boundary, "suggestion"); suggestion != "" {
			b.WriteString("  ")
			b.WriteString(suggestion)
			b.WriteString("\n")
		}
	}
	return b.String()
}

func renderFallback(env Envelope) string {
	var b strings.Builder
	if env.Status == StatusOK {
		if len(env.Logs) == 0 {
			return "ok\n"
		}
		for _, line := range env.Logs {
			b.WriteString(line)
			b.WriteString("\n")
		}
		return b.String()
	}
	for _, line := range env.Logs {
		b.WriteString(line)
		b.WriteString("\n")
	}
	if env.Error != nil {
		b.WriteString(env.Error.Message)
		b.WriteString("\n")
		return b.String()
	}
	return "failed\n"
}

type renderer struct {
	color bool
}

func (r renderer) title(env Envelope) string {
	status := env.Status
	if status == "" {
		status = StatusOK
	}
	label := fmt.Sprintf("xlflow %s", env.Command)
	if env.Command == "" {
		label = "xlflow"
	}
	if status == StatusOK {
		return r.style("OK", "42", true) + " " + r.style(label, "", true)
	}
	return r.style("FAILED", "196", true) + " " + r.style(label, "", true)
}

func (r renderer) renderDoctor(env Envelope) string {
	diag := objectMap(env.Diagnostics)
	if len(diag) == 0 {
		return r.renderLogs(env)
	}
	excel := objectMap(diag["excel"])
	host := objectMap(diag["host"])
	windows := objectMap(diag["windows"])
	pathTranslation := objectMap(diag["path_translation"])
	workbook := objectMap(env.Workbook)
	var b strings.Builder
	b.WriteString("\n")
	if boolValue(host, "is_wsl") {
		distro := stringValue(host, "distro")
		if distro == "" {
			distro = "detected"
		}
		b.WriteString(kv("WSL", distro))
	}
	if len(windows) > 0 {
		detail := stringValue(windows, "xlflow_path")
		if version := stringValue(windows, "xlflow_version"); version != "" {
			if detail != "" {
				detail += " "
			}
			detail += "(" + version + ")"
		}
		if detail == "" {
			detail = "Windows xlflow.exe"
		}
		b.WriteString(r.checkLine(boolValue(windows, "xlflow_found"), "Windows xlflow", detail))
		b.WriteString(r.checkLine(boolValue(windows, "bridge_found"), "Windows bridge", "Excel bridge is available"))
	}
	if len(pathTranslation) > 0 {
		detail := stringValue(pathTranslation, "windows_path")
		if detail == "" {
			detail = "Windows-visible project path"
		}
		b.WriteString(r.checkLine(boolValue(pathTranslation, "supported"), "Path translation", detail))
	}
	if selected := stringValue(diag, "selected_bridge"); selected != "" {
		b.WriteString(kv("Selected bridge", selected))
	}
	if requested := stringValue(diag, "requested_bridge"); requested != "" {
		b.WriteString(kv("Requested bridge", requested))
	}
	if _, ok := boolValueOK(diag, "fallback"); ok {
		b.WriteString(kv("Fallback", yesNo(boolValue(diag, "fallback"))))
	}
	if _, ok := boolValueOK(diag, "legacy"); ok {
		role := "primary"
		if boolValue(diag, "legacy") {
			role = "legacy fallback"
		}
		b.WriteString(kv("Bridge role", role))
	}
	b.WriteString(r.checkLine(r.doctorBool(diag, excel, "excel_installed", "com_activation"), "Excel automation", "Excel COM can be created"))
	systemProfileDesktop := objectMap(excel["systemprofile_desktop"])
	if len(systemProfileDesktop) > 0 {
		systemProfileStatus, systemProfileDetail := summarizeSystemProfileDesktop(systemProfileDesktop)
		switch systemProfileStatus {
		case "ok":
			b.WriteString(r.checkLine(true, "Systemprofile Desktop", systemProfileDetail))
		case "warning":
			b.WriteString(r.warnLine("Systemprofile Desktop", systemProfileDetail))
		default:
			b.WriteString(r.checkLine(false, "Systemprofile Desktop", systemProfileDetail))
		}
		if systemProfileStatus == "failed" && (env.Error == nil || env.Error.Code != "systemprofile_desktop_missing") {
			system32Path, syswow64Path := systemProfileDesktopPaths(systemProfileDesktop)
			b.WriteString("\n")
			b.WriteString("systemprofile Desktop directories are missing.\n")
			b.WriteString("Create both directories:\n")
			b.WriteString("- ")
			b.WriteString(system32Path)
			b.WriteString("\n")
			b.WriteString("- ")
			b.WriteString(syswow64Path)
			b.WriteString("\n\n")
			b.WriteString("This is required for Excel COM automation in non-interactive sessions such as SSH, services, or CI.\n")
		}
	}
	if path := stringValue(workbook, "path"); path != "" {
		b.WriteString(r.checkLine(boolValue(diag, "workbook_openable"), "Workbook", path))
	} else {
		b.WriteString(r.skipLine("Workbook", "Not checked; run xlflow doctor --workbook to open the configured workbook"))
	}
	b.WriteString(r.checkLine(r.doctorBool(diag, excel, "vbide_access", ""), "VBIDE access", "VBA project object model is available"))
	if fix := stringValue(diag, "fix"); fix != "" {
		b.WriteString("\n")
		b.WriteString(r.style("Fix:", "214", true))
		b.WriteString(" ")
		b.WriteString(fix)
		b.WriteString("\n")
	}
	return b.String()
}

func summarizeSystemProfileDesktop(systemProfileDesktop map[string]any) (string, string) {
	system32 := objectMap(systemProfileDesktop["system32"])
	syswow64 := objectMap(systemProfileDesktop["syswow64"])
	system32Status := stringValue(system32, "status")
	syswow64Status := stringValue(syswow64, "status")
	switch {
	case system32Status == "exists" && syswow64Status == "exists":
		return "ok", "systemprofile Desktop directories are available"
	case system32Status == "missing" || syswow64Status == "missing":
		return "failed", "systemprofile Desktop directories are missing"
	case system32Status == "access_denied" && syswow64Status == "access_denied":
		return "warning", "systemprofile Desktop paths could not be inspected without elevated permissions"
	case system32Status == "access_denied":
		return "warning", "System32 path could not be inspected without elevated permissions"
	case syswow64Status == "access_denied":
		return "warning", "SysWOW64 path could not be inspected without elevated permissions"
	default:
		return "warning", "systemprofile Desktop readiness could not be fully inspected"
	}
}

func systemProfileDesktopPaths(systemProfileDesktop map[string]any) (string, string) {
	system32 := objectMap(systemProfileDesktop["system32"])
	syswow64 := objectMap(systemProfileDesktop["syswow64"])
	system32Path := stringValue(system32, "path")
	if system32Path == "" {
		system32Path = "C:\\Windows\\System32\\config\\systemprofile\\Desktop"
	}
	syswow64Path := stringValue(syswow64, "path")
	if syswow64Path == "" {
		syswow64Path = "C:\\Windows\\SysWOW64\\config\\systemprofile\\Desktop"
	}
	return system32Path, syswow64Path
}

func (r renderer) doctorBool(diag map[string]any, excel map[string]any, flatKey string, nestedKey string) bool {
	if v, ok := boolValueOK(diag, flatKey); ok {
		return v
	}
	key := nestedKey
	if key == "" {
		key = flatKey
	}
	if v, ok := boolValueOK(excel, key); ok {
		return v
	}
	return false
}

func (r renderer) renderVersion(env Envelope) string {
	version := objectMap(env.Version)
	if len(version) == 0 {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	if value := stringValue(version, "version"); value != "" {
		b.WriteString(kv("Version", value))
	}
	if value := stringValue(version, "commit"); value != "" {
		b.WriteString(kv("Commit", value))
	}
	if value := stringValue(version, "date"); value != "" {
		b.WriteString(kv("Date", value))
	}
	if value := stringValue(version, "executable_path"); value != "" {
		b.WriteString(kv("Executable", value))
	}
	if value := stringValue(version, "go_version"); value != "" {
		b.WriteString(kv("Go", value))
	}
	if value := stringValue(version, "module_path"); value != "" {
		b.WriteString(kv("Module", value))
	}
	settings := listOfObjects(version["build_settings"])
	if len(settings) > 0 {
		b.WriteString("\n")
		b.WriteString(r.style("Build settings", "", true))
		b.WriteString("\n")
		for _, setting := range settings {
			b.WriteString(kv(stringValue(setting, "key"), stringValue(setting, "value")))
		}
	}
	scripts := listOfObjects(version["scripts"])
	if len(scripts) > 0 {
		b.WriteString("\n")
		b.WriteString(r.style("Scripts", "", true))
		b.WriteString("\n")
		for _, script := range scripts {
			label := stringValue(script, "command")
			summary := stringValue(script, "source")
			if path := stringValue(script, "path"); path != "" {
				summary += " (" + path + ")"
			}
			b.WriteString(kv(label, summary))
		}
	}
	features := listOfObjects(version["features"])
	if len(features) > 0 {
		b.WriteString("\n")
		b.WriteString(r.style("Features", "", true))
		b.WriteString("\n")
		for _, feature := range features {
			fmt.Fprintf(&b, "- %s: %s\n", stringValue(feature, "name"), stringValue(feature, "description"))
		}
	}
	return b.String()
}

func (r renderer) checkLine(ok bool, name, detail string) string {
	marker := r.style("[x]", "196", true)
	if ok {
		marker = r.style("[ok]", "42", true)
	}
	return fmt.Sprintf("%s %s - %s\n", marker, r.style(name, "", true), detail)
}

func (r renderer) skipLine(name, detail string) string {
	return fmt.Sprintf("%s %s - %s\n", r.style("[-]", "244", true), r.style(name, "", true), detail)
}

func (r renderer) warnLine(name, detail string) string {
	return fmt.Sprintf("%s %s - %s\n", r.style("[!]", "214", true), r.style(name, "", true), detail)
}

func summarizeRuntime(runtime map[string]any) string {
	if len(runtime) == 0 {
		return ""
	}
	mode := stringValue(runtime, "mode_name")
	if mode == "" {
		mode = stringValue(runtime, "mode")
	}
	if mode == "" {
		return ""
	}
	parts := []string{mode}
	if source := stringValue(runtime, "source"); source != "" {
		parts = append(parts, "source="+source)
	}
	if injected, ok := boolValueOK(runtime, "injected"); ok {
		if injected {
			parts = append(parts, "workbook marker injected")
		} else {
			parts = append(parts, "workbook marker not injected")
		}
	}
	return strings.Join(parts, ", ")
}

func (r renderer) renderRun(env Envelope) string {
	macro := objectMap(env.Macro)
	workbook := objectMap(env.Workbook)
	runtime := objectMap(env.Runtime)
	debugResult := objectMap(env.Debug)
	ui := objectMap(env.UI)
	if len(macro) == 0 && len(workbook) == 0 && len(runtime) == 0 && len(debugResult) == 0 && len(ui) == 0 && env.RunDiagnostic == nil {
		var b strings.Builder
		b.WriteString(r.renderLogs(env))
		b.WriteString(r.renderDebug(env))
		b.WriteString(r.renderWarningsAndHints(env))
		return b.String()
	}
	var b strings.Builder
	b.WriteString("\n")
	if name := stringValue(macro, "name"); name != "" {
		b.WriteString(kv("Macro", name))
	}
	if summary := summarizeRuntime(runtime); summary != "" {
		b.WriteString(kv("Runtime", summary))
	}
	if duration, ok := numberValue(macro, "duration_ms"); ok {
		b.WriteString(kv("Duration", fmt.Sprintf("%dms", int(duration))))
	}
	if path := stringValue(workbook, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	}
	if sessionSummary := summarizeSessionUsage(workbook); sessionSummary != "" {
		b.WriteString(kv("Session", sessionSummary))
	}
	if save := summarizeSaveRequirement(workbook); save != "" {
		b.WriteString(kv("Save", save))
	}
	if summary := summarizeRunWorkbookResult(workbook); summary != "" {
		b.WriteString(kv("Result", summary))
	}
	b.WriteString(r.renderLogs(env))
	b.WriteString(r.renderUI(env))
	b.WriteString(r.renderDebug(env))
	if diag := objectMap(env.RunDiagnostic); len(diag) > 0 {
		b.WriteString("\n")
		b.WriteString(r.style("Diagnostic", "", true))
		b.WriteString("\n")
		if kind := stringValue(diag, "kind"); kind != "" {
			b.WriteString(kv("Kind", kind))
		}
		if messages := stringList(diag["message"]); len(messages) > 0 {
			b.WriteString("Message:\n")
			for _, line := range messages {
				b.WriteString("  ")
				b.WriteString(line)
				b.WriteString("\n")
			}
		} else if message := stringValue(diag, "message"); message != "" {
			b.WriteString(kv("Message", message))
		}
		if loc := objectMap(diag["location"]); len(loc) > 0 {
			parts := []string{}
			for _, key := range []string{"module", "procedure", "file", "source_path"} {
				if v := stringValue(loc, key); v != "" {
					parts = append(parts, v)
				}
			}
			if n, ok := numberValue(loc, "line"); ok && n > 0 {
				parts = append(parts, fmt.Sprintf("line %d", int(n)))
			}
			if n, ok := numberValue(loc, "column"); ok && n > 0 {
				parts = append(parts, fmt.Sprintf("column %d", int(n)))
			}
			if token := stringValue(loc, "token"); token != "" {
				parts = append(parts, token)
			} else if text := stringValue(loc, "text"); text != "" {
				parts = append(parts, text)
			}
			if len(parts) > 0 {
				b.WriteString(kv("Location", strings.Join(parts, " ")))
			}
		}
		if cause := stringValue(diag, "likely_cause"); cause != "" {
			b.WriteString(kv("Likely cause", cause))
		}
		if suggestion := stringValue(diag, "suggestion"); suggestion != "" {
			b.WriteString(kv("Suggestion", suggestion))
		}
		if nearby := stringList(diag["nearby_code"]); len(nearby) > 0 {
			b.WriteString("Nearby code:\n")
			for _, line := range nearby {
				b.WriteString("  ")
				b.WriteString(line)
				b.WriteString("\n")
			}
		}
	}
	b.WriteString(r.renderWarningsAndHints(env))
	return b.String()
}

func (r renderer) renderTest(env Envelope) string {
	tests := listOfObjects(env.Tests)
	if env.Tests == nil {
		var b strings.Builder
		b.WriteString(r.renderLogs(env))
		b.WriteString(r.renderUI(env))
		b.WriteString(r.renderDebug(env))
		b.WriteString(r.renderWarningsAndHints(env))
		return b.String()
	}
	workbook := objectMap(env.Workbook)
	runtime := objectMap(env.Runtime)
	passed := 0
	failed := 0
	inconclusive := 0
	notRun := 0
	for _, test := range tests {
		switch stringValue(test, "status") {
		case "passed":
			passed++
		case "failed":
			failed++
		case "inconclusive":
			inconclusive++
		default:
			notRun++
		}
	}
	var b strings.Builder
	b.WriteString("\n")
	if path := stringValue(workbook, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	}
	if summary := summarizeRuntime(runtime); summary != "" {
		b.WriteString(kv("Runtime", summary))
	}
	if sessionSummary := summarizeSessionUsage(workbook); sessionSummary != "" {
		b.WriteString(kv("Session", sessionSummary))
	}
	if needsSave := summarizeSaveRequirement(workbook); needsSave != "" {
		b.WriteString(kv("Save", needsSave))
	}
	summary := fmt.Sprintf("%d passed, %d failed", passed, failed)
	if inconclusive > 0 {
		summary += fmt.Sprintf(", %d inconclusive", inconclusive)
	}
	if notRun > 0 {
		summary += fmt.Sprintf(", %d not run", notRun)
	}
	summary += fmt.Sprintf(", %d total", len(tests))
	b.WriteString(kv("Summary", summary))
	for _, test := range tests {
		status := stringValue(test, "status")
		marker := r.style("[-]", "241", true)
		switch status {
		case "passed":
			marker = r.style("[ok]", "42", true)
		case "failed":
			marker = r.style("[x]", "196", true)
		case "inconclusive":
			marker = r.style("[?]", "214", true)
		}
		name := stringValue(test, "name")
		module := stringValue(test, "module")
		duration := ""
		if n, ok := numberValue(test, "duration_ms"); ok {
			duration = fmt.Sprintf(" (%dms)", int(n))
		}
		fmt.Fprintf(&b, "%s %s.%s%s\n", marker, module, name, duration)
		if errMap := objectMap(test["error"]); len(errMap) > 0 {
			b.WriteString("  ")
			b.WriteString(stringValue(errMap, "message"))
			b.WriteString("\n")
		}
	}
	b.WriteString(r.renderUI(env))
	b.WriteString(r.renderDebug(env))
	b.WriteString(r.renderWarningsAndHints(env))
	return b.String()
}

func (r renderer) renderDiagnostic(title string, value any) string {
	diag := objectMap(value)
	if len(diag) == 0 {
		return ""
	}
	var b strings.Builder
	b.WriteString("\n")
	b.WriteString(r.style(title, "", true))
	b.WriteString("\n")
	if kind := stringValue(diag, "kind"); kind != "" {
		b.WriteString(kv("Kind", kind))
	}
	if loc := objectMap(diag["location"]); len(loc) > 0 {
		parts := []string{}
		for _, key := range []string{"module", "component", "procedure", "file", "source_path"} {
			if v := stringValue(loc, key); v != "" {
				parts = append(parts, v)
			}
		}
		if n, ok := numberValue(loc, "line"); ok && n > 0 {
			parts = append(parts, fmt.Sprintf("line %d", int(n)))
		}
		if n, ok := numberValue(loc, "column"); ok && n > 0 {
			parts = append(parts, fmt.Sprintf("column %d", int(n)))
		}
		if text := stringValue(loc, "text"); text != "" {
			parts = append(parts, text)
		}
		if len(parts) > 0 {
			b.WriteString(kv("Location", strings.Join(parts, " ")))
		}
	}
	return b.String()
}

func (r renderer) renderDebug(env Envelope) string {
	debugResult := objectMap(env.Debug)
	events := listOfObjects(debugResult["events"])
	count := len(events)
	if total, ok := numberValue(debugResult, "count"); ok && int(total) > count {
		count = int(total)
	}
	if count == 0 {
		return ""
	}
	var b strings.Builder
	b.WriteString("\n")
	b.WriteString(r.style("Debug", "", true))
	b.WriteString("\n")
	b.WriteString(kv("Events", fmt.Sprintf("%d", count)))
	if boolValue(debugResult, "truncated") {
		b.WriteString(kv("Retention", "truncated to recent events"))
	}
	for _, event := range events {
		b.WriteString("- ")
		b.WriteString(summarizeDebugEvent(event))
		b.WriteString("\n")
	}
	return b.String()
}

func summarizeDebugEvent(event map[string]any) string {
	parts := []string{"log"}
	if message := stringValue(event, "message"); message != "" {
		parts = append(parts, "message="+message)
	}
	if source := stringValue(event, "source"); source != "" {
		parts = append(parts, "source="+source)
	}
	if mode := stringValue(event, "runtime_mode"); mode != "" {
		parts = append(parts, "mode="+mode)
	}
	if errText := stringValue(event, "error"); errText != "" {
		parts = append(parts, "error="+errText)
	}
	return strings.Join(parts, " ")
}

func (r renderer) renderUI(env Envelope) string {
	ui := objectMap(env.UI)
	events := listOfObjects(ui["events"])
	if len(events) == 0 {
		return ""
	}
	var b strings.Builder
	b.WriteString("\n")
	b.WriteString(r.style("UI", "", true))
	b.WriteString("\n")
	b.WriteString(kv("Events", fmt.Sprintf("%d", len(events))))
	for _, event := range events {
		b.WriteString("- ")
		b.WriteString(summarizeUIEvent(event))
		b.WriteString("\n")
	}
	return b.String()
}

func summarizeUIEvent(event map[string]any) string {
	kind := stringValue(event, "kind")
	if kind == "" {
		kind = "dialog"
	}
	id := stringValue(event, "dialog_id")
	parts := []string{kind}
	if id != "" {
		parts = append(parts, "id="+id)
	}
	if source := stringValue(event, "response_source"); source != "" {
		parts = append(parts, "source="+source)
	}
	if result := stringValue(event, "resolved_result"); result != "" {
		parts = append(parts, "result="+result)
	}
	if value := stringValue(event, "resolved_value"); value != "" {
		if boolValue(event, "redacted") {
			value = "[redacted]"
		}
		parts = append(parts, "value="+value)
	}
	if errText := stringValue(event, "error"); errText != "" {
		parts = append(parts, "error="+errText)
	}
	return strings.Join(parts, " ")
}

func (r renderer) renderLint(env Envelope) string {
	issues := listOfObjects(env.Issues)
	if env.Issues == nil && env.Status == StatusFailed {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	if len(issues) == 0 {
		b.WriteString("No lint issues found.\n")
		b.WriteString(r.renderWarningsAndHints(env))
		return b.String()
	}
	b.WriteString(kv("Issues", fmt.Sprintf("%d", len(issues))))
	for _, issue := range issues {
		loc := stringValue(issue, "file")
		if n, ok := numberValue(issue, "line"); ok && n > 0 {
			loc = fmt.Sprintf("%s:%d", loc, int(n))
		}
		fmt.Fprintf(&b, "%s %s %s - %s\n", r.style("["+stringValue(issue, "severity")+"]", "214", true), stringValue(issue, "code"), loc, stringValue(issue, "message"))
	}
	b.WriteString(r.renderWarningsAndHints(env))
	return b.String()
}

func (r renderer) renderAnalysis(env Envelope) string {
	findings := listOfObjects(env.Analysis)
	if env.Analysis == nil && env.Status == StatusFailed {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	if len(findings) == 0 {
		b.WriteString("No analysis findings found.\n")
		b.WriteString(r.renderWarningsAndHints(env))
		return b.String()
	}
	b.WriteString(kv("Findings", fmt.Sprintf("%d", len(findings))))
	for _, finding := range findings {
		loc := stringValue(finding, "file")
		if n, ok := numberValue(finding, "line"); ok && n > 0 {
			loc = fmt.Sprintf("%s:%d", loc, int(n))
		}
		fmt.Fprintf(&b, "%s %s %s - %s\n", r.style("["+stringValue(finding, "severity")+"]", "214", true), stringValue(finding, "code"), loc, stringValue(finding, "message"))
		if suggestion := stringValue(finding, "suggestion"); suggestion != "" {
			b.WriteString("  Suggestion: ")
			b.WriteString(suggestion)
			b.WriteString("\n")
		}
	}
	b.WriteString(r.renderWarningsAndHints(env))
	return b.String()
}

func (r renderer) renderCheck(env Envelope) string {
	check := objectMap(env.Check)
	if len(check) == 0 {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	for _, name := range []string{"lint", "analyze", "doctor"} {
		item := objectMap(check[name])
		if len(item) == 0 {
			continue
		}
		status := stringValue(item, "status")
		if status == "" {
			status = "ok"
		}
		count := ""
		if n, ok := numberValue(item, "count"); ok {
			count = fmt.Sprintf(" (%d)", int(n))
		}
		b.WriteString(kv(name, status+count))
	}
	b.WriteString(r.renderLogs(env))
	return b.String()
}

func (r renderer) renderMacros(env Envelope) string {
	macros := listOfObjects(env.Macros)
	if env.Macros == nil && env.Status == StatusFailed {
		return r.renderLogs(env)
	}
	workbook := objectMap(env.Workbook)
	var b strings.Builder
	b.WriteString("\n")
	if path := stringValue(workbook, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	}
	if sessionSummary := summarizeSessionUsage(workbook); sessionSummary != "" {
		b.WriteString(kv("Session", sessionSummary))
	}
	b.WriteString(kv("Entrypoints", fmt.Sprintf("%d", len(macros))))
	for _, macro := range macros {
		name := stringValue(macro, "qualified_name")
		if name == "" {
			name = strings.Trim(strings.Join([]string{stringValue(macro, "module"), stringValue(macro, "name")}, "."), ".")
		}
		args := strings.Join(stringList(macro["args"]), ", ")
		if args != "" {
			args = "(" + args + ")"
		}
		kind := stringValue(macro, "kind")
		if kind != "" {
			kind = " [" + kind + "]"
		}
		fmt.Fprintf(&b, "- %s%s%s\n", name, args, kind)
		runnable := boolValue(macro, "runnable")
		if runnable {
			b.WriteString("    Runnable: yes\n")
			if cmd := stringValue(macro, "run_command"); cmd != "" {
				fmt.Fprintf(&b, "    Run: %s\n", cmd)
			}
		} else {
			reason := stringValue(macro, "reason_not_runnable")
			if reason == "" {
				reason = "unknown"
			}
			fmt.Fprintf(&b, "    Runnable: no\n    Reason: %s\n", reason)
		}
	}
	if entry, ok := env.DefaultEntry.(string); ok && entry != "" {
		b.WriteString("\n")
		fmt.Fprintf(&b, "Default entry:\n  %s\n", entry)
	}
	suggestions := listOfObjects(env.Suggestions)
	if len(suggestions) > 0 {
		b.WriteString("\nNext:\n")
		for _, s := range suggestions {
			title := stringValue(s, "title")
			cmd := stringValue(s, "command")
			if title != "" {
				fmt.Fprintf(&b, "  %s:\n", title)
			}
			if cmd != "" {
				fmt.Fprintf(&b, "    %s\n", cmd)
			}
		}
	}
	b.WriteString(r.renderWarningsAndHints(env))
	return b.String()
}

func (r renderer) renderList(env Envelope) string {
	forms := listOfObjects(env.Forms)
	workbook := objectMap(env.Workbook)
	if env.Forms == nil && len(workbook) == 0 && env.Warnings == nil && env.Hints == nil {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	if path := stringValue(workbook, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	}
	if sessionSummary := summarizeSessionUsage(workbook); sessionSummary != "" {
		b.WriteString(kv("Session", sessionSummary))
	}
	if save := summarizeSaveRequirement(workbook); save != "" {
		b.WriteString(kv("Save", save))
	}
	if env.Forms == nil {
		b.WriteString(kv("Forms", "unavailable"))
	} else {
		b.WriteString(kv("Forms", fmt.Sprintf("%d", len(forms))))
	}
	for _, form := range forms {
		name := stringValue(form, "name")
		if name == "" {
			continue
		}
		details := []string{}
		if value, ok := boolValueOK(form, "has_frx"); ok && value {
			details = append(details, "has .frx")
		}
		if path := stringValue(form, "source_path"); path != "" {
			details = append(details, path)
		}
		if len(details) == 0 {
			fmt.Fprintf(&b, "- %s\n", name)
			continue
		}
		fmt.Fprintf(&b, "- %s (%s)\n", name, strings.Join(details, ", "))
	}
	b.WriteString(r.renderWarningsAndHints(env))
	b.WriteString(r.renderLogs(env))
	return b.String()
}

func (r renderer) renderBackupList(env Envelope) string {
	backups := listOfObjects(env.Backups)
	if env.Backups == nil {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	b.WriteString(kv("Backups", fmt.Sprintf("%d", len(backups))))
	for _, item := range backups {
		line := []string{stringValue(item, "id")}
		if created := stringValue(item, "created_at"); created != "" {
			line = append(line, created)
		}
		if reason := stringValue(item, "reason"); reason != "" {
			line = append(line, reason)
		}
		if workbook := stringValue(item, "workbook"); workbook != "" {
			line = append(line, workbook)
		}
		if path := stringValue(item, "path"); path != "" {
			line = append(line, path)
		}
		b.WriteString("- ")
		b.WriteString(strings.Join(line, " | "))
		b.WriteString("\n")
	}
	b.WriteString(r.renderWarningsAndHints(env))
	b.WriteString(r.renderLogs(env))
	return b.String()
}

func (r renderer) renderRollback(env Envelope) string {
	rollback := objectMap(env.Rollback)
	target := objectMap(env.Target)
	if len(rollback) == 0 && len(target) == 0 {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	if summary := summarizeTarget(target); summary != "" {
		b.WriteString(kv("Target", summary))
	}
	restored := objectMap(rollback["restored_from"])
	if id := stringValue(restored, "id"); id != "" {
		b.WriteString(kv("Backup ID", id))
	}
	if path := stringValue(restored, "path"); path != "" {
		b.WriteString(kv("Restored from", path))
	}
	if reason := stringValue(restored, "reason"); reason != "" {
		b.WriteString(kv("Reason", reason))
	}
	if created := stringValue(restored, "created_at"); created != "" {
		b.WriteString(kv("Created", created))
	}
	safety := objectMap(rollback["safety_backup"])
	if id := stringValue(safety, "id"); id != "" {
		b.WriteString(kv("Safety backup", id))
	}
	if path := stringValue(safety, "path"); path != "" {
		b.WriteString(kv("Safety path", path))
	}
	b.WriteString(r.renderWarningsAndHints(env))
	b.WriteString(r.renderLogs(env))
	return b.String()
}

func (r renderer) renderWorkbookSource(env Envelope) string {
	workbook := objectMap(env.Workbook)
	backup := objectMap(env.Backup)
	source := objectMap(env.Source)
	if len(workbook) == 0 && len(backup) == 0 && len(source) == 0 && env.PushDiagnostic == nil {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	if path := stringValue(workbook, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	}
	if path := stringValue(backup, "path"); path != "" {
		b.WriteString(kv("Backup", path))
	}
	if path := stringValue(source, "path"); path != "" {
		b.WriteString(kv("Source", path))
	}
	if sessionSummary := summarizeSessionUsage(workbook); sessionSummary != "" {
		b.WriteString(kv("Session", sessionSummary))
	}
	if summary := summarizeWorkbookSourceResult(env.Command, workbook, source); summary != "" {
		b.WriteString(kv("Result", summary))
	}
	if save := summarizeSaveRequirement(workbook); save != "" {
		b.WriteString(kv("Save", save))
	}
	if updated, ok := boolValueOK(source, "updated"); ok {
		b.WriteString(kv("Source updated", fmt.Sprintf("%t", updated)))
	}
	b.WriteString(r.renderDiagnostic("Diagnostic", env.PushDiagnostic))
	b.WriteString(r.renderWarningsAndHints(env))
	b.WriteString(r.renderLogs(env))
	return b.String()
}

func (r renderer) renderExportImage(env Envelope) string {
	workbook := objectMap(env.Workbook)
	target := objectMap(env.Target)
	outputPayload := objectMap(env.Output)
	if len(workbook) == 0 && len(target) == 0 && len(outputPayload) == 0 && env.Warnings == nil && env.Hints == nil {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	if path := stringValue(workbook, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	} else if path := stringValue(target, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	}
	if summary := summarizeExportImageTarget(target); summary != "" {
		b.WriteString(kv("Export target", summary))
	}
	if sessionSummary := summarizeSessionUsage(workbook); sessionSummary != "" {
		b.WriteString(kv("Session", sessionSummary))
	}
	if sheet := stringValue(target, "sheet"); sheet != "" {
		b.WriteString(kv("Sheet", sheet))
	}
	if cellRange := stringValue(target, "range"); cellRange != "" {
		b.WriteString(kv("Selection", cellRange))
	}
	if save := summarizeSaveRequirement(workbook); save != "" {
		b.WriteString(kv("Save", save))
	}
	if path := stringValue(outputPayload, "path"); path != "" {
		b.WriteString(kv("Output", path))
	}
	if format := stringValue(outputPayload, "format"); format != "" {
		b.WriteString(kv("Format", strings.ToUpper(format)))
	}
	if width, ok := numberValue(outputPayload, "width_px"); ok {
		if height, ok := numberValue(outputPayload, "height_px"); ok {
			b.WriteString(kv("Size", fmt.Sprintf("%d x %d px", int(width), int(height))))
		}
	}
	if value, ok := boolValueOK(outputPayload, "default"); ok {
		b.WriteString(kv("Default path", fmt.Sprintf("%t", value)))
	}
	if value, ok := boolValueOK(outputPayload, "created_parent_dirs"); ok {
		b.WriteString(kv("Created dirs", fmt.Sprintf("%t", value)))
	}
	b.WriteString(r.renderWarningsAndHints(env))
	b.WriteString(r.renderLogs(env))
	return b.String()
}

func (r renderer) renderFormSnapshot(env Envelope) string {
	workbook := objectMap(env.Workbook)
	target := objectMap(env.Target)
	form := objectMap(env.Forms)
	outputPayload := objectMap(env.Output)
	if len(workbook) == 0 && len(target) == 0 && len(form) == 0 && len(outputPayload) == 0 && env.Warnings == nil && env.Hints == nil {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	if path := stringValue(workbook, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	} else if path := stringValue(target, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	}
	if summary := summarizeTarget(target); summary != "" {
		b.WriteString(kv("Snapshot target", summary))
	}
	if sessionSummary := summarizeSessionUsage(workbook); sessionSummary != "" {
		b.WriteString(kv("Session", sessionSummary))
	}
	if name := stringValue(form, "name"); name != "" {
		b.WriteString(kv("Form", name))
	}
	if basis := stringValue(form, "basis"); basis != "" {
		b.WriteString(kv("Basis", basis))
	}
	if caption := stringValue(form, "caption"); caption != "" {
		b.WriteString(kv("Caption", caption))
	}
	if coord := stringValue(form, "coordinate_system"); coord != "" {
		b.WriteString(kv("Coordinates", coord))
	}
	if count, ok := numberValue(form, "control_count"); ok {
		b.WriteString(kv("Controls", fmt.Sprintf("%d", int(count))))
	}
	if save := summarizeSaveRequirement(workbook); save != "" {
		b.WriteString(kv("Save", save))
	}
	if path := stringValue(outputPayload, "path"); path != "" {
		b.WriteString(kv("Output", path))
	}
	if format := stringValue(outputPayload, "format"); format != "" {
		b.WriteString(kv("Format", strings.ToUpper(format)))
	}
	b.WriteString(r.renderWarningsAndHints(env))
	b.WriteString(r.renderLogs(env))
	return b.String()
}

func (r renderer) renderFormExportImage(env Envelope) string {
	workbook := objectMap(env.Workbook)
	target := objectMap(env.Target)
	form := objectMap(env.Forms)
	outputPayload := objectMap(env.Output)
	if len(workbook) == 0 && len(target) == 0 && len(form) == 0 && len(outputPayload) == 0 && env.Warnings == nil && env.Hints == nil {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	if path := stringValue(workbook, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	} else if path := stringValue(target, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	}
	if summary := summarizeTarget(target); summary != "" {
		b.WriteString(kv("Export target", summary))
	}
	if sessionSummary := summarizeSessionUsage(workbook); sessionSummary != "" {
		b.WriteString(kv("Session", sessionSummary))
	}
	if name := stringValue(form, "name"); name != "" {
		b.WriteString(kv("Form", name))
	} else if name := stringValue(target, "form"); name != "" {
		b.WriteString(kv("Form", name))
	}
	if basis := stringValue(form, "basis"); basis != "" {
		b.WriteString(kv("Basis", basis))
	}
	if initializer := stringValue(form, "initializer"); initializer != "" {
		b.WriteString(kv("Initializer", initializer))
	}
	if captureState := stringValue(target, "capture_state"); captureState != "" {
		b.WriteString(kv("Capture", captureState))
	}
	if save := summarizeSaveRequirement(workbook); save != "" {
		b.WriteString(kv("Save", save))
	}
	if path := stringValue(outputPayload, "path"); path != "" {
		b.WriteString(kv("Output", path))
	}
	if format := stringValue(outputPayload, "format"); format != "" {
		b.WriteString(kv("Format", strings.ToUpper(format)))
	}
	if width, ok := numberValue(outputPayload, "width_px"); ok {
		if height, ok := numberValue(outputPayload, "height_px"); ok {
			b.WriteString(kv("Size", fmt.Sprintf("%d x %d px", int(width), int(height))))
		}
	}
	b.WriteString(r.renderWarningsAndHints(env))
	b.WriteString(r.renderLogs(env))
	return b.String()
}

func (r renderer) renderFormWrite(env Envelope) string {
	workbook := objectMap(env.Workbook)
	target := objectMap(env.Target)
	form := objectMap(env.Forms)
	spec := objectMap(env.Spec)
	if len(workbook) == 0 && len(target) == 0 && len(form) == 0 && len(spec) == 0 && env.Warnings == nil && env.Hints == nil {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	if path := stringValue(workbook, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	} else if path := stringValue(target, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	}
	if summary := summarizeTarget(target); summary != "" {
		b.WriteString(kv("Write target", summary))
	}
	if sessionSummary := summarizeSessionUsage(workbook); sessionSummary != "" {
		b.WriteString(kv("Session", sessionSummary))
	}
	if action := stringValue(form, "action"); action != "" {
		b.WriteString(kv("Action", action))
	}
	if name := stringValue(form, "name"); name != "" {
		b.WriteString(kv("Form", name))
	}
	if basis := stringValue(form, "basis"); basis != "" {
		b.WriteString(kv("Basis", basis))
	}
	if coord := stringValue(form, "coordinate_system"); coord != "" {
		b.WriteString(kv("Coordinates", coord))
	}
	if count, ok := numberValue(form, "control_count"); ok {
		b.WriteString(kv("Controls", fmt.Sprintf("%d", int(count))))
	}
	if specPath := stringValue(form, "spec_path"); specPath != "" {
		b.WriteString(kv("Spec", specPath))
	}
	if overwrite, ok := boolValueOK(form, "overwrite"); ok {
		b.WriteString(kv("Overwrite", fmt.Sprintf("%t", overwrite)))
	}
	if len(spec) > 0 && stringValue(form, "spec_path") == "" {
		b.WriteString(renderSpecMetadata(spec))
	}
	if save := summarizeSaveRequirement(workbook); save != "" {
		b.WriteString(kv("Save", save))
	}
	b.WriteString(r.renderWarningsAndHints(env))
	if suggestion := stringValue(spec, "suggestion"); suggestion != "" {
		b.WriteString(kv("Remediation", suggestion))
	}
	b.WriteString(r.renderLogs(env))
	return b.String()
}

func (r renderer) renderEdit(env Envelope) string {
	workbook := objectMap(env.Workbook)
	target := objectMap(env.Target)
	edit := objectMap(env.Edit)
	if len(workbook) == 0 && len(target) == 0 && len(edit) == 0 && env.Warnings == nil && env.Hints == nil {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	if path := stringValue(workbook, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	} else if path := stringValue(target, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	}
	if summary := summarizeTarget(target); summary != "" {
		b.WriteString(kv("Edit target", summary))
	}
	if sessionSummary := summarizeSessionUsage(workbook); sessionSummary != "" {
		b.WriteString(kv("Session", sessionSummary))
	}
	if selector := summarizeEditSelector(edit); selector != "" {
		b.WriteString(kv("Selection", selector))
	}
	if summary := summarizeEditMutation(edit); summary != "" {
		b.WriteString(kv("Mutation", summary))
	}
	if events := summarizeEditEvents(edit); events != "" {
		b.WriteString(kv("Events", events))
	}
	if save := summarizeSaveRequirement(workbook); save != "" {
		b.WriteString(kv("Save", save))
	}
	b.WriteString(r.renderWarningsAndHints(env))
	b.WriteString(r.renderLogs(env))
	return b.String()
}

func (r renderer) renderProcessList(env Envelope) string {
	processes := listOfObjects(env.Process)
	if env.Process == nil && env.Status == StatusFailed {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	if env.Process == nil {
		b.WriteString(kv("Processes", "unavailable"))
		return b.String()
	}
	b.WriteString(kv("Processes", fmt.Sprintf("%d", len(processes))))
	for _, proc := range processes {
		pid := intNumber(proc, "pid")
		hasWorkbook, workbookKnown := boolValueOK(proc, "has_workbook")
		workbookLabel := "unknown"
		if workbookKnown {
			if hasWorkbook {
				workbookLabel = "has workbook"
			} else {
				workbookLabel = "no workbook"
			}
		}
		fmt.Fprintf(&b, "- PID %d %s\n", pid, workbookLabel)
	}
	if len(processes) == 0 {
		b.WriteString("0 processes found.\n")
	}
	b.WriteString(r.renderLogs(env))
	return b.String()
}

func (r renderer) renderProcessCleanup(env Envelope) string {
	if env.Process == nil && env.Status == StatusFailed {
		return r.renderLogs(env)
	}
	payload := objectMap(env.Process)
	var b strings.Builder
	b.WriteString("\n")
	if env.Process == nil {
		b.WriteString("Process cleanup result unavailable.\n")
		return b.String()
	}
	mode := stringValue(payload, "mode")
	if mode == "" {
		mode = "unknown"
	}
	total := intNumber(payload, "total")
	results := listOfObjects(payload["results"])
	terminatedCount := 0
	failedCount := 0
	for _, res := range results {
		if boolValue(res, "terminated") {
			terminatedCount++
		} else {
			failedCount++
		}
	}
	b.WriteString(kv("Mode", mode))
	b.WriteString(kv("Total", fmt.Sprintf("%d", total)))
	for _, res := range results {
		pid := intNumber(res, "pid")
		method := stringValue(res, "method")
		terminated := boolValue(res, "terminated")
		statusLabel := "terminated"
		if !terminated {
			statusLabel = "failed"
		}
		fmt.Fprintf(&b, "- PID: %d %s (%s)\n", pid, statusLabel, method)
	}
	fmt.Fprintf(&b, "%d terminated, %d failed\n", terminatedCount, failedCount)
	b.WriteString(r.renderLogs(env))
	return b.String()
}

func (r renderer) renderFmt(env Envelope) string {
	output := objectMap(env.Output)
	warnings := listOfObjects(env.Warnings)
	hints := listOfObjects(env.Hints)
	if len(output) == 0 && len(warnings) == 0 && len(hints) == 0 {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")

	changed, _ := numberValue(output, "changed")
	unchanged, _ := numberValue(output, "unchanged")
	skipped, _ := numberValue(output, "skipped")
	total, _ := numberValue(output, "total")
	mode := stringValue(output, "mode")

	summaryParts := []string{}
	if changed > 0 {
		summaryParts = append(summaryParts, fmt.Sprintf("%d %s", int(changed), fmtSummaryLabel(mode, int(changed))))
	}
	if unchanged > 0 {
		summaryParts = append(summaryParts, fmt.Sprintf("%d unchanged", int(unchanged)))
	}
	if skipped > 0 {
		summaryParts = append(summaryParts, fmt.Sprintf("%d skipped", int(skipped)))
	}
	if total > 0 {
		summaryParts = append(summaryParts, fmt.Sprintf("%d total", int(total)))
	}
	if len(summaryParts) > 0 {
		b.WriteString(kv("Summary", strings.Join(summaryParts, ", ")))
	}

	if paths := stringList(output["changed_paths"]); len(paths) > 0 {
		b.WriteString("Changed:\n")
		for _, p := range paths {
			fmt.Fprintf(&b, "- %s\n", p)
		}
	}
	if paths := stringList(output["skipped_paths"]); len(paths) > 0 {
		b.WriteString("Skipped:\n")
		for _, p := range paths {
			fmt.Fprintf(&b, "- %s\n", p)
		}
	}

	b.WriteString(r.renderWarningsAndHints(env))
	b.WriteString(r.renderLogs(env))
	return b.String()
}

func fmtSummaryLabel(mode string, count int) string {
	switch mode {
	case "write":
		return pluralize(count, "formatted")
	case "check":
		return pluralize(count, "not formatted")
	case "diff":
		return pluralize(count, "would be reformatted")
	case "inspect", "":
		return pluralize(count, "would be formatted")
	default:
		return pluralize(count, "changed")
	}
}

func pluralize(count int, label string) string {
	return label
}

func (r renderer) renderDiff(env Envelope) string {
	diff := objectMap(env.Diff)
	if len(diff) == 0 {
		return r.renderLogs(env)
	}
	summary := objectMap(diff["summary"])
	var b strings.Builder
	b.WriteString("\n")
	total, _ := numberValue(summary, "total_diffs")
	b.WriteString(kv("Total diffs", fmt.Sprintf("%d", int(total))))
	for _, key := range []string{"sheet_diffs", "cell_diffs", "vba_diffs"} {
		if value, ok := numberValue(summary, key); ok {
			b.WriteString(kv(labelFromKey(key), fmt.Sprintf("%d", int(value))))
		}
	}
	b.WriteString(r.renderLogs(env))
	return b.String()
}

func (r renderer) renderInspectStandalone(env Envelope) string {
	if env.Command != "inspect" || env.Status != StatusOK {
		return ""
	}
	payload := objectMap(env.Inspect)
	if len(payload) == 0 {
		return ""
	}
	switch stringValue(payload, "format") {
	case "json":
		standalone := map[string]any{}
		for key, value := range payload {
			standalone[key] = value
		}
		if env.Target != nil {
			standalone["target_state"] = env.Target
		}
		if env.Session != nil {
			standalone["session"] = env.Session
		}
		if env.Warnings != nil {
			standalone["warnings"] = env.Warnings
		}
		if env.Hints != nil {
			standalone["hints"] = env.Hints
		}
		data, err := json.MarshalIndent(standalone, "", "  ")
		if err != nil {
			return ""
		}
		return string(data) + "\n"
	case "markdown":
		text := strings.TrimRight(r.renderInspectMarkdown(env, payload), "\n")
		if text == "" {
			return ""
		}
		return text + "\n"
	default:
		return ""
	}
}

func (r renderer) renderInspect(env Envelope) string {
	payload := objectMap(env.Inspect)
	if len(payload) == 0 {
		return r.renderLogs(env)
	}
	switch stringValue(payload, "target") {
	case "workbook":
		return r.renderInspectWorkbook(env, payload)
	case "sheets":
		return r.renderInspectSheets(env, payload)
	case "form":
		return r.renderInspectForm(env, payload)
	case "range", "used-range":
		return r.renderInspectRange(env, payload)
	case "cell":
		return r.renderInspectCell(env, payload)
	case "calls":
		return r.renderInspectCalls(env, payload)
	case "symbols":
		return r.renderInspectSymbols(env, payload)
	default:
		return r.renderLogs(env)
	}
}

func (r renderer) renderInspectWorkbook(env Envelope, payload map[string]any) string {
	workbook := objectMap(payload["workbook"])
	if len(workbook) == 0 {
		return ""
	}
	var b strings.Builder
	b.WriteString("\n")
	b.WriteString(r.renderInspectTargetSession(env))
	b.WriteString(renderInspectTargetInfo(payload))
	if path := stringValue(workbook, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	}
	if name := stringValue(workbook, "name"); name != "" {
		b.WriteString(kv("Name", name))
	}
	if active := stringValue(workbook, "active_sheet"); active != "" {
		b.WriteString(kv("Active sheet", active))
	}
	sheets := listOfObjects(workbook["sheets"])
	b.WriteString(kv("Sheets", fmt.Sprintf("%d", len(sheets))))
	for _, sheet := range sheets {
		fmt.Fprintf(
			&b,
			"- %d. %s (%s, used %s, %d row(s) x %d column(s))\n",
			intNumber(sheet, "index"),
			stringValue(sheet, "name"),
			visibleLabel(sheet),
			emptyDash(stringValue(sheet, "used_range")),
			intNumber(sheet, "row_count"),
			intNumber(sheet, "column_count"),
		)
	}
	b.WriteString(r.renderWarningsAndHints(env))
	return b.String()
}

func (r renderer) renderInspectSheets(env Envelope, payload map[string]any) string {
	sheets := listOfObjects(payload["sheets"])
	var b strings.Builder
	b.WriteString("\n")
	b.WriteString(r.renderInspectTargetSession(env))
	b.WriteString(renderInspectTargetInfo(payload))
	b.WriteString(kv("Sheets", fmt.Sprintf("%d", len(sheets))))
	for _, sheet := range sheets {
		fmt.Fprintf(
			&b,
			"- %d. %s (%s, used %s, %d row(s) x %d column(s))\n",
			intNumber(sheet, "index"),
			stringValue(sheet, "name"),
			visibleLabel(sheet),
			emptyDash(stringValue(sheet, "used_range")),
			intNumber(sheet, "row_count"),
			intNumber(sheet, "column_count"),
		)
	}
	b.WriteString(r.renderWarningsAndHints(env))
	return b.String()
}

func (r renderer) renderInspectRange(env Envelope, payload map[string]any) string {
	snapshot := objectMap(payload["range"])
	if len(snapshot) == 0 {
		return ""
	}
	var b strings.Builder
	b.WriteString("\n")
	b.WriteString(r.renderInspectTargetSession(env))
	b.WriteString(renderInspectTargetInfo(payload))
	if sheet := stringValue(snapshot, "sheet"); sheet != "" {
		b.WriteString(kv("Sheet", sheet))
	}
	if value := stringValue(snapshot, "range"); value != "" {
		b.WriteString(kv("Range", value))
	}
	if value := stringValue(snapshot, "used_range"); value != "" {
		b.WriteString(kv("Used range", value))
	}
	if value := stringValue(snapshot, "returned_range"); value != "" {
		b.WriteString(kv("Returned", value))
	}
	b.WriteString(kv("Size", fmt.Sprintf("%d row(s) x %d column(s)", intNumber(snapshot, "row_count"), intNumber(snapshot, "column_count"))))
	if boolValue(snapshot, "truncated") {
		b.WriteString(kv("Truncated", "true"))
	}
	if boolValue(snapshot, "style_included") {
		b.WriteString(kv("Style", "included"))
	}
	for _, warning := range stringList(snapshot["warnings"]) {
		b.WriteString("! ")
		b.WriteString(warning)
		b.WriteString("\n")
	}
	values := matrixStrings(snapshot["values"])
	if len(values) == 0 {
		b.WriteString("Values: <empty>\n")
		b.WriteString(r.renderWarningsAndHints(env))
		return b.String()
	}
	b.WriteString("Values:\n")
	for _, row := range values {
		b.WriteString("  ")
		b.WriteString(strings.Join(row, " | "))
		b.WriteString("\n")
	}
	b.WriteString(r.renderWarningsAndHints(env))
	return b.String()
}

func (r renderer) renderInspectCell(env Envelope, payload map[string]any) string {
	cell := objectMap(payload["cell"])
	if len(cell) == 0 {
		return ""
	}
	var b strings.Builder
	b.WriteString("\n")
	b.WriteString(r.renderInspectTargetSession(env))
	b.WriteString(renderInspectTargetInfo(payload))
	if sheet := stringValue(cell, "sheet"); sheet != "" {
		b.WriteString(kv("Sheet", sheet))
	}
	if address := stringValue(cell, "address"); address != "" {
		b.WriteString(kv("Cell", address))
	}
	b.WriteString(kv("Value", emptyDash(stringValue(cell, "value"))))
	b.WriteString(r.renderWarningsAndHints(env))
	return b.String()
}

func (r renderer) renderInspectForm(env Envelope, payload map[string]any) string {
	var b strings.Builder
	b.WriteString("\n")
	b.WriteString(r.renderInspectTargetSession(env))
	b.WriteString(renderInspectTargetInfo(payload))
	if form := objectMap(payload["form"]); len(form) > 0 {
		b.WriteString(renderInspectFormSnapshot(form, ""))
		b.WriteString(r.renderWarningsAndHints(env))
		return b.String()
	}
	forms := objectMap(payload["forms"])
	if len(forms) == 0 {
		b.WriteString(kv("Forms", "unavailable"))
		b.WriteString(r.renderWarningsAndHints(env))
		return b.String()
	}
	for _, basis := range []string{"runtime", "designer"} {
		snapshot := objectMap(forms[basis])
		if len(snapshot) == 0 {
			continue
		}
		if b.Len() > 1 {
			b.WriteString("\n")
		}
		b.WriteString(renderInspectFormSnapshot(snapshot, basis))
	}
	b.WriteString(r.renderWarningsAndHints(env))
	return b.String()
}

func (r renderer) renderInspectMarkdown(env Envelope, payload map[string]any) string {
	switch stringValue(payload, "target") {
	case "workbook":
		workbook := objectMap(payload["workbook"])
		if len(workbook) == 0 {
			return ""
		}
		var b strings.Builder
		b.WriteString(renderInspectTargetSessionMarkdown(env))
		if info := renderInspectTargetInfoMarkdown(payload); info != "" {
			b.WriteString(info)
		}
		if path := stringValue(workbook, "path"); path != "" {
			b.WriteString("Workbook: ")
			b.WriteString(path)
			b.WriteString("\n")
		}
		if active := stringValue(workbook, "active_sheet"); active != "" {
			b.WriteString("Active sheet: ")
			b.WriteString(active)
			b.WriteString("\n\n")
		} else {
			b.WriteString("\n")
		}
		b.WriteString(markdownSheetTable(listOfObjects(workbook["sheets"])))
		b.WriteString(renderWarningsAndHintsMarkdown(env))
		return b.String()
	case "sheets":
		return renderInspectTargetSessionMarkdown(env) + renderInspectTargetInfoMarkdown(payload) + markdownSheetTable(listOfObjects(payload["sheets"])) + renderWarningsAndHintsMarkdown(env)
	case "range", "used-range":
		snapshot := objectMap(payload["range"])
		if len(snapshot) == 0 {
			return ""
		}
		var b strings.Builder
		b.WriteString(renderInspectTargetSessionMarkdown(env))
		if info := renderInspectTargetInfoMarkdown(payload); info != "" {
			b.WriteString(info)
		}
		if sheet := stringValue(snapshot, "sheet"); sheet != "" {
			b.WriteString("Sheet: ")
			b.WriteString(sheet)
			b.WriteString("\n")
		}
		if value := stringValue(snapshot, "range"); value != "" {
			b.WriteString("Range: ")
			b.WriteString(value)
			b.WriteString("\n")
		}
		if value := stringValue(snapshot, "used_range"); value != "" {
			b.WriteString("Used range: ")
			b.WriteString(value)
			b.WriteString("\n")
		}
		if value := stringValue(snapshot, "returned_range"); value != "" {
			b.WriteString("Returned range: ")
			b.WriteString(value)
			b.WriteString("\n")
		}
		if boolValue(snapshot, "style_included") {
			b.WriteString("Style: included\n")
		}
		for _, warning := range stringList(snapshot["warnings"]) {
			b.WriteString("\n> ")
			b.WriteString(warning)
			b.WriteString("\n")
		}
		values := matrixStrings(snapshot["values"])
		if len(values) == 0 {
			b.WriteString("\n_No values_\n")
			b.WriteString(renderWarningsAndHintsMarkdown(env))
			return b.String()
		}
		b.WriteString("\n")
		b.WriteString(markdownMatrixTable(values))
		b.WriteString(renderWarningsAndHintsMarkdown(env))
		return b.String()
	case "cell":
		cell := objectMap(payload["cell"])
		if len(cell) == 0 {
			return ""
		}
		prefix := renderInspectTargetSessionMarkdown(env) + renderInspectTargetInfoMarkdown(payload)
		rows := [][]string{
			{"Sheet", stringValue(cell, "sheet")},
			{"Cell", stringValue(cell, "address")},
			{"Value", emptyDash(stringValue(cell, "value"))},
		}
		return prefix + markdownTable([]string{"Field", "Value"}, rows) + renderWarningsAndHintsMarkdown(env)
	case "form":
		var b strings.Builder
		b.WriteString(renderInspectTargetSessionMarkdown(env))
		if info := renderInspectTargetInfoMarkdown(payload); info != "" {
			b.WriteString(info)
		}
		if form := objectMap(payload["form"]); len(form) > 0 {
			b.WriteString(markdownInspectFormSnapshot(form, ""))
			b.WriteString(renderWarningsAndHintsMarkdown(env))
			return b.String()
		}
		forms := objectMap(payload["forms"])
		for _, basis := range []string{"runtime", "designer"} {
			snapshot := objectMap(forms[basis])
			if len(snapshot) == 0 {
				continue
			}
			b.WriteString(markdownInspectFormSnapshot(snapshot, basis))
		}
		b.WriteString(renderWarningsAndHintsMarkdown(env))
		return b.String()
	case "calls":
		return renderInspectCallsMarkdown(env, payload)
	case "symbols":
		return renderInspectSymbolsMarkdown(env, payload)
	default:
		return ""
	}
}

func (r renderer) renderInspectSymbols(env Envelope, payload map[string]any) string {
	files := listOfObjects(payload["files"])
	var b strings.Builder
	b.WriteString("\n")
	b.WriteString(r.renderInspectTargetSession(env))
	if root := stringValue(payload, "root"); root != "" {
		b.WriteString(kv("Source", root))
	}
	summary := objectMap(payload["summary"])
	if len(summary) > 0 {
		b.WriteString(kv("Files", fmt.Sprintf("%d", intNumber(summary, "files"))))
		b.WriteString(kv("Symbols", fmt.Sprintf("%d", intNumber(summary, "symbols"))))
	}
	if len(files) == 0 {
		b.WriteString("No symbols found.\n")
		b.WriteString(r.renderWarningsAndHints(env))
		return b.String()
	}
	for _, file := range files {
		path := stringValue(file, "path")
		if path == "" {
			path = "<unknown>"
		}
		b.WriteString(path)
		b.WriteString("\n")
		moduleName := stringValue(file, "moduleName")
		if moduleName == "" {
			moduleName = strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
		}
		moduleKind := stringValue(file, "moduleKind")
		if moduleKind == "" {
			moduleKind = "module"
		}
		fmt.Fprintf(&b, "  %s %s\n", inspectModuleKindLabel(moduleKind), moduleName)
		parse := objectMap(file["parse"])
		if boolValue(parse, "hasError") || boolValue(parse, "hasMissing") {
			recovery := []string{}
			if boolValue(parse, "hasError") {
				recovery = append(recovery, "ERROR")
			}
			if boolValue(parse, "hasMissing") {
				recovery = append(recovery, "MISSING")
			}
			fmt.Fprintf(&b, "  Parse recovery: %s\n", strings.Join(recovery, ", "))
		}
		for _, symbol := range listOfObjects(file["symbols"]) {
			if stringValue(symbol, "kind") == "module" {
				continue
			}
			line := inspectSymbolLine(symbol)
			if line == "" {
				continue
			}
			b.WriteString("  ")
			b.WriteString(line)
			b.WriteString("\n")
		}
	}
	b.WriteString(r.renderWarningsAndHints(env))
	return b.String()
}

func (r renderer) renderInspectCalls(env Envelope, payload map[string]any) string {
	calls := listOfObjects(payload["calls"])
	var b strings.Builder
	b.WriteString("\n")
	b.WriteString(r.renderInspectTargetSession(env))
	if root := stringValue(payload, "root"); root != "" {
		b.WriteString(kv("Source", root))
	}
	summary := objectMap(payload["summary"])
	if len(summary) > 0 {
		b.WriteString(kv("Files", fmt.Sprintf("%d", intNumber(summary, "files"))))
		b.WriteString(kv("Calls", fmt.Sprintf("%d", intNumber(summary, "calls"))))
	}
	if len(calls) == 0 {
		b.WriteString("No calls found.\n")
		b.WriteString(r.renderWarningsAndHints(env))
		return b.String()
	}
	for _, fileGroup := range groupInspectCalls(calls) {
		b.WriteString(fileGroup.file)
		b.WriteString("\n")
		for _, callerGroup := range fileGroup.callers {
			b.WriteString(callerGroup.caller)
			b.WriteString("\n")
			for _, call := range callerGroup.calls {
				callee := objectMap(call["callee"])
				name := stringValue(callee, "text")
				if name == "" {
					name = "<unknown>"
				}
				fmt.Fprintf(&b, "  -> %-28s %s:%d\n", name, stringValue(call, "file"), intNumber(objectMap(call["range"]), "startLine"))
			}
		}
	}
	b.WriteString(r.renderWarningsAndHints(env))
	return b.String()
}

func renderInspectCallsMarkdown(env Envelope, payload map[string]any) string {
	calls := listOfObjects(payload["calls"])
	var b strings.Builder
	b.WriteString(renderInspectTargetSessionMarkdown(env))
	if root := stringValue(payload, "root"); root != "" {
		b.WriteString("Source: ")
		b.WriteString(root)
		b.WriteString("\n")
	}
	summary := objectMap(payload["summary"])
	if len(summary) > 0 {
		fmt.Fprintf(&b, "Files: %d\nCalls: %d\n\n", intNumber(summary, "files"), intNumber(summary, "calls"))
	}
	if len(calls) == 0 {
		b.WriteString("_No calls found._\n")
		b.WriteString(renderWarningsAndHintsMarkdown(env))
		return b.String()
	}
	for _, fileGroup := range groupInspectCalls(calls) {
		fmt.Fprintf(&b, "### `%s`\n\n", fileGroup.file)
		for _, callerGroup := range fileGroup.callers {
			fmt.Fprintf(&b, "#### `%s`\n\n", callerGroup.caller)
			for _, call := range callerGroup.calls {
				callee := objectMap(call["callee"])
				name := stringValue(callee, "text")
				if name == "" {
					name = "<unknown>"
				}
				fmt.Fprintf(&b, "- `%s` at `%s:%d`\n", name, stringValue(call, "file"), intNumber(objectMap(call["range"]), "startLine"))
			}
			b.WriteString("\n")
		}
	}
	b.WriteString(renderWarningsAndHintsMarkdown(env))
	return b.String()
}

type inspectCallFileGroup struct {
	file    string
	callers []inspectCallCallerGroup
}

type inspectCallCallerGroup struct {
	caller string
	calls  []map[string]any
}

func groupInspectCalls(calls []map[string]any) []inspectCallFileGroup {
	fileIndex := map[string]int{}
	result := []inspectCallFileGroup{}
	for _, call := range calls {
		file := stringValue(call, "file")
		if file == "" {
			file = "<unknown>"
		}
		fi, ok := fileIndex[file]
		if !ok {
			fi = len(result)
			fileIndex[file] = fi
			result = append(result, inspectCallFileGroup{file: file})
		}
		caller := inspectCallCallerLabel(call)
		ci := -1
		for i, group := range result[fi].callers {
			if group.caller == caller {
				ci = i
				break
			}
		}
		if ci < 0 {
			ci = len(result[fi].callers)
			result[fi].callers = append(result[fi].callers, inspectCallCallerGroup{caller: caller})
		}
		result[fi].callers[ci].calls = append(result[fi].callers[ci].calls, call)
	}
	return result
}

func inspectCallCallerLabel(call map[string]any) string {
	caller := objectMap(call["caller"])
	if qualified := stringValue(caller, "qualifiedName"); qualified != "" {
		return qualified
	}
	if name := stringValue(caller, "name"); name != "" {
		return name
	}
	if module := stringValue(call, "module"); module != "" {
		return module
	}
	return "<module>"
}

func renderInspectSymbolsMarkdown(env Envelope, payload map[string]any) string {
	files := listOfObjects(payload["files"])
	var b strings.Builder
	b.WriteString(renderInspectTargetSessionMarkdown(env))
	if root := stringValue(payload, "root"); root != "" {
		b.WriteString("Source: ")
		b.WriteString(root)
		b.WriteString("\n")
	}
	summary := objectMap(payload["summary"])
	if len(summary) > 0 {
		fmt.Fprintf(&b, "Files: %d\nSymbols: %d\n\n", intNumber(summary, "files"), intNumber(summary, "symbols"))
	}
	if len(files) == 0 {
		b.WriteString("_No symbols found._\n")
		b.WriteString(renderWarningsAndHintsMarkdown(env))
		return b.String()
	}
	for _, file := range files {
		path := stringValue(file, "path")
		if path == "" {
			path = "<unknown>"
		}
		fmt.Fprintf(&b, "### `%s`\n\n", path)
		moduleName := stringValue(file, "moduleName")
		if moduleName == "" {
			moduleName = strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
		}
		moduleKind := stringValue(file, "moduleKind")
		if moduleKind == "" {
			moduleKind = "module"
		}
		fmt.Fprintf(&b, "- %s %s\n", inspectModuleKindLabel(moduleKind), moduleName)
		parse := objectMap(file["parse"])
		if boolValue(parse, "hasError") || boolValue(parse, "hasMissing") {
			recovery := []string{}
			if boolValue(parse, "hasError") {
				recovery = append(recovery, "ERROR")
			}
			if boolValue(parse, "hasMissing") {
				recovery = append(recovery, "MISSING")
			}
			fmt.Fprintf(&b, "- Parse recovery: %s\n", strings.Join(recovery, ", "))
		}
		for _, symbol := range listOfObjects(file["symbols"]) {
			if stringValue(symbol, "kind") == "module" {
				continue
			}
			if line := inspectSymbolLine(symbol); line != "" {
				fmt.Fprintf(&b, "- `%s`\n", line)
			}
		}
		b.WriteString("\n")
	}
	b.WriteString(renderWarningsAndHintsMarkdown(env))
	return b.String()
}

func inspectSymbolLine(symbol map[string]any) string {
	if signature := stringValue(symbol, "signature"); signature != "" {
		return signature
	}
	kind := strings.ReplaceAll(stringValue(symbol, "kind"), "_", " ")
	name := stringValue(symbol, "name")
	if kind == "" {
		return name
	}
	if name == "" {
		return kind
	}
	return titleCase(kind) + " " + name
}

func titleCase(value string) string {
	value = strings.TrimSpace(value)
	if value == "" {
		return ""
	}
	parts := strings.Fields(strings.ReplaceAll(value, "_", " "))
	for i, part := range parts {
		lower := strings.ToLower(part)
		parts[i] = strings.ToUpper(lower[:1]) + lower[1:]
	}
	return strings.Join(parts, " ")
}

func inspectModuleKindLabel(kind string) string {
	if strings.EqualFold(kind, "standard") {
		return "Module"
	}
	return titleCase(kind)
}

func renderInspectFormSnapshot(snapshot map[string]any, heading string) string {
	var b strings.Builder
	label := stringValue(snapshot, "basis")
	if heading != "" {
		label = heading
	}
	if label != "" {
		b.WriteString(kv("Basis", label))
	}
	if name := stringValue(snapshot, "name"); name != "" {
		b.WriteString(kv("Form", name))
	}
	if caption := stringValue(snapshot, "caption"); caption != "" {
		b.WriteString(kv("Caption", caption))
	}
	if width, ok := numberValue(snapshot, "width"); ok {
		if height, okHeight := numberValue(snapshot, "height"); okHeight {
			b.WriteString(kv("Size", fmt.Sprintf("%.0f x %.0f", width, height)))
		}
	}
	if coord := stringValue(snapshot, "coordinate_system"); coord != "" {
		b.WriteString(kv("Coordinates", coord))
	}
	controls := listOfObjects(snapshot["controls"])
	b.WriteString(kv("Controls", fmt.Sprintf("%d", len(controls))))
	for _, control := range controls {
		renderInspectControlLine(&b, control, 0)
	}
	return b.String()
}

func renderInspectControlLine(b *strings.Builder, control map[string]any, depth int) {
	indent := strings.Repeat("  ", depth)
	name := stringValue(control, "name")
	kind := stringValue(control, "type")
	summary := inspectControlSummary(control)
	if name == "" {
		name = "<unnamed>"
	}
	line := fmt.Sprintf("%s- %s [%s]", indent, name, kind)
	if summary != "" {
		line += " " + summary
	}
	b.WriteString(line)
	b.WriteString("\n")
	for _, child := range listOfObjects(control["controls"]) {
		renderInspectControlLine(b, child, depth+1)
	}
}

func inspectControlSummary(control map[string]any) string {
	parts := make([]string, 0, 4)
	if caption := stringValue(control, "caption"); caption != "" {
		parts = append(parts, "caption="+caption)
	}
	if value := stringValue(control, "value"); value != "" {
		parts = append(parts, "value="+value)
	}
	if text := stringValue(control, "text"); text != "" && text != stringValue(control, "value") {
		parts = append(parts, "text="+text)
	}
	if left, ok := numberValue(control, "left"); ok {
		if top, okTop := numberValue(control, "top"); okTop {
			parts = append(parts, fmt.Sprintf("@ %.0f,%.0f", left, top))
		}
	}
	return strings.Join(parts, " | ")
}

func markdownInspectFormSnapshot(snapshot map[string]any, heading string) string {
	var b strings.Builder
	label := stringValue(snapshot, "basis")
	if heading != "" {
		label = heading
	}
	if label != "" {
		b.WriteString("Basis: ")
		b.WriteString(label)
		b.WriteString("\n")
	}
	rows := [][]string{
		{"Form", stringValue(snapshot, "name")},
		{"Caption", stringValue(snapshot, "caption")},
		{"Coordinates", stringValue(snapshot, "coordinate_system")},
		{"Controls", fmt.Sprintf("%d", len(listOfObjects(snapshot["controls"])))},
	}
	if width, ok := numberValue(snapshot, "width"); ok {
		if height, okHeight := numberValue(snapshot, "height"); okHeight {
			rows = append(rows, []string{"Size", fmt.Sprintf("%.0f x %.0f", width, height)})
		}
	}
	b.WriteString(markdownTable([]string{"Field", "Value"}, rows))
	controls := flattenInspectControls(listOfObjects(snapshot["controls"]), 0)
	if len(controls) > 0 {
		b.WriteString("\n")
		b.WriteString(markdownTable([]string{"Control", "Type", "Summary"}, controls))
	}
	b.WriteString("\n")
	return b.String()
}

func flattenInspectControls(controls []map[string]any, depth int) [][]string {
	rows := make([][]string, 0, len(controls))
	for _, control := range controls {
		name := stringValue(control, "name")
		if name == "" {
			name = "<unnamed>"
		}
		rows = append(rows, []string{
			strings.Repeat("  ", depth) + name,
			stringValue(control, "type"),
			inspectControlSummary(control),
		})
		rows = append(rows, flattenInspectControls(listOfObjects(control["controls"]), depth+1)...)
	}
	return rows
}

func renderInspectTargetInfo(payload map[string]any) string {
	info := objectMap(payload["target_info"])
	if len(info) == 0 {
		return ""
	}
	var b strings.Builder
	if kind := stringValue(info, "kind"); kind != "" {
		label := kind
		switch kind {
		case "file":
			label = "saved workbook file"
		case "live_session":
			label = "live session workbook"
		}
		b.WriteString(kv("Snapshot", label))
	}
	if path := stringValue(info, "path"); path != "" {
		b.WriteString(kv("Path", path))
	}
	if note := stringValue(info, "note"); note != "" {
		b.WriteString(kv("Note", note))
	}
	return b.String()
}

func renderInspectTargetInfoMarkdown(payload map[string]any) string {
	info := objectMap(payload["target_info"])
	if len(info) == 0 {
		return ""
	}
	var b strings.Builder
	if kind := stringValue(info, "kind"); kind != "" {
		label := kind
		switch kind {
		case "file":
			label = "saved workbook file"
		case "live_session":
			label = "live session workbook"
		}
		b.WriteString("Snapshot: ")
		b.WriteString(label)
		b.WriteString("\n")
	}
	if path := stringValue(info, "path"); path != "" {
		b.WriteString("Path: ")
		b.WriteString(path)
		b.WriteString("\n")
	}
	if note := stringValue(info, "note"); note != "" {
		b.WriteString("Note: ")
		b.WriteString(note)
		b.WriteString("\n\n")
	}
	return b.String()
}

func (r renderer) renderCreated(env Envelope) string {
	var b strings.Builder
	b.WriteString("\n")
	b.WriteString(r.renderLogs(env))
	if workbook := objectMap(env.Workbook); len(workbook) > 0 {
		for _, key := range sortedKeys(workbook) {
			if value := stringValue(workbook, key); value != "" {
				b.WriteString(kv(labelFromKey(key), value))
			}
		}
	}
	return b.String()
}

func (r renderer) renderSession(env Envelope) string {
	session := objectMap(env.Session)
	workbook := objectMap(env.Workbook)
	if len(session) == 0 && len(workbook) == 0 {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	if path := stringValue(workbook, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	}
	if running, ok := boolValueOK(session, "running"); ok {
		b.WriteString(kv("Running", fmt.Sprintf("%t", running)))
	}
	if open, ok := boolValueOK(session, "workbook_open"); ok {
		b.WriteString(kv("Workbook open", fmt.Sprintf("%t", open)))
	}
	if source := stringValue(session, "source_of_truth"); source != "" {
		b.WriteString(kv("Source of truth", source))
	}
	if known, ok := boolValueOK(session, "userforms_known"); ok && !known {
		b.WriteString(kv("UserForms", "unknown"))
	} else if present, ok := boolValueOK(session, "userforms_present"); ok {
		value := "false"
		if present {
			value = "true"
			if count, ok := numberValue(session, "userform_count"); ok {
				value = fmt.Sprintf("true (%d)", int(count))
			}
		}
		b.WriteString(kv("UserForms", value))
	}
	if save := summarizeSaveRequirement(workbook); save != "" {
		b.WriteString(kv("Save", save))
	}
	b.WriteString(r.renderWarningsAndHints(env))
	b.WriteString(r.renderLogs(env))
	return b.String()
}

func (r renderer) renderSave(env Envelope) string {
	workbook := objectMap(env.Workbook)
	if len(workbook) == 0 {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	if path := stringValue(workbook, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	}
	if sessionSummary := summarizeSessionUsage(workbook); sessionSummary != "" {
		b.WriteString(kv("Session", sessionSummary))
	}
	if saved, ok := boolValueOK(workbook, "saved"); ok && saved {
		b.WriteString(kv("Result", "saved live session workbook to disk"))
	}
	if save := summarizeSaveRequirement(workbook); save != "" {
		b.WriteString(kv("Save", save))
	}
	b.WriteString(r.renderWarningsAndHints(env))
	b.WriteString(r.renderLogs(env))
	return b.String()
}

func (r renderer) renderStatus(env Envelope) string {
	project := objectMap(env.Project)
	session := objectMap(env.Session)
	state := objectMap(env.State)
	if len(project) == 0 && len(session) == 0 && len(state) == 0 {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	if len(project) > 0 {
		b.WriteString("Project:\n")
		b.WriteString(kv("Root", stringValue(project, "root")))
		b.WriteString(kv("Workbook", stringValue(project, "workbook_path")))
		srcPaths := stringList(project["src_paths"])
		for _, srcPath := range srcPaths {
			b.WriteString(kv("Source", srcPath))
		}
	}
	if len(session) > 0 {
		b.WriteString("Session:\n")
		if active, ok := boolValueOK(session, "active"); ok {
			dirty, dirtyKnown := boolValueOK(session, "dirty")
			b.WriteString(kv("Status", sessionSummary(active, dirty, dirtyKnown)))
		}
		if dirty, ok := boolValueOK(session, "dirty"); ok && dirty {
			b.WriteString(kv("Dirty", "true"))
		}
	}
	if len(state) > 0 {
		b.WriteString("State:\n")
		if srcNewer, ok := boolValueOK(state, "src_newer_than_workbook"); ok {
			b.WriteString(kv("Source newer", fmt.Sprintf("%t", srcNewer)))
		}
		if liveNewer, ok := boolValueOK(state, "live_session_newer_than_disk"); ok {
			b.WriteString(kv("Live newer", fmt.Sprintf("%t", liveNewer)))
		}
		if saved, ok := boolValueOK(state, "workbook_saved"); ok {
			b.WriteString(kv("Workbook saved", fmt.Sprintf("%t", saved)))
		}
		if sourceOfTruth := stringValue(state, "source_of_truth"); sourceOfTruth != "" {
			b.WriteString(kv("Source of truth", sourceOfTruth))
		}
		if mtime := stringValue(state, "workbook_last_modified_at"); mtime != "" {
			b.WriteString(kv("Workbook mtime", mtime))
		}
		if mtime := stringValue(state, "latest_source_modified_at"); mtime != "" {
			b.WriteString(kv("Source mtime", mtime))
		}
		if mtime := stringValue(state, "push_state_last_modified_at"); mtime != "" {
			b.WriteString(kv("Push state mtime", mtime))
		}
	}
	b.WriteString(r.renderWarningsAndHints(env))
	b.WriteString(r.renderLogs(env))
	return b.String()
}

func sessionSummary(active, dirty, dirtyKnown bool) string {
	if !active {
		return "inactive"
	}
	if !dirtyKnown {
		return "active"
	}
	if dirty {
		return "active, dirty"
	}
	return "active, clean"
}

func (r renderer) renderLogs(env Envelope) string {
	return r.renderLogsSkipping(env, nil)
}

func (r renderer) renderLogsSkipping(env Envelope, skip map[string]bool) string {
	if len(env.Logs) == 0 {
		return ""
	}
	var b strings.Builder
	for _, line := range env.Logs {
		if skip[line] {
			continue
		}
		if strings.HasPrefix(line, "diff ") ||
			strings.HasPrefix(line, "+++ b/") ||
			strings.HasPrefix(line, "--- a/") ||
			strings.HasPrefix(line, "@@ ") ||
			strings.HasPrefix(line, "+") ||
			strings.HasPrefix(line, "-") ||
			strings.HasPrefix(line, " ") {
			b.WriteString(line)
		} else {
			b.WriteString("- ")
			b.WriteString(line)
		}
		b.WriteString("\n")
	}
	return b.String()
}

func renderWarningList(warnings []map[string]any) string {
	if len(warnings) == 0 {
		return ""
	}
	var b strings.Builder
	for _, warning := range warnings {
		b.WriteString("- ")
		if code := stringValue(warning, "code"); code != "" {
			b.WriteString("[")
			b.WriteString(code)
			b.WriteString("] ")
		}
		b.WriteString(stringValue(warning, "message"))
		b.WriteString("\n")
	}
	return b.String()
}

func renderHintList(hints []map[string]any) string {
	if len(hints) == 0 {
		return ""
	}
	var b strings.Builder
	for _, hint := range hints {
		b.WriteString("- ")
		if code := stringValue(hint, "code"); code != "" {
			b.WriteString("[")
			b.WriteString(code)
			b.WriteString("] ")
		}
		b.WriteString(stringValue(hint, "message"))
		b.WriteString("\n")
	}
	return b.String()
}

func (r renderer) renderTargetSession(env Envelope) string {
	target := objectMap(env.Target)
	session := objectMap(env.Session)
	if len(target) == 0 && len(session) == 0 {
		return ""
	}
	var b strings.Builder
	if len(target) > 0 {
		b.WriteString(kv("Target", summarizeTarget(target)))
	}
	if len(session) > 0 {
		b.WriteString(kv("Session state", summarizeSessionState(session)))
	}
	if note := summarizeStateNote(target, session, listOfObjects(env.Warnings)); note != "" {
		b.WriteString(kv("State note", note))
	}
	return b.String()
}

func (r renderer) renderInspectTargetSession(env Envelope) string {
	text := r.renderTargetSession(env)
	if text == "" {
		return ""
	}
	return text
}

func (r renderer) renderWarningsAndHints(env Envelope) string {
	warnings := listOfObjects(env.Warnings)
	hints := listOfObjects(env.Hints)
	if len(warnings) == 0 && len(hints) == 0 {
		return ""
	}
	var b strings.Builder
	if renderedWarnings := renderWarningList(warnings); renderedWarnings != "" {
		b.WriteString("Warnings:\n")
		b.WriteString(renderedWarnings)
	}
	if renderedHints := renderHintList(hints); renderedHints != "" {
		b.WriteString("Hints:\n")
		b.WriteString(renderedHints)
	}
	return b.String()
}

func renderInspectTargetSessionMarkdown(env Envelope) string {
	target := objectMap(env.Target)
	session := objectMap(env.Session)
	if len(target) == 0 && len(session) == 0 {
		return ""
	}
	var b strings.Builder
	if len(target) > 0 {
		b.WriteString("Target: ")
		b.WriteString(summarizeTarget(target))
		b.WriteString("\n")
	}
	if len(session) > 0 {
		b.WriteString("Session state: ")
		b.WriteString(summarizeSessionState(session))
		b.WriteString("\n")
	}
	if note := summarizeStateNote(target, session, listOfObjects(env.Warnings)); note != "" {
		b.WriteString("State note: ")
		b.WriteString(note)
		b.WriteString("\n")
	}
	if b.Len() > 0 {
		b.WriteString("\n")
	}
	return b.String()
}

func renderWarningsAndHintsMarkdown(env Envelope) string {
	warnings := listOfObjects(env.Warnings)
	hints := listOfObjects(env.Hints)
	if len(warnings) == 0 && len(hints) == 0 {
		return ""
	}
	var b strings.Builder
	if len(warnings) > 0 {
		for _, warning := range warnings {
			b.WriteString("\n> Warning")
			if code := stringValue(warning, "code"); code != "" {
				b.WriteString(" [")
				b.WriteString(code)
				b.WriteString("]")
			}
			b.WriteString(": ")
			b.WriteString(stringValue(warning, "message"))
			b.WriteString("\n")
		}
	}
	if len(hints) > 0 {
		for _, hint := range hints {
			b.WriteString("\n> Hint")
			if code := stringValue(hint, "code"); code != "" {
				b.WriteString(" [")
				b.WriteString(code)
				b.WriteString("]")
			}
			b.WriteString(": ")
			b.WriteString(stringValue(hint, "message"))
			b.WriteString("\n")
		}
	}
	return b.String()
}

func (r renderer) renderBridge(env Envelope) string {
	bridge := objectMap(env.Bridge)
	if len(bridge) == 0 {
		return ""
	}
	summary := stringValue(bridge, "host")
	details := make([]string, 0, 2)
	if summary == "" {
		summary = stringValue(bridge, "name")
	}
	if edition := stringValue(bridge, "edition"); edition != "" {
		details = append(details, edition)
	}
	if version := stringValue(bridge, "version"); version != "" {
		details = append(details, version)
	}
	if runtime := stringValue(bridge, "runtime"); runtime != "" {
		details = append(details, runtime)
	}
	if arch := stringValue(bridge, "architecture"); arch != "" {
		details = append(details, arch)
	}
	if summary == "" && len(details) == 0 {
		return ""
	}
	if len(details) > 0 {
		if summary != "" {
			summary += " (" + strings.Join(details, " ") + ")"
		} else {
			summary = strings.Join(details, " ")
		}
	}
	return kv("Bridge", summary)
}

func summarizeRunWorkbookResult(workbook map[string]any) string {
	if len(workbook) == 0 {
		return ""
	}
	if saveAs := stringValue(workbook, "save_as"); saveAs != "" {
		return "copied to " + saveAs
	}
	saved, ok := boolValueOK(workbook, "saved")
	if !ok {
		return ""
	}
	session := boolValue(workbook, "session")
	if saved {
		if session {
			return "saved live session workbook to disk"
		}
		return "saved in place"
	}
	return "left unchanged on disk"
}

func summarizeWorkbookSourceResult(command string, workbook, source map[string]any) string {
	if len(workbook) == 0 {
		return ""
	}
	session := boolValue(workbook, "session")
	if save := summarizeSaveRequirement(workbook); save != "" {
		return save
	}
	saved, savedKnown := boolValueOK(workbook, "saved")
	switch command {
	case "push":
		changedOnly, _ := boolValueOK(source, "changed_only")
		changed, changedKnown := boolValueOK(source, "changed")
		if changedOnly && changedKnown && !changed {
			return "skipped workbook import; source unchanged"
		}
		if savedKnown && saved {
			if session {
				return "saved live session workbook to disk"
			}
			return "saved workbook in place"
		}
	case "pull":
		if session {
			return "exported from live session workbook"
		}
	case "attach":
		return "inspected the active workbook"
	}
	return ""
}

func summarizeExportImageTarget(target map[string]any) string {
	switch stringValue(target, "kind") {
	case "live_session":
		return "live session workbook"
	case "file":
		return "saved workbook file"
	default:
		return ""
	}
}

func summarizeEditSelector(edit map[string]any) string {
	if len(edit) == 0 {
		return ""
	}
	sheet := stringValue(edit, "sheet")
	for _, key := range []string{"cell", "range", "rows", "columns"} {
		if selector := stringValue(edit, key); selector != "" {
			if sheet != "" {
				return sheet + "!" + selector
			}
			return selector
		}
	}
	return sheet
}

func summarizeEditMutation(edit map[string]any) string {
	mutation := objectMap(edit["mutation"])
	if len(mutation) == 0 {
		return ""
	}
	if formula := objectMap(mutation["formula"]); len(formula) > 0 {
		return "formula -> " + stringValue(formula, "after")
	}
	if value := objectMap(mutation["value"]); len(value) > 0 {
		return "value -> " + stringValue(value, "after")
	}
	if style := objectMap(mutation["style"]); len(style) > 0 {
		if fill := objectMap(style["fill"]); len(fill) > 0 {
			return "fill -> " + stringValue(fill, "after")
		}
		if clear := stringValue(style, "cleared"); clear != "" {
			return "clear " + clear
		}
	}
	if clear := objectMap(mutation["clear"]); len(clear) > 0 {
		if mode := stringValue(clear, "mode"); mode != "" {
			return "clear " + mode
		}
	}
	if rowHeight := objectMap(mutation["row_height"]); len(rowHeight) > 0 {
		return "row height -> " + stringValue(rowHeight, "after")
	}
	if columnWidth := objectMap(mutation["column_width"]); len(columnWidth) > 0 {
		return "column width -> " + stringValue(columnWidth, "after")
	}
	return ""
}

func summarizeEditEvents(edit map[string]any) string {
	events := objectMap(edit["events"])
	if len(events) == 0 {
		return ""
	}
	parts := []string{}
	if mode := stringValue(events, "mode"); mode != "" {
		parts = append(parts, "mode="+mode)
	}
	if before, ok := boolValueOK(events, "enable_events_before"); ok {
		parts = append(parts, "before="+yesNo(before))
	}
	if after, ok := boolValueOK(events, "enable_events_after"); ok {
		parts = append(parts, "after="+yesNo(after))
	}
	if restored, ok := boolValueOK(events, "restored"); ok {
		parts = append(parts, "restored="+yesNo(restored))
	}
	return strings.Join(parts, ", ")
}

func summarizeSessionUsage(workbook map[string]any) string {
	if len(workbook) == 0 || !boolValue(workbook, "session") {
		return ""
	}
	switch stringValue(workbook, "session_mode") {
	case "explicit":
		return "explicit xlflow session workbook"
	case "auto":
		return "auto-reused matching xlflow session workbook"
	case "managed":
		return "managed xlflow session workbook"
	default:
		return "xlflow session workbook"
	}
}

func summarizeTarget(target map[string]any) string {
	if len(target) == 0 {
		return ""
	}
	kind := stringValue(target, "kind")
	description := stringValue(target, "description")
	path := stringValue(target, "path")
	if description == "" {
		switch kind {
		case "source":
			description = "source files"
		case "file":
			description = "saved file"
		case "live_session":
			description = "live session"
		}
	}
	if path != "" {
		if description != "" {
			return description + " (" + path + ")"
		}
		return path
	}
	return description
}

func summarizeSessionState(session map[string]any) string {
	if len(session) == 0 {
		return ""
	}
	parts := make([]string, 0, 4)
	if active, ok := boolValueOK(session, "active"); ok {
		if active {
			parts = append(parts, "active")
		} else {
			parts = append(parts, "inactive")
		}
	}
	if dirty, ok := boolValueOK(session, "dirty"); ok && dirty {
		parts = append(parts, "dirty")
	}
	if saveRequired, ok := boolValueOK(session, "live_newer_than_disk"); ok && saveRequired {
		parts = append(parts, "SAVE REQUIRED")
		parts = append(parts, "live workbook is newer than disk")
	} else if saveRequired, ok := boolValueOK(session, "save_required"); ok && saveRequired {
		parts = append(parts, "SAVE REQUIRED")
		parts = append(parts, "live workbook is newer than disk")
	}
	if len(parts) == 0 {
		return ""
	}
	return strings.Join(parts, ", ")
}

func summarizeSaveRequirement(workbook map[string]any) string {
	if len(workbook) == 0 || !boolValue(workbook, "needs_save") {
		return ""
	}
	return "SAVE REQUIRED: live workbook is newer than disk; run xlflow save before session stop"
}

func summarizeStateNote(target, session map[string]any, warnings []map[string]any) string {
	if len(session) == 0 && len(target) == 0 {
		return ""
	}
	liveNewer := boolValue(session, "live_newer_than_disk") || boolValue(session, "save_required")
	userFormsPresent := boolValue(session, "userforms_present")
	userFormsKnown, userFormsKnownSet := boolValueOK(session, "userforms_known")
	if !userFormsPresent {
		for _, warning := range warnings {
			switch stringValue(warning, "code") {
			case "userform_state_partial", "userform_unsaved_session_state", "userform_inspect_saved_file":
				userFormsPresent = true
			case "userform_detection_unavailable":
				userFormsKnown = false
				userFormsKnownSet = true
			}
		}
	}
	if liveNewer && userFormsPresent {
		return "UserForm project: save before disk inspect/pull review."
	}
	if liveNewer && userFormsKnownSet && !userFormsKnown {
		return "Workbook may contain UserForms; save before disk inspect/pull review."
	}
	if note := stringValue(target, "note"); note != "" {
		return note
	}
	if stringValue(target, "capture_state") == "temporary_copy" {
		return "Runtime inspection/export used a temporary workbook copy."
	}
	return ""
}

func visibleLabel(m map[string]any) string {
	if boolValue(m, "visible") {
		return "visible"
	}
	return "hidden"
}

func emptyDash(value string) string {
	if value == "" {
		return "-"
	}
	return value
}

func intNumber(m map[string]any, key string) int {
	n, _ := numberValue(m, key)
	return int(n)
}

func matrixStrings(value any) [][]string {
	if value == nil {
		return nil
	}
	rows, ok := value.([]any)
	if !ok {
		data, err := json.Marshal(value)
		if err != nil {
			return nil
		}
		if err := json.Unmarshal(data, &rows); err != nil {
			return nil
		}
	}
	out := make([][]string, 0, len(rows))
	for _, rawRow := range rows {
		row, ok := rawRow.([]any)
		if !ok {
			data, err := json.Marshal(rawRow)
			if err != nil {
				continue
			}
			if err := json.Unmarshal(data, &row); err != nil {
				continue
			}
		}
		line := make([]string, 0, len(row))
		for _, cell := range row {
			if cell == nil {
				line = append(line, "")
				continue
			}
			line = append(line, fmt.Sprint(cell))
		}
		out = append(out, line)
	}
	return out
}

func markdownSheetTable(sheets []map[string]any) string {
	rows := make([][]string, 0, len(sheets))
	for _, sheet := range sheets {
		rows = append(rows, []string{
			fmt.Sprintf("%d", intNumber(sheet, "index")),
			stringValue(sheet, "name"),
			fmt.Sprintf("%t", boolValue(sheet, "visible")),
			emptyDash(stringValue(sheet, "used_range")),
			fmt.Sprintf("%d", intNumber(sheet, "row_count")),
			fmt.Sprintf("%d", intNumber(sheet, "column_count")),
		})
	}
	return markdownTable([]string{"Index", "Name", "Visible", "Used range", "Rows", "Columns"}, rows)
}

func markdownMatrixTable(values [][]string) string {
	width := 0
	for _, row := range values {
		if len(row) > width {
			width = len(row)
		}
	}
	if width == 0 {
		return "_No values_\n"
	}
	headers := make([]string, 0, width)
	rows := make([][]string, 0, len(values))
	for i := 1; i <= width; i++ {
		headers = append(headers, fmt.Sprintf("C%d", i))
	}
	for _, row := range values {
		line := make([]string, width)
		copy(line, row)
		rows = append(rows, line)
	}
	return markdownTable(headers, rows)
}

func markdownTable(headers []string, rows [][]string) string {
	if len(headers) == 0 {
		return ""
	}
	var b strings.Builder
	b.WriteString("| ")
	for i, header := range headers {
		if i > 0 {
			b.WriteString(" | ")
		}
		b.WriteString(escapeMarkdownTableCell(header))
	}
	b.WriteString(" |\n| ")
	for i := range headers {
		if i > 0 {
			b.WriteString(" | ")
		}
		b.WriteString("---")
	}
	b.WriteString(" |\n")
	for _, row := range rows {
		b.WriteString("| ")
		for i := range headers {
			if i > 0 {
				b.WriteString(" | ")
			}
			value := ""
			if i < len(row) {
				value = row[i]
			}
			b.WriteString(escapeMarkdownTableCell(value))
		}
		b.WriteString(" |\n")
	}
	return b.String()
}

func escapeMarkdownTableCell(value string) string {
	replacer := strings.NewReplacer("|", "\\|", "\n", "<br>", "\r", "")
	return replacer.Replace(value)
}

func yesNo(v bool) string {
	if v {
		return "yes"
	}
	return "no"
}

func (r renderer) errorBlock(env Envelope) string {
	if env.Error == nil {
		return "\nError: command failed\n"
	}
	var b strings.Builder
	b.WriteString("\n")
	b.WriteString(r.style("Error:", "196", true))
	b.WriteString(" ")
	b.WriteString(env.Error.Message)
	b.WriteString("\n")
	if env.Error.Code != "" {
		b.WriteString(kv("Code", env.Error.Code))
	}
	if env.Error.Phase != "" {
		b.WriteString(kv("Phase", env.Error.Phase))
	}
	if env.Error.Source != "" {
		b.WriteString(kv("Source", env.Error.Source))
	}
	if env.Error.Number != 0 {
		b.WriteString(kv("Number", fmt.Sprintf("%d", env.Error.Number)))
	}
	if env.Error.Line != 0 {
		b.WriteString(kv("Line", fmt.Sprintf("%d", env.Error.Line)))
	}
	return b.String()
}

func (r renderer) style(s, color string, bold bool) string {
	if !r.color {
		return s
	}
	style := lipgloss.NewStyle()
	if color != "" {
		style = style.Foreground(lipgloss.Color(color))
	}
	if bold {
		style = style.Bold(true)
	}
	return style.Render(s)
}

func kv(key, value string) string {
	if value == "" {
		return ""
	}
	return fmt.Sprintf("%-14s %s\n", key+":", value)
}

func renderSpecMetadata(spec map[string]any) string {
	if len(spec) == 0 {
		return ""
	}
	var b strings.Builder
	if path := stringValue(spec, "path"); path != "" {
		b.WriteString(kv("Spec", path))
	}
	if format := stringValue(spec, "format"); format != "" {
		b.WriteString(kv("Spec format", strings.ToUpper(format)))
	}
	if field := stringValue(spec, "field"); field != "" {
		b.WriteString(kv("Spec field", field))
	}
	line, lineOK := numberValue(spec, "line")
	column, columnOK := numberValue(spec, "column")
	if lineOK && columnOK {
		b.WriteString(kv("Spec location", fmt.Sprintf("line %d, column %d", int(line), int(column))))
	} else if lineOK {
		b.WriteString(kv("Spec location", fmt.Sprintf("line %d", int(line))))
	}
	return b.String()
}

func objectMap(value any) map[string]any {
	if value == nil {
		return map[string]any{}
	}
	switch v := value.(type) {
	case map[string]any:
		return v
	case map[string]string:
		out := make(map[string]any, len(v))
		for key, item := range v {
			out[key] = item
		}
		return out
	default:
		data, err := json.Marshal(value)
		if err != nil {
			return map[string]any{}
		}
		var out map[string]any
		if err := json.Unmarshal(data, &out); err != nil {
			return map[string]any{}
		}
		if out == nil {
			return map[string]any{}
		}
		return out
	}
}

func listOfObjects(value any) []map[string]any {
	if value == nil {
		return nil
	}
	switch v := value.(type) {
	case []map[string]any:
		return v
	case []any:
		out := make([]map[string]any, 0, len(v))
		for _, item := range v {
			out = append(out, objectMap(item))
		}
		return out
	default:
		data, err := json.Marshal(value)
		if err != nil {
			return nil
		}
		var out []map[string]any
		if err := json.Unmarshal(data, &out); err != nil {
			return nil
		}
		return out
	}
}

func stringList(value any) []string {
	switch v := value.(type) {
	case []string:
		return v
	case []any:
		out := make([]string, 0, len(v))
		for _, item := range v {
			out = append(out, fmt.Sprint(item))
		}
		return out
	default:
		return nil
	}
}

func stringValue(m map[string]any, key string) string {
	value, ok := m[key]
	if !ok || value == nil {
		return ""
	}
	switch v := value.(type) {
	case string:
		return v
	default:
		return fmt.Sprint(v)
	}
}

func boolValue(m map[string]any, key string) bool {
	value, _ := boolValueOK(m, key)
	return value
}

func boolValueOK(m map[string]any, key string) (bool, bool) {
	value, ok := m[key]
	if !ok {
		return false, false
	}
	switch v := value.(type) {
	case bool:
		return v, true
	default:
		return false, false
	}
}

func numberValue(m map[string]any, key string) (float64, bool) {
	value, ok := m[key]
	if !ok {
		return 0, false
	}
	switch v := value.(type) {
	case float64:
		return v, true
	case float32:
		return float64(v), true
	case int:
		return float64(v), true
	case int64:
		return float64(v), true
	case json.Number:
		n, err := v.Float64()
		return n, err == nil
	default:
		return 0, false
	}
}

func sortedKeys(m map[string]any) []string {
	keys := make([]string, 0, len(m))
	for key := range m {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	return keys
}

func labelFromKey(key string) string {
	parts := strings.Split(key, "_")
	for i, part := range parts {
		if part == "" {
			continue
		}
		if strings.EqualFold(part, "vba") {
			parts[i] = "VBA"
			continue
		}
		parts[i] = strings.ToUpper(part[:1]) + part[1:]
	}
	return strings.Join(parts, " ")
}
