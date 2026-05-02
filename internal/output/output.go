package output

import (
	"encoding/json"
	"errors"
	"fmt"
	"io"
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
}

type Envelope struct {
	Status  string   `json:"status"`
	Command string   `json:"command"`
	Error   *Error   `json:"error"`
	Logs    []string `json:"logs"`

	Diagnostics   any `json:"diagnostics,omitempty"`
	Workbook      any `json:"workbook,omitempty"`
	Backup        any `json:"backup,omitempty"`
	Source        any `json:"source,omitempty"`
	Bridge        any `json:"bridge,omitempty"`
	Macro         any `json:"macro,omitempty"`
	Macros        any `json:"macros,omitempty"`
	Issues        any `json:"issues,omitempty"`
	Tests         any `json:"tests,omitempty"`
	Diff          any `json:"diff,omitempty"`
	Trace         any `json:"trace,omitempty"`
	GUIBoundaries any `json:"gui_boundaries,omitempty"`
	UI            any `json:"ui,omitempty"`
	Session       any `json:"session,omitempty"`
	Runner        any `json:"runner,omitempty"`
	Analysis      any `json:"analysis,omitempty"`
	Check         any `json:"check,omitempty"`
	RunDiagnostic any `json:"run_diagnostic,omitempty"`
}

type Options struct {
	JSON        bool
	Interactive bool
	Color       bool
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
		return Write(w, env, true)
	}
	text := renderHuman(env, opts)
	if text == "" {
		text = renderFallback(env)
	}
	_, err := fmt.Fprint(w, text)
	return err
}

func renderHuman(env Envelope, opts Options) string {
	r := renderer{color: opts.Color}
	var b strings.Builder
	b.WriteString(r.title(env))
	b.WriteString("\n")
	if env.Status == StatusFailed {
		b.WriteString(r.errorBlock(env))
	}
	b.WriteString(r.renderBridge(env))
	switch env.Command {
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
	case "trace":
		if env.Issues != nil {
			b.WriteString(r.renderLint(env))
		}
		if env.Analysis != nil {
			b.WriteString(r.renderAnalysis(env))
		}
		b.WriteString(r.renderTraceCommand(env))
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
	case "diff":
		b.WriteString(r.renderDiff(env))
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
	workbook := objectMap(env.Workbook)
	var b strings.Builder
	b.WriteString("\n")
	b.WriteString(r.checkLine(boolValue(diag, "excel_installed"), "Excel automation", "Excel COM can be created"))
	if path := stringValue(workbook, "path"); path != "" {
		b.WriteString(r.checkLine(boolValue(diag, "workbook_openable"), "Workbook", path))
	} else {
		b.WriteString(r.checkLine(false, "Workbook", "No configured workbook was checked"))
	}
	b.WriteString(r.checkLine(boolValue(diag, "vbide_access"), "VBIDE access", "VBA project object model is available"))
	if fix := stringValue(diag, "fix"); fix != "" {
		b.WriteString("\n")
		b.WriteString(r.style("Fix:", "214", true))
		b.WriteString(" ")
		b.WriteString(fix)
		b.WriteString("\n")
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

func (r renderer) renderRun(env Envelope) string {
	macro := objectMap(env.Macro)
	workbook := objectMap(env.Workbook)
	trace := objectMap(env.Trace)
	if len(macro) == 0 && len(workbook) == 0 && len(trace) == 0 && env.RunDiagnostic == nil {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	if name := stringValue(macro, "name"); name != "" {
		b.WriteString(kv("Macro", name))
	}
	if duration, ok := numberValue(macro, "duration_ms"); ok {
		b.WriteString(kv("Duration", fmt.Sprintf("%dms", int(duration))))
	}
	if path := stringValue(workbook, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	}
	if summary := summarizeRunWorkbookResult(workbook); summary != "" {
		b.WriteString(kv("Result", summary))
	}
	if helper := summarizeRunTraceHelper(trace); helper != "" {
		b.WriteString(kv("Trace helper", helper))
	}
	if hint := stringValue(trace, "hint"); hint != "" {
		b.WriteString(kv("Trace hint", hint))
	}
	events := listOfObjects(trace["events"])
	b.WriteString(r.renderLogsSkipping(env, traceEventLogLines(events)))
	if len(events) > 0 {
		b.WriteString("\n")
		b.WriteString(r.style("Trace", "", true))
		b.WriteString("\n")
		for _, event := range events {
			b.WriteString("  ")
			if ts := stringValue(event, "timestamp"); ts != "" {
				b.WriteString("[")
				b.WriteString(ts)
				b.WriteString("] ")
			}
			b.WriteString(stringValue(event, "message"))
			b.WriteString("\n")
		}
	}
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
			for _, key := range []string{"module", "procedure", "file"} {
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
	return b.String()
}

func (r renderer) renderTest(env Envelope) string {
	tests := listOfObjects(env.Tests)
	if env.Tests == nil {
		return r.renderLogs(env)
	}
	passed := 0
	failed := 0
	notRun := 0
	for _, test := range tests {
		switch stringValue(test, "status") {
		case "passed":
			passed++
		case "failed":
			failed++
		default:
			notRun++
		}
	}
	var b strings.Builder
	b.WriteString("\n")
	summary := fmt.Sprintf("%d passed, %d failed", passed, failed)
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
	return b.String()
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
	var b strings.Builder
	b.WriteString("\n")
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
	}
	return b.String()
}

func (r renderer) renderWorkbookSource(env Envelope) string {
	workbook := objectMap(env.Workbook)
	backup := objectMap(env.Backup)
	source := objectMap(env.Source)
	if len(workbook) == 0 && len(backup) == 0 && len(source) == 0 {
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
	if summary := summarizeWorkbookSourceResult(env.Command, workbook, source); summary != "" {
		b.WriteString(kv("Result", summary))
	}
	if updated, ok := boolValueOK(source, "updated"); ok {
		b.WriteString(kv("Source updated", fmt.Sprintf("%t", updated)))
	}
	b.WriteString(r.renderLogs(env))
	return b.String()
}

func (r renderer) renderTraceCommand(env Envelope) string {
	workbook := objectMap(env.Workbook)
	source := objectMap(env.Source)
	trace := objectMap(env.Trace)
	if len(workbook) == 0 && len(source) == 0 && len(trace) == 0 {
		return r.renderLogs(env)
	}
	var b strings.Builder
	b.WriteString("\n")
	if path := stringValue(workbook, "path"); path != "" {
		b.WriteString(kv("Workbook", path))
	}
	if path := stringValue(source, "path"); path != "" {
		b.WriteString(kv("Source", path))
	}
	if summary := summarizeTraceCommandResult(workbook, trace); summary != "" {
		b.WriteString(kv("Result", summary))
	}
	if helper := summarizeTraceCommandHelper(trace, source); helper != "" {
		b.WriteString(kv("Trace helper", helper))
	}
	if logDir := stringValue(trace, "log_dir"); logDir != "" {
		b.WriteString(kv("Trace dir", logDir))
	}
	if updated, ok := boolValueOK(source, "updated"); ok {
		b.WriteString(kv("Source updated", fmt.Sprintf("%t", updated)))
	}
	b.WriteString(r.renderLogs(env))
	return b.String()
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
		b.WriteString("- ")
		b.WriteString(line)
		b.WriteString("\n")
	}
	return b.String()
}

func traceEventLogLines(events []map[string]any) map[string]bool {
	if len(events) == 0 {
		return nil
	}
	lines := make(map[string]bool, len(events))
	for _, event := range events {
		timestamp := stringValue(event, "timestamp")
		message := stringValue(event, "message")
		if timestamp == "" || message == "" {
			continue
		}
		lines["["+timestamp+"] "+message] = true
	}
	return lines
}

func (r renderer) renderBridge(env Envelope) string {
	bridge := objectMap(env.Bridge)
	if len(bridge) == 0 {
		return ""
	}
	summary := stringValue(bridge, "host")
	details := make([]string, 0, 2)
	if edition := stringValue(bridge, "edition"); edition != "" {
		details = append(details, edition)
	}
	if version := stringValue(bridge, "version"); version != "" {
		details = append(details, version)
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
	if session {
		return "UNSAVED session changes: live workbook differs from disk; run xlflow save --session before session stop"
	}
	return "left unchanged on disk"
}

func summarizeRunTraceHelper(trace map[string]any) string {
	if len(trace) == 0 {
		return ""
	}
	switch stringValue(trace, "lifecycle") {
	case "temporary":
		if boolValue(trace, "temporary_reverted") {
			return "temporary helper injected for this run and reverted afterward"
		}
		return "temporary helper injected for this run"
	case "existing":
		return "using an existing workbook trace helper"
	}
	return ""
}

func summarizeWorkbookSourceResult(command string, workbook, source map[string]any) string {
	if len(workbook) == 0 {
		return ""
	}
	session := boolValue(workbook, "session")
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
		if session {
			return "UNSAVED session changes: live workbook differs from disk; run xlflow save --session before session stop"
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

func summarizeTraceCommandResult(workbook, trace map[string]any) string {
	if len(trace) == 0 {
		return ""
	}
	session := boolValue(workbook, "session")
	saved := boolValue(workbook, "saved")
	switch stringValue(trace, "lifecycle") {
	case "enabled":
		if session && saved {
			return "saved live session workbook with trace helper"
		}
		if saved {
			return "saved workbook with trace helper"
		}
		return "enabled trace helper"
	case "disabled":
		if session && saved {
			return "saved live session workbook after trace helper removal"
		}
		if saved {
			return "saved workbook after trace helper removal"
		}
		return "disabled trace helper"
	}
	if stringValue(trace, "status") != "" {
		if session {
			return "inspected trace helper state on the live session workbook"
		}
		return "inspected trace helper state"
	}
	return ""
}

func summarizeTraceCommandHelper(trace, source map[string]any) string {
	if len(trace) == 0 {
		return ""
	}
	switch stringValue(trace, "lifecycle") {
	case "enabled":
		if stringValue(source, "path") != "" {
			return "persisted in workbook and source"
		}
		return "enabled in workbook only"
	case "disabled":
		parts := make([]string, 0, 2)
		if boolValue(trace, "workbook_removed") {
			parts = append(parts, "removed from workbook")
		}
		if boolValue(trace, "source_removed") {
			parts = append(parts, "removed from source")
		}
		if len(parts) > 0 {
			return strings.Join(parts, "; ")
		}
		return "disabled"
	}
	if stringValue(trace, "status") != "" {
		return fmt.Sprintf("workbook=%s source=%s bundled=%s",
			yesNo(boolValue(trace, "workbook_injected")),
			yesNo(boolValue(trace, "source_exists")),
			yesNo(boolValue(trace, "source_matches_bundled")),
		)
	}
	return ""
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
