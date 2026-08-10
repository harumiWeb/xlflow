package analyze

import (
	"context"
	"fmt"
	"strings"

	vbadf "github.com/harumiWeb/xlflow/internal/vba/dataflow"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func (a Analyzer) dataFlowFindingsContext(ctx context.Context, file parsedFile, proc sourceProcedure) ([]Finding, error) {
	// VBA224 owns generic source-to-sink flows. Process-launch sinks are
	// projected through VBA236 so a Shell/Run/Exec call is never reported by
	// both rules. Keep the shared data-flow run enabled when either rule is on.
	if (!a.Config.Analyze.DetectUntrustedDataFlow && !a.Config.Analyze.DetectUnsafeCommandConstruction) || proc.Graph == nil {
		return nil, nil
	}
	ir, ok := procedureIRForSource(file.IR, proc)
	if !ok {
		return nil, nil
	}
	result, err := vbadf.AnalyzeProcedureContext(ctx, ir, *proc.Graph, vbadf.Options{
		Conservative:    true,
		IsKnownConstant: func(name string) bool { return a.isKnownDataFlowConstant(file, name) },
	})
	if err != nil {
		return nil, err
	}
	findings := make([]Finding, 0, len(result.Findings)+len(result.CommandFindings))
	commandKeys := make(map[string]bool, len(result.CommandFindings))
	for _, command := range result.CommandFindings {
		commandKeys[commandFindingKey(command)] = true
	}
	for _, flow := range result.Findings {
		commandSink := isCommandExecutionSink(string(flow.Sink.Kind))
		if commandSink {
			if a.Config.Analyze.DetectUnsafeCommandConstruction {
				if commandKeys[commandFlowKey(flow)] {
					// The command-specific observation is richer than the generic
					// source-to-sink flow; project it once below.
					continue
				}
				findings = append(findings, a.commandFlowFinding(file, proc, flow))
				continue
			}
			// Preserve the legacy generic finding when the specialized rule is
			// explicitly disabled. With both rules enabled (the default), process
			// sinks are owned exclusively by VBA236.
			if !a.Config.Analyze.DetectUntrustedDataFlow {
				continue
			}
		}
		if !a.Config.Analyze.DetectUntrustedDataFlow {
			continue
		}
		line := flow.Sink.Range.StartLine
		if line <= 0 {
			line = flow.Source.Range.StartLine
		}
		if line <= 0 {
			line = proc.StartLine
		}
		sourceLine := flow.Source.Range.StartLine
		if sourceLine <= 0 {
			sourceLine = line
		}
		sinkLine := flow.Sink.Range.StartLine
		if sinkLine <= 0 {
			sinkLine = line
		}
		path := formatDataFlowPath(flow.Path)
		message := fmt.Sprintf("Conservative analysis: %s flows to %s. Source: %s; sink: %s; path: %s.", flow.Source.Label, flow.Sink.Label, flow.Source.Label, flow.Sink.Label, path)
		reason := fmt.Sprintf("Conservative analysis keeps potentially untrusted data from %s through the propagation path %s until it reaches the sensitive sink %s.", flow.Source.Label, path, flow.Sink.Label)
		suggestion := "Validate against a narrow allowlist or use the sink-specific sanitizer before calling the sensitive API."
		finding := a.simpleFinding(file, proc, line, "VBA224", "warning", message, reason, suggestion)
		finding.DataFlow = &DataFlowContext{
			Source: DataFlowEndpoint{Kind: string(flow.Source.Kind), Label: flow.Source.Label, Line: sourceLine},
			Sink:   DataFlowEndpoint{Kind: string(flow.Sink.Kind), Label: flow.Sink.Label, Line: sinkLine},
			Path:   convertDataFlowPath(flow.Path, line),
		}
		findings = append(findings, finding)
	}
	if a.Config.Analyze.DetectUnsafeCommandConstruction {
		for _, command := range result.CommandFindings {
			findings = append(findings, a.commandFinding(file, proc, command))
		}
	}
	return findings, nil
}

// isCommandExecutionSink is intentionally name-based rather than tied to a
// particular version of the protocol-neutral data-flow catalog. The catalog
// has historically exposed shell sinks as SinkShell and the adapter also
// accepts the more specific launcher names introduced for VBA236.
func isCommandExecutionSink(kind string) bool {
	switch strings.ToLower(strings.TrimSpace(kind)) {
	case "shell", "wscript_shell_run", "wscript_shell_exec", "shellexecute", "shell_execute", "cmd", "cmd_exe", "powershell", "script_host", "url_document_launch":
		return true
	default:
		return false
	}
}

func (a Analyzer) commandFlowFinding(file parsedFile, proc sourceProcedure, flow vbadf.Finding) Finding {
	line := flow.Sink.Range.StartLine
	if line <= 0 {
		line = flow.Source.Range.StartLine
	}
	if line <= 0 {
		line = proc.StartLine
	}
	sourceLine := flow.Source.Range.StartLine
	if sourceLine <= 0 {
		sourceLine = line
	}
	sinkLine := flow.Sink.Range.StartLine
	if sinkLine <= 0 {
		sinkLine = line
	}
	path := formatDataFlowPath(flow.Path)
	call, _ := commandCallForFlow(proc, flow)
	launcher, interpreter, role := commandLauncherDetails(flow, call, proc)
	class, kind := commandRiskClassification(flow, role)
	origin := string(flow.State)

	var message, reason, suggestion string
	if class == "injection" {
		message = fmt.Sprintf("Potential command-injection risk: %s reaches %s command text.", flow.Source.Label, flow.Sink.Label)
		reason = fmt.Sprintf("Known external input from %s reaches a process-launch sink through %s. This is a potential injection path, not a confirmed exploit.", flow.Source.Label, path)
	} else {
		message = fmt.Sprintf("Process-launch risk: input reaches %s, but its origin or command role is unknown.", flow.Sink.Label)
		reason = fmt.Sprintf("The value passed to %s crossed an unknown origin or transformation on the path %s; review the executable, arguments, and error handling before treating it as safe.", flow.Sink.Label, path)
	}
	suggestion = windowsCommandSuggestion(interpreter, role)
	finding := a.simpleFinding(file, proc, line, "VBA236", "warning", message, reason, suggestion)
	finding.DataFlow = &DataFlowContext{
		Source: DataFlowEndpoint{Kind: string(flow.Source.Kind), Label: flow.Source.Label, Line: sourceLine},
		Sink:   DataFlowEndpoint{Kind: string(flow.Sink.Kind), Label: flow.Sink.Label, Line: sinkLine},
		Path:   convertDataFlowPath(flow.Path, line),
	}
	finding.CommandExecution = &CommandExecutionContext{
		RiskClass: class, RiskKind: kind, Launcher: launcher,
		Interpreter: interpreter, CommandRole: role, OriginState: origin,
	}
	return finding
}

func (a Analyzer) commandFinding(file parsedFile, proc sourceProcedure, command vbadf.CommandFinding) Finding {
	execution := command.Execution
	if execution.Interpreter == "" {
		execution.Interpreter = commandInterpreterForProcedure(proc, execution)
	}
	line := execution.Range.StartLine
	if line <= 0 {
		line = command.Source.Range.StartLine
	}
	if line <= 0 {
		line = proc.StartLine
	}
	class := string(command.RiskClass)
	if class == "" {
		class = "process_launch"
	}
	kind := string(command.RiskKind)
	if kind == "" {
		kind = "unknown_origin"
	}
	message, reason := command.Message, command.Reason
	if message == "" {
		switch kind {
		case "unquoted_executable_path":
			message = fmt.Sprintf("Unquoted executable path in %s may launch the wrong program.", execution.Launcher)
			reason = "Windows command parsing treats an executable path containing spaces as multiple tokens unless the complete path is quoted."
		case "credential_exposure":
			message = fmt.Sprintf("Command-line credential exposure risk in %s.", execution.Launcher)
			reason = "Secrets passed as process arguments can be exposed through process listings, diagnostics, or child-process inspection."
		case "observability":
			message = fmt.Sprintf("Process launched by %s without observable completion handling.", execution.Launcher)
			reason = "Hidden or unobserved execution makes failures and exit status unavailable to the caller."
		case "tainted_command_text":
			message = fmt.Sprintf("Potential command-injection risk: input reaches %s.", execution.Launcher)
			reason = "Known external input reaches a command-bearing process-launch role; this is a potential injection path, not a confirmed exploit."
		default:
			message = fmt.Sprintf("Process-launch risk: %s requires review.", execution.Launcher)
			reason = "The command origin or role is not precise enough to claim injection; review executable quoting, argument validation, and error handling."
		}
	}
	// Core observations intentionally use protocol-neutral wording. Prefix
	// the analyzer projection with the stable risk class so CLI and LSP users
	// can distinguish potential injection from a general launch warning even
	// when the core supplied its own message.
	if class == "injection" && !strings.Contains(strings.ToLower(message), "injection") {
		message = "Potential command-injection risk: " + message
	}
	if kind == "unquoted_executable_path" && !strings.Contains(strings.ToLower(message), "unquoted") {
		message = "Unquoted executable path risk: " + message
	}
	if kind == "credential_exposure" && !strings.Contains(strings.ToLower(message), "credential") {
		message = "Command-line credential exposure risk: " + message
	}
	if kind == "observability" && !strings.Contains(strings.ToLower(message), "observ") {
		message = "Unobserved process launch risk: " + message
	}
	suggestion := windowsCommandSuggestion(execution.Interpreter, string(execution.Role))
	if kind == "credential_exposure" {
		suggestion = "Do not put secrets on the command line; use environment variables, standard input, or a protected credential store and redact diagnostics."
	}
	finding := a.simpleFinding(file, proc, line, "VBA236", "warning", message, reason, suggestion)
	if kind == "credential_exposure" {
		finding.NearbyCode = redactedNearbyCode(file.Lines, line, line, 2)
	}
	if command.Source.Kind != "" || command.Source.Label != "" {
		sourceLine := command.Source.Range.StartLine
		if sourceLine <= 0 {
			sourceLine = line
		}
		sinkLine := execution.Range.StartLine
		if sinkLine <= 0 {
			sinkLine = line
		}
		finding.DataFlow = &DataFlowContext{
			Source: DataFlowEndpoint{Kind: string(command.Source.Kind), Label: command.Source.Label, Line: sourceLine},
			Sink:   DataFlowEndpoint{Kind: commandSinkKind(execution.Launcher), Label: execution.Launcher, Line: sinkLine},
			Path:   convertDataFlowPath(command.Path, line),
		}
	}
	finding.CommandExecution = &CommandExecutionContext{
		RiskClass:   class,
		RiskKind:    kind,
		Launcher:    execution.Launcher,
		Interpreter: execution.Interpreter,
		CommandRole: string(execution.Role),
		OriginState: string(command.OriginState),
	}
	return finding
}

// commandInterpreterForProcedure repairs older core observations that did not
// retain the interpreter token on CommandExecution. The procedure IR still
// carries the argument expression, so the adapter can recover cmd/PowerShell
// guidance without exposing parser nodes in the public finding.
func commandInterpreterForProcedure(proc sourceProcedure, execution vbadf.CommandExecution) string {
	for _, call := range proc.Calls {
		if call.Range.StartByte != execution.Range.StartByte || execution.Argument < 0 || execution.Argument >= len(call.Arguments.ExpressionIDs) {
			continue
		}
		expressionID := call.Arguments.ExpressionIDs[execution.Argument]
		for _, expression := range proc.Expressions {
			if expression.ID == expressionID {
				return commandInterpreter(expression.Text)
			}
		}
	}
	return ""
}

func commandSinkKind(launcher string) string {
	lower := strings.ToLower(strings.TrimSpace(launcher))
	switch {
	case strings.Contains(lower, "wscript.shell") && strings.Contains(lower, ".run"):
		return "wscript_shell_run"
	case strings.Contains(lower, "wscript.shell") && strings.Contains(lower, ".exec"):
		return "wscript_shell_exec"
	case strings.Contains(lower, "shellexecute"):
		return "shell_execute"
	case strings.HasPrefix(lower, "shell") || lower == "shell":
		return "shell"
	default:
		return "command_execution"
	}
}

func commandFlowKey(flow vbadf.Finding) string {
	return fmt.Sprintf("%d:%s:%s", flow.Sink.Range.StartByte, flow.Source.Kind, strings.ToLower(flow.Source.Label))
}

func commandFindingKey(command vbadf.CommandFinding) string {
	return fmt.Sprintf("%d:%s:%s", command.Execution.Range.StartByte, command.Source.Kind, strings.ToLower(command.Source.Label))
}

func commandRiskClassification(flow vbadf.Finding, role string) (string, string) {
	if flow.State == vbadf.StateTainted && (role == string(vbadf.CommandRoleExecutable) || role == string(vbadf.CommandRoleShellCommand)) {
		return "injection", "tainted_command_text"
	}
	return "process_launch", "unknown_origin"
}

func commandCallForFlow(proc sourceProcedure, flow vbadf.Finding) (procedureir.CallSite, bool) {
	for _, call := range proc.Calls {
		if call.StatementID == flow.Sink.StatementID && call.Range.StartByte == flow.Sink.Range.StartByte {
			return call, true
		}
	}
	for _, call := range proc.Calls {
		if call.StatementID == flow.Sink.StatementID {
			return call, true
		}
	}
	return procedureir.CallSite{}, false
}

func commandLauncherDetails(flow vbadf.Finding, call procedureir.CallSite, proc sourceProcedure) (launcher, interpreter, role string) {
	launcher = strings.TrimSpace(flow.Sink.Label)
	if call.Callee.Text != "" {
		launcher = strings.TrimSpace(call.Callee.Text)
	}
	role = string(vbadf.CommandRoleUnknown)
	command := ""
	if len(call.Arguments.ExpressionIDs) > 0 {
		for _, expressionID := range call.Arguments.ExpressionIDs {
			for _, expression := range proc.Expressions {
				if expression.ID == expressionID {
					command = expression.Text
					break
				}
			}
			if command != "" {
				break
			}
		}
	}
	if command == "" && call.Arguments.Count > 0 {
		// The sink range/statement still gives us a useful launcher label when
		// the parser recovered the argument list without expression IDs.
		command = launcher
	}
	interpreter = commandInterpreter(command)
	return launcher, interpreter, role
}

func commandInterpreter(command string) string {
	command = strings.ToLower(strings.TrimSpace(command))
	switch {
	case strings.Contains(command, "cmd.exe") || strings.HasPrefix(command, "cmd ") || strings.Contains(command, "cmd /c"):
		return "cmd.exe"
	case strings.Contains(command, "powershell.exe") || strings.Contains(command, "pwsh") || strings.Contains(command, "powershell "):
		return "powershell"
	case strings.Contains(command, "wscript.exe") || strings.Contains(command, "cscript.exe") || strings.Contains(command, "mshta.exe"):
		return "script_host"
	default:
		return ""
	}
}

func windowsCommandSuggestion(interpreter, role string) string {
	if interpreter == "cmd.exe" {
		return "Quote the executable path and keep arguments separate; avoid concatenating external input into `cmd.exe /c`, and escape `& | < > ( ) ^ % !` when a shell is unavoidable."
	}
	if interpreter == "powershell" {
		return "Prefer `powershell -File` with fixed, separately validated arguments over `-Command`; quote the executable path and keep secrets out of the command line."
	}
	if interpreter == "script_host" {
		return "Use a fixed script path and validated arguments; quote paths containing spaces and avoid passing external input as script-host command text."
	}
	if role == "command_text" {
		return "Use a fixed executable path, quote it for Windows, validate arguments against a narrow allowlist, and avoid putting secrets on the command line."
	}
	return "Use a fixed executable path, quote it for Windows, validate arguments against a narrow allowlist, and inspect the process result or exit code."
}

func dataFlowBindings(declarations []procedureir.Declaration) map[string]bool {
	bindings := make(map[string]bool, len(declarations))
	for _, declaration := range declarations {
		name := strings.ToLower(strings.TrimSpace(declaration.Name))
		if name == "" {
			continue
		}
		bindings[name] = declaration.Kind == "const"
	}
	return bindings
}

func (a Analyzer) isKnownDataFlowConstant(file parsedFile, name string) bool {
	key := strings.ToLower(strings.TrimSpace(name))
	if constant, bound := file.DataFlowModuleBindings[key]; bound {
		return constant
	}
	if a.typeDB == nil {
		return false
	}
	_, ok := a.typeDB.ResolveConstant(name)
	return ok
}

func procedureIRForSource(document procedureir.DocumentIR, proc sourceProcedure) (procedureir.ProcedureIR, bool) {
	for _, candidate := range document.Procedures {
		if !strings.EqualFold(candidate.Symbol.Name, proc.Name) {
			continue
		}
		if proc.StartLine > 0 && candidate.Symbol.DeclarationRange.StartLine != proc.StartLine {
			continue
		}
		return candidate, true
	}
	return procedureir.ProcedureIR{}, false
}

func convertDataFlowPath(path []vbadf.PathStep, fallbackLine int) []DataFlowStep {
	steps := make([]DataFlowStep, 0, len(path))
	for _, step := range path {
		line := step.Range.StartLine
		if line <= 0 {
			line = fallbackLine
		}
		steps = append(steps, DataFlowStep{Kind: step.Kind, Label: step.Label, Line: line})
	}
	return steps
}

func formatDataFlowPath(path []vbadf.PathStep) string {
	if len(path) == 0 {
		return "source to sink"
	}
	labels := make([]string, 0, len(path))
	for _, step := range path {
		label := strings.TrimSpace(step.Label)
		if label == "" {
			label = step.Kind
		}
		labels = append(labels, label)
	}
	return strings.Join(labels, " -> ")
}
