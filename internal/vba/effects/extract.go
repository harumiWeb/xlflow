package effects

import (
	"regexp"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

var whitespace = regexp.MustCompile(`\s+`)

func extractStatements(summary *ProcedureSummary, proc procedureir.ProcedureIR, reachable map[int]bool) {
	for index, statement := range proc.Statements {
		if !reachable[statement.ID] || statement.Recovered {
			continue
		}
		// The parser represents a chained parameterized receiver and its final
		// no-argument member as adjacent statements. Join only the exact same
		// physical expression; never infer a receiver from another line.
		if strings.HasPrefix(strings.TrimSpace(statement.Text), ".") && index > 0 {
			previous := proc.Statements[index-1]
			member := strings.ToLower(strings.TrimPrefix(compact(statement.Text), "."))
			if reachable[previous.ID] && !previous.Recovered && previous.Range.EndByte == statement.Range.StartByte &&
				previous.Range.EndLine == statement.Range.StartLine && isCellSurface(strings.ToLower(compact(previous.Text))) &&
				(member == "clear" || member == "clearcontents" || member == "insert" || member == "delete") {
				addStatementEffect(summary, statement, WritesCells, previous.Text+statement.Text, "")
				addStatementEffect(summary, statement, ChangesWorkbook, previous.Text+statement.Text, "")
			}
		}
		if statement.Kind == procedureir.StatementOnError && statement.Control != nil && statement.Control.Transfer == procedureir.TransferOnErrorResumeNext {
			addStatementEffect(summary, statement, SuppressesErrors, "On Error", "Resume Next")
		}
		if statement.Kind != procedureir.StatementAssignment || statement.Target == nil {
			continue
		}
		target := compact(statement.Target.Text)
		value := ""
		if statement.Value != nil {
			value = strings.TrimSpace(statement.Value.Text)
		}
		lower := strings.ToLower(target)
		switch {
		case isCellSurface(lower):
			addStatementEffect(summary, statement, WritesCells, target, value)
			addStatementEffect(summary, statement, ChangesWorkbook, target, value)
		case applicationStateTarget(lower):
			targetName := applicationStateTargetName(target)
			addStatementEffect(summary, statement, ChangesApplicationState, targetName, value)
			if strings.EqualFold(lower, "application.enableevents") && (strings.EqualFold(strings.TrimSpace(value), "false") || strings.TrimSpace(value) == "0") {
				addStatementEffect(summary, statement, DisablesEvents, targetName, value)
			}
			if strings.EqualFold(lower, "application.calculation") {
				addStatementEffect(summary, statement, ChangesCalculation, targetName, value)
			}
			if applicationStateRestoreCandidate(lower, value) {
				if strings.EqualFold(lower, "application.enableevents") {
					addStatementEffect(summary, statement, RestoresEvents, targetName, value)
				}
				addStatementEffect(summary, statement, RestoresApplicationState, targetName, value)
			}
		}
	}
}

func applicationStateTarget(target string) bool {
	switch target {
	case "application.enableevents", "application.displayalerts", "application.screenupdating", "application.calculation",
		"application.statusbar", "application.cursor", "application.interactive", "application.asktoupdatelinks",
		"application.automationsecurity", "application.cutcopymode":
		return true
	default:
		return false
	}
}

func applicationStateTargetName(target string) string {
	if len(target) <= len("application.") {
		return target
	}
	name := target[len("application."):]
	switch name {
	case "enableevents":
		return "Application.EnableEvents"
	case "displayalerts":
		return "Application.DisplayAlerts"
	case "screenupdating":
		return "Application.ScreenUpdating"
	case "calculation":
		return "Application.Calculation"
	case "statusbar":
		return "Application.StatusBar"
	case "cursor":
		return "Application.Cursor"
	case "interactive":
		return "Application.Interactive"
	case "asktoupdatelinks":
		return "Application.AskToUpdateLinks"
	case "automationsecurity":
		return "Application.AutomationSecurity"
	case "cutcopymode":
		return "Application.CutCopyMode"
	default:
		return target
	}
}

// Restore evidence remains intentionally broader than the CFG proof used by
// VBA203. It exists only for the established Push/Pop helper convention,
// where a stored value or an Excel default is an obvious cleanup candidate.
func applicationStateRestoreCandidate(target, value string) bool {
	value = strings.ToLower(strings.TrimSpace(value))
	if value == "true" || value == "false" || value == "0" || value == "xldefault" ||
		value == "xlcalculationautomatic" || value == "msoautomationsecuritybyui" {
		return true
	}
	if strings.ContainsAny(value, ".() +-*/&=<>:") {
		return false
	}
	return value != ""
}

func extractCall(summary *ProcedureSummary, call procedureir.CallSite, statement procedureir.Statement) {
	// A matched call names project code even when it happens to share a name
	// with a VBA/Excel builtin.
	if call.Resolution.Status == procedureir.ResolutionMatched {
		return
	}
	full := strings.ToLower(compact(call.Callee.Text))
	base := strings.ToLower(call.Callee.BaseName)
	member := strings.ToLower(call.Callee.Member)
	if member == "" {
		member = base
	}
	receiver := ""
	if call.Callee.Receiver != nil {
		receiver = strings.ToLower(compact(*call.Callee.Receiver))
		if !strings.Contains(full, receiver) {
			full = receiver + "." + member
		}
	}
	builtin := call.Resolution.Status == procedureir.ResolutionBuiltinLike
	switch {
	case builtin && receiver == "" && (base == "msgbox" || base == "inputbox") || receiver == "application" && member == "inputbox":
		addCallEffect(summary, call, ShowsDialog, call.Callee.Text)
	case builtin && receiver == "" && base == "shell" || isWScriptShellReceiver(receiver) && (member == "run" || member == "exec"):
		addCallEffect(summary, call, LaunchesProcess, call.Callee.Text)
	case member == "open" && (receiver == "workbooks" || receiver == "application.workbooks"):
		addCallEffect(summary, call, OpensWorkbook, call.Callee.Text)
	case member == "close" && isWorkbookReceiver(receiver):
		addCallEffect(summary, call, ClosesWorkbook, call.Callee.Text)
	case builtin && receiver == "" && base == "error" && statement.Kind == procedureir.StatementCall ||
		(member == "raise" && receiver == "err"):
		addCallEffect(summary, call, RaisesError, call.Callee.Text)
	}
	if isWorkbookMutation(full, member) {
		addCallEffect(summary, call, ChangesWorkbook, call.Callee.Text)
	}
	if isCellMutation(full, member) {
		addCallEffect(summary, call, WritesCells, call.Callee.Text)
		addCallEffect(summary, call, ChangesWorkbook, call.Callee.Text)
	}
}

func addStatementEffect(summary *ProcedureSummary, statement procedureir.Statement, kind EffectKind, target, value string) {
	summary.Direct = append(summary.Direct, Evidence{Effect: kind, Origin: summary.Identity, Range: statement.Range, StatementID: statement.ID, Target: target, Value: value})
}

func addCallEffect(summary *ProcedureSummary, call procedureir.CallSite, kind EffectKind, target string) {
	summary.Direct = append(summary.Direct, Evidence{Effect: kind, Origin: summary.Identity, Range: call.Range, StatementID: call.StatementID, CallID: call.ID, Target: target})
}

func compact(value string) string { return whitespace.ReplaceAllString(strings.TrimSpace(value), "") }

func isCellSurface(lower string) bool {
	if strings.HasPrefix(lower, "range(") || strings.HasPrefix(lower, "cells(") {
		return true
	}
	for _, prefix := range []string{"activesheet.range(", "activesheet.cells(", "application.range(", "application.cells(", "thisworkbook.worksheets(", "thisworkbook.sheets(", "worksheets(", "sheets("} {
		if strings.HasPrefix(lower, prefix) && (strings.Contains(lower, ".range(") || strings.Contains(lower, ".cells(")) {
			return true
		}
	}
	return false
}

func isCellMutation(full, member string) bool {
	if member != "clear" && member != "clearcontents" && member != "insert" && member != "delete" {
		return false
	}
	return isCellSurface(full) || strings.HasPrefix(full, "rows(") || strings.HasPrefix(full, "columns(")
}

func isWorkbookMutation(full, member string) bool {
	if member == "save" || member == "saveas" || member == "savecopyas" {
		receiver := strings.TrimSuffix(full, "."+member)
		return isWorkbookReceiver(receiver)
	}
	if member == "add" || member == "delete" || member == "move" || member == "copy" {
		receiver := strings.TrimSuffix(full, "."+member)
		return receiver == "worksheets" || receiver == "sheets" || receiver == "thisworkbook.worksheets" || receiver == "thisworkbook.sheets" || receiver == "activeworkbook.worksheets" || receiver == "activeworkbook.sheets" || strings.HasPrefix(receiver, "worksheets(") || strings.HasPrefix(receiver, "sheets(") || strings.HasPrefix(receiver, "thisworkbook.worksheets(") || strings.HasPrefix(receiver, "thisworkbook.sheets(") || strings.HasPrefix(receiver, "activeworkbook.worksheets(") || strings.HasPrefix(receiver, "activeworkbook.sheets(")
	}
	return false
}

func isWorkbookReceiver(receiver string) bool {
	return receiver == "thisworkbook" || receiver == "activeworkbook" || receiver == "application.activeworkbook" || strings.HasPrefix(receiver, "workbooks(")
}

func isWScriptShellReceiver(receiver string) bool {
	return receiver == "wscript.shell" || strings.HasPrefix(receiver, `createobject("wscript.shell")`)
}
