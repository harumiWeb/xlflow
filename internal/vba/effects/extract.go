package effects

import (
	"regexp"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

var (
	whitespace           = regexp.MustCompile(`\s+`)
	bareVBAIdentifierRE  = regexp.MustCompile(`^[a-z_][a-z0-9_]*$`)
	formControlTargetRE  = regexp.MustCompile(`^(?:me\.)?([a-z_][a-z0-9_]*)\.(?:value|text|listindex|list)$`)
	formControlsTargetRE = regexp.MustCompile(`^(?:me\.)?controls\((?:"[^"]+"|'[^']+')\)\.(?:value|text|listindex|list)$`)
	fileOpenEffectRE     = regexp.MustCompile(`(?i)^\s*open\b.*\bfor\s+(?:append|binary|input|output|random)\b.*\bas\s+#?\s*(?:\d+|\[[^\]]+\]|[a-z_][a-z0-9_]*(?:[$%&!#@^])?)\b`)
)

func extractStatements(summary *ProcedureSummary, proc procedureir.ProcedureIR, reachable map[int]bool) {
	statementsByID := make(map[int]procedureir.Statement, len(proc.Statements))
	for _, statement := range proc.Statements {
		statementsByID[statement.ID] = statement
	}
	for index, statement := range proc.Statements {
		if !statementReachable(statement, reachable, statementsByID) || statement.Recovered {
			continue
		}
		// The parser represents a chained parameterized receiver and its final
		// no-argument member as adjacent statements. Join only the exact same
		// physical expression; never infer a receiver from another line.
		if strings.HasPrefix(strings.TrimSpace(statement.Text), ".") && index > 0 {
			previous := proc.Statements[index-1]
			member := strings.ToLower(strings.TrimPrefix(compact(statement.Text), "."))
			if statementReachable(previous, reachable, statementsByID) && !previous.Recovered && previous.Range.EndByte == statement.Range.StartByte &&
				previous.Range.EndLine == statement.Range.StartLine && isCellSurface(strings.ToLower(compact(previous.Text))) &&
				(member == "clear" || member == "clearcontents" || member == "insert" || member == "delete") {
				addStatementEffect(summary, statement, WritesCells, previous.Text+statement.Text, "")
			}
		}
		if statement.Kind == procedureir.StatementOnError && statement.Control != nil && statement.Control.Transfer == procedureir.TransferOnErrorResumeNext {
			addStatementEffect(summary, statement, SuppressesErrors, "On Error", "Resume Next")
		}
		if fileOpenEffectRE.MatchString(statement.Text) {
			addStatementEffect(summary, statement, OpensFile, "VBA file handle", "")
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
		property, applicationState := applicationStateTarget(statement, lower, statementsByID)
		switch {
		case isCellTarget(proc, lower):
			addStatementEffect(summary, statement, WritesCells, target, value)
		case isSheetNameTarget(lower):
			addStatementEffect(summary, statement, ChangesWorkbook, target, value)
		case controlTarget(summary.Identity.ModuleKind, lower):
			addStatementEffect(summary, statement, ChangesControls, target, value)
		case applicationState:
			targetName := applicationStateTargetName("application." + property)
			addStatementEffect(summary, statement, ChangesApplicationState, targetName, value)
			if property == "enableevents" && (strings.EqualFold(strings.TrimSpace(value), "false") || strings.TrimSpace(value) == "0") {
				addStatementEffect(summary, statement, DisablesEvents, targetName, value)
			}
			if property == "calculation" {
				addStatementEffect(summary, statement, ChangesCalculation, targetName, value)
			}
			if applicationStateRestoreCandidate(proc, statement, property, value) {
				if property == "enableevents" {
					addStatementEffect(summary, statement, RestoresEvents, targetName, value)
				}
				addStatementEffect(summary, statement, RestoresApplicationState, targetName, value)
			}
		}
	}
}

func statementReachable(statement procedureir.Statement, reachable map[int]bool, statementsByID map[int]procedureir.Statement) bool {
	for {
		if reachable[statement.ID] {
			return true
		}
		if statement.ParentID == 0 {
			return false
		}
		parent, ok := statementsByID[statement.ParentID]
		if !ok {
			return false
		}
		statement = parent
	}
}

func applicationStateTarget(statement procedureir.Statement, target string, statementsByID map[int]procedureir.Statement) (string, bool) {
	property := strings.TrimPrefix(target, "application.")
	if property == target {
		if !strings.HasPrefix(target, ".") || !statementWithinApplicationWith(statement, statementsByID) {
			return "", false
		}
		property = strings.TrimPrefix(target, ".")
	}
	switch property {
	case "enableevents", "displayalerts", "screenupdating", "calculation", "statusbar", "cursor", "interactive", "asktoupdatelinks", "automationsecurity", "cutcopymode":
		return property, true
	default:
		return "", false
	}
}

func statementWithinApplicationWith(statement procedureir.Statement, statementsByID map[int]procedureir.Statement) bool {
	for parentID := statement.ParentID; parentID != 0; {
		parent, ok := statementsByID[parentID]
		if !ok {
			return false
		}
		if parent.Kind == procedureir.StatementWith {
			return parent.Value != nil && strings.EqualFold(compact(parent.Value.Text), "application")
		}
		parentID = parent.ParentID
	}
	return false
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

// Restore evidence remains intentionally narrower than the CFG proof used by
// VBA203. It exists only for the established Push/Pop helper convention and
// accepts a known property reset or a resolved saved-value variable. In
// particular, a disabling assignment must never claim restoration evidence.
func applicationStateRestoreCandidate(proc procedureir.ProcedureIR, statement procedureir.Statement, property, value string) bool {
	normalized := strings.ToLower(strings.TrimSpace(value))
	if knownApplicationStateReset(property, normalized) {
		return true
	}
	if !bareVBAIdentifier(normalized) {
		return false
	}
	for _, access := range proc.Accesses {
		if access.StatementID == statement.ID && access.Mode != procedureir.AccessWrite &&
			access.Scope != procedureir.ScopeUnresolved && strings.EqualFold(access.Name, normalized) {
			return true
		}
	}
	return false
}

func knownApplicationStateReset(property, value string) bool {
	switch property {
	case "enableevents", "displayalerts", "screenupdating", "interactive", "asktoupdatelinks":
		return value == "true"
	case "calculation":
		return value == "xlcalculationautomatic"
	case "statusbar", "cutcopymode":
		return value == "false" || value == "0"
	case "cursor":
		return value == "xldefault"
	case "automationsecurity":
		return value == "msoautomationsecuritybyui"
	default:
		return false
	}
}

func bareVBAIdentifier(value string) bool {
	return bareVBAIdentifierRE.MatchString(value)
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
	}
	if member == "calculate" && (isCalculationReceiver(receiver) || builtin && receiver == "") {
		addCallEffect(summary, call, Recalculates, call.Callee.Text)
	}
	if (member == "select" || member == "activate" || member == "goto") && isSelectionReceiver(receiver) {
		addCallEffect(summary, call, ChangesSelection, call.Callee.Text)
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
	for _, prefix := range []string{"me.range(", "me.cells(", "activesheet.range(", "activesheet.cells(", "application.range(", "application.cells(", "thisworkbook.worksheets(", "thisworkbook.sheets(", "worksheets(", "sheets("} {
		if strings.HasPrefix(lower, prefix) && (strings.Contains(lower, ".range(") || strings.Contains(lower, ".cells(")) {
			return true
		}
	}
	return false
}

func isCellTarget(proc procedureir.ProcedureIR, target string) bool {
	if isCellSurface(target) {
		return true
	}
	root := target
	if index := strings.IndexByte(root, '.'); index >= 0 {
		root = root[:index]
	}
	root = strings.TrimSuffix(root, "()")
	for _, parameter := range proc.Symbol.Parameters {
		if strings.EqualFold(parameter.Name, root) && strings.EqualFold(strings.TrimSpace(parameter.Type), "Range") {
			return true
		}
	}
	for _, declaration := range proc.Declarations {
		if strings.EqualFold(declaration.Name, root) && strings.EqualFold(strings.TrimSpace(declaration.Type), "Range") {
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

func isSheetNameTarget(target string) bool {
	return strings.HasSuffix(target, ".name") && (strings.Contains(target, "worksheets(") || strings.Contains(target, "sheets(") || strings.HasPrefix(target, "activesheet.") || strings.HasPrefix(target, "thisworkbook.worksheets"))
}

func controlTarget(moduleKind, target string) bool {
	if !strings.EqualFold(moduleKind, "form") {
		return false
	}
	return formControlTargetRE.MatchString(target) || formControlsTargetRE.MatchString(target)
}

func isCalculationReceiver(receiver string) bool {
	return receiver == "application" || receiver == "thisworkbook" || receiver == "activeworkbook" || receiver == "activesheet" ||
		strings.Contains(receiver, ".range(") || strings.HasPrefix(receiver, "range(") || strings.HasPrefix(receiver, "cells(")
}

func isSelectionReceiver(receiver string) bool {
	return receiver == "application" || receiver == "activesheet" || receiver == "activeworkbook" || receiver == "thisworkbook" ||
		receiver == "selection" || receiver == "activecell" || strings.HasPrefix(receiver, "workbooks(") ||
		strings.HasPrefix(receiver, "worksheets(") || strings.HasPrefix(receiver, "sheets(") ||
		strings.Contains(receiver, ".worksheets(") || strings.Contains(receiver, ".sheets(") ||
		strings.Contains(receiver, ".range(") || strings.HasPrefix(receiver, "range(") || strings.HasPrefix(receiver, "cells(") ||
		strings.Contains(receiver, ".cells(")
}

func isWorkbookReceiver(receiver string) bool {
	return receiver == "thisworkbook" || receiver == "activeworkbook" || receiver == "application.activeworkbook" || strings.HasPrefix(receiver, "workbooks(")
}

func isWScriptShellReceiver(receiver string) bool {
	return receiver == "wscript.shell" || strings.HasPrefix(receiver, `createobject("wscript.shell")`)
}
