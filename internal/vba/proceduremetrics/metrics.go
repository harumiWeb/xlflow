// Package proceduremetrics computes deterministic, protocol-neutral
// maintainability metrics for one VBA procedure.  The package deliberately
// does not depend on analyzer findings or rule enablement; callers can use
// the resulting values for reporting, thresholds, or other future consumers.
package proceduremetrics

import (
	"fmt"
	"sort"
	"strings"

	vbast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// JSONSchemaVersion is the version of the public procedure metrics object.
// Adding fields is backward compatible; removing a field or changing its
// meaning requires incrementing this value.
const JSONSchemaVersion = 1

// MetricName is the stable identifier used by configuration and threshold
// diagnostics.  Metric names are intentionally snake_case so they are also
// suitable as JSON and TOML keys.
type MetricName string

const (
	MetricCyclomaticComplexity MetricName = "cyclomatic_complexity"
	MetricMaxNestingDepth      MetricName = "max_nesting_depth"
	MetricStatementCount       MetricName = "statement_count"
	MetricSourceLineCount      MetricName = "source_line_count"
	MetricBranchCount          MetricName = "branch_count"
	MetricLoopCount            MetricName = "loop_count"
	MetricGotoCount            MetricName = "goto_count"
	MetricExitPointCount       MetricName = "exit_point_count"
	MetricParameterCount       MetricName = "parameter_count"
	MetricByRefParameterCount  MetricName = "byref_parameter_count"
	MetricLocalVariableCount   MetricName = "local_variable_count"
	MetricCallFanOut           MetricName = "call_fan_out"
)

var canonicalMetricNames = [...]MetricName{
	MetricCyclomaticComplexity,
	MetricMaxNestingDepth,
	MetricStatementCount,
	MetricSourceLineCount,
	MetricBranchCount,
	MetricLoopCount,
	MetricGotoCount,
	MetricExitPointCount,
	MetricParameterCount,
	MetricByRefParameterCount,
	MetricLocalVariableCount,
	MetricCallFanOut,
}

// MetricNames is a copy of the canonical order used for display and
// configuration help.  Evaluation uses a private immutable array so a caller
// mutating this compatibility slice cannot make results non-deterministic.
var MetricNames = append([]MetricName(nil), canonicalMetricNames[:]...)

// Metrics contains all procedure metrics. The fields are kept as
// integers because every metric is a count in schema version 1.
type Metrics struct {
	CyclomaticComplexity            int `json:"cyclomatic_complexity"`
	MaxNestingDepth                 int `json:"max_nesting_depth"`
	StatementCount                  int `json:"statement_count"`
	SourceLineCount                 int `json:"source_line_count"`
	BranchCount                     int `json:"branch_count"`
	LoopCount                       int `json:"loop_count"`
	GotoCount                       int `json:"goto_count"`
	ExitPointCount                  int `json:"exit_point_count"`
	ParameterCount                  int `json:"parameter_count"`
	ByRefParameterCount             int `json:"byref_parameter_count"`
	LocalVariableCount              int `json:"local_variable_count"`
	CallFanOut                      int `json:"call_fan_out"`
	BooleanParameterCount           int `json:"boolean_parameter_count"`
	OptionalBooleanParameterCount   int `json:"optional_boolean_parameter_count"`
	VagueBooleanParameterCount      int `json:"vague_boolean_parameter_count"`
	BooleanControlBranchCount       int `json:"boolean_control_branch_count"`
	BooleanControlledStatementCount int `json:"boolean_controlled_statement_count"`
}

// Value returns a metric by its stable name.  The bool is false for an
// unknown name, allowing configuration validation to reject typos.
func (m Metrics) Value(name MetricName) (int, bool) {
	switch name {
	case MetricCyclomaticComplexity:
		return m.CyclomaticComplexity, true
	case MetricMaxNestingDepth:
		return m.MaxNestingDepth, true
	case MetricStatementCount:
		return m.StatementCount, true
	case MetricSourceLineCount:
		return m.SourceLineCount, true
	case MetricBranchCount:
		return m.BranchCount, true
	case MetricLoopCount:
		return m.LoopCount, true
	case MetricGotoCount:
		return m.GotoCount, true
	case MetricExitPointCount:
		return m.ExitPointCount, true
	case MetricParameterCount:
		return m.ParameterCount, true
	case MetricByRefParameterCount:
		return m.ByRefParameterCount, true
	case MetricLocalVariableCount:
		return m.LocalVariableCount, true
	case MetricCallFanOut:
		return m.CallFanOut, true
	default:
		return 0, false
	}
}

// ProcedureMetrics combines the procedure identity and its metrics.  The
// anonymous Metrics field intentionally flattens the metric values in JSON,
// while keeping identity available to threshold/reporting consumers.
type ProcedureMetrics struct {
	File                    string                    `json:"file"`
	Module                  string                    `json:"module"`
	ModuleKind              string                    `json:"module_kind"`
	Name                    string                    `json:"name"`
	Kind                    procedureir.ProcedureKind `json:"kind"`
	Visibility              string                    `json:"visibility,omitempty"`
	ResolvedCallees         []string                  `json:"-"`
	ExternalDependencyCount int                       `json:"-"`
	ExcelEffectCount        int                       `json:"-"`
	MutableStateReads       int                       `json:"-"`
	MutableStateWrites      int                       `json:"-"`
	MutableStateMutations   int                       `json:"-"`
	ErrorHandlingCount      int                       `json:"-"`
	ResourceOwnershipCount  int                       `json:"-"`
	AmbiguousCallCount      int                       `json:"-"`
	UnresolvedCallCount     int                       `json:"-"`
	DynamicCallCount        int                       `json:"-"`
	DeclarationRange        vbast.Range               `json:"declaration_range"`
	Metrics
}

// Input is the protocol-neutral input to Collect.  IR is authoritative when
// present.  Graph is accepted so callers that already built a CFG need not
// discard it; when IR statements are absent, statement blocks from Graph are
// used as a conservative fallback.  Calls overrides IR.Calls when non-nil.
type Input struct {
	IR         procedureir.ProcedureIR
	Graph      *cfg.Graph
	Calls      []procedureir.CallSite
	File       string
	Module     string
	ModuleKind string
}

// Collect computes all metrics for one procedure.  It never mutates input
// slices and does not inspect analyzer findings.
func Collect(input Input) ProcedureMetrics {
	procedure := input.IR
	symbol := procedure.Symbol
	if symbol.Name == "" && input.Graph != nil {
		symbol = input.Graph.Procedure
	}
	statements := orderedStatements(procedure.Statements)
	if len(statements) == 0 && input.Graph != nil {
		statements = statementsFromGraph(*input.Graph)
	}
	calls := input.Calls
	if calls == nil {
		calls = procedure.Calls
	}

	result := ProcedureMetrics{
		File:             input.File,
		Module:           input.Module,
		ModuleKind:       input.ModuleKind,
		Name:             symbol.Name,
		Kind:             symbol.Kind,
		Visibility:       symbol.Visibility,
		DeclarationRange: symbol.DeclarationRange,
	}
	result.SourceLineCount = sourceLineCount(symbol.DeclarationRange)
	result.StatementCount = statementCount(statements)
	result.BranchCount = branchCount(statements)
	result.LoopCount = loopCount(statements)
	result.CyclomaticComplexity = 1 + result.BranchCount + result.LoopCount
	result.MaxNestingDepth = maxNestingDepth(statements)
	result.GotoCount = gotoCount(statements)
	result.ExitPointCount = exitPointCount(symbol.Kind, statements)
	result.ParameterCount = len(symbol.Parameters)
	for _, parameter := range symbol.Parameters {
		// VBA's default parameter passing mode is ByRef.  Count that effective
		// mode even when the source omitted the ByRef keyword.
		if !strings.EqualFold(strings.TrimSpace(parameter.Passing), "ByVal") {
			result.ByRefParameterCount++
		}
	}
	for _, declaration := range procedure.Declarations {
		kind := strings.ToLower(strings.TrimSpace(declaration.Kind))
		if declaration.Scope == procedureir.ScopeLocal &&
			kind != "return_slot" && kind != "parameter" {
			result.LocalVariableCount++
		}
	}
	result.CallFanOut = callFanOut(calls)
	booleanMetrics := booleanControlMetrics(symbol.Parameters, statements)
	result.BooleanParameterCount = booleanMetrics.booleanParameterCount
	result.OptionalBooleanParameterCount = booleanMetrics.optionalBooleanParameterCount
	result.VagueBooleanParameterCount = booleanMetrics.vagueBooleanParameterCount
	result.BooleanControlBranchCount = booleanMetrics.controlBranchCount
	result.BooleanControlledStatementCount = booleanMetrics.controlledStatementCount
	result.ResolvedCallees = resolvedCalleeNames(calls)
	result.ExternalDependencyCount = externalDependencyCount(calls)
	result.AmbiguousCallCount, result.UnresolvedCallCount, result.DynamicCallCount = callUncertaintyCounts(calls)
	for _, access := range procedure.Accesses {
		if access.Scope != procedureir.ScopeModule && access.Scope != procedureir.ScopeProject {
			continue
		}
		switch access.Mode {
		case procedureir.AccessRead:
			result.MutableStateReads++
		case procedureir.AccessWrite:
			result.MutableStateWrites++
			result.MutableStateMutations++
		case procedureir.AccessReadWrite:
			result.MutableStateReads++
			result.MutableStateWrites++
			result.MutableStateMutations++
		}
	}
	children := make(map[int]bool, len(statements))
	for _, statement := range statements {
		if statement.ParentID != 0 {
			children[statement.ParentID] = true
		}
	}
	for _, statement := range statements {
		texts := make([]string, 0, 4)
		// Statement.Text for a control node contains its complete nested body.
		// Only leaf text is used for executable effects; control expressions are
		// added separately so `If Range(...) Then` remains observable.
		if !children[statement.ID] {
			texts = append(texts, statement.Text)
		}
		for _, expression := range []*procedureir.Expression{statement.Target, statement.Value, statement.Condition} {
			if expression != nil {
				texts = append(texts, expression.Text)
			}
		}
		if containsExcelEffect(texts) {
			result.ExcelEffectCount++
		}
		text := strings.ToLower(statement.Text)
		if statement.Kind == procedureir.StatementOnError || statement.Kind == procedureir.StatementResume || strings.Contains(text, "err.raise") {
			result.ErrorHandlingCount++
		}
		if strings.Contains(text, "workbooks.open") || strings.HasPrefix(strings.TrimSpace(text), "open ") {
			result.ResourceOwnershipCount++
		}
	}
	return result
}

func containsExcelEffect(texts []string) bool {
	for _, text := range texts {
		lower := strings.ToLower(text)
		for _, name := range []string{"application", "range", "cells", "worksheets", "workbooks"} {
			if containsVBAReference(lower, name) {
				return true
			}
		}
	}
	return false
}

func containsVBAReference(text, name string) bool {
	for offset := 0; offset < len(text); {
		relative := strings.Index(text[offset:], name)
		if relative < 0 {
			return false
		}
		index := offset + relative
		beforeOK := index == 0 || !isVBAIdentifierByte(text[index-1])
		end := index + len(name)
		after := end
		for after < len(text) && (text[after] == ' ' || text[after] == '\t') {
			after++
		}
		afterOK := after < len(text) && (text[after] == '(' || text[after] == '.')
		if beforeOK && afterOK {
			return true
		}
		offset = index + len(name)
	}
	return false
}

func isVBAIdentifierByte(value byte) bool {
	return value == '_' || value >= 'a' && value <= 'z' || value >= 'A' && value <= 'Z' || value >= '0' && value <= '9'
}

// CollectProcedure is a convenience for callers that have only a procedure
// and an optional CFG.  It is equivalent to Collect(Input{IR: procedure,
// Graph: graph}).
func CollectProcedure(procedure procedureir.ProcedureIR, graph *cfg.Graph) ProcedureMetrics {
	return Collect(Input{IR: procedure, Graph: graph})
}

// CollectDocument computes and canonically sorts all procedures in a document.
// Graphs are matched by source declaration offset, making the result stable
// even when a caller supplies an independently sorted CFG document.
func CollectDocument(document procedureir.DocumentIR, graphs cfg.Document) []ProcedureMetrics {
	graphByStart := make(map[int]*cfg.Graph, len(graphs.Graphs))
	for i := range graphs.Graphs {
		graph := &graphs.Graphs[i]
		graphByStart[graph.Procedure.DeclarationRange.StartByte] = graph
	}
	result := make([]ProcedureMetrics, 0, len(document.Procedures))
	for _, procedure := range document.Procedures {
		var graph *cfg.Graph
		if candidate, ok := graphByStart[procedure.Symbol.DeclarationRange.StartByte]; ok {
			graph = candidate
		}
		result = append(result, Collect(Input{
			IR: procedure, Graph: graph, File: document.Path,
			Module: document.ModuleName, ModuleKind: document.ModuleKind,
		}))
	}
	Sort(result)
	return result
}

// Sort orders procedure metrics by project file, declaration position, and a
// stable identity tie-breaker.  It sorts in place and returns the same slice
// for convenient chaining.
func Sort(metrics []ProcedureMetrics) []ProcedureMetrics {
	sort.SliceStable(metrics, func(i, j int) bool {
		a, b := metrics[i], metrics[j]
		if a.File != b.File {
			return a.File < b.File
		}
		if a.DeclarationRange.StartByte != b.DeclarationRange.StartByte {
			return a.DeclarationRange.StartByte < b.DeclarationRange.StartByte
		}
		if a.DeclarationRange.StartLine != b.DeclarationRange.StartLine {
			return a.DeclarationRange.StartLine < b.DeclarationRange.StartLine
		}
		if a.DeclarationRange.EndByte != b.DeclarationRange.EndByte {
			return a.DeclarationRange.EndByte < b.DeclarationRange.EndByte
		}
		if a.DeclarationRange.EndLine != b.DeclarationRange.EndLine {
			return a.DeclarationRange.EndLine < b.DeclarationRange.EndLine
		}
		if !strings.EqualFold(a.Module, b.Module) {
			return strings.ToLower(a.Module) < strings.ToLower(b.Module)
		}
		if a.Module != b.Module {
			return a.Module < b.Module
		}
		if !strings.EqualFold(a.Name, b.Name) {
			return strings.ToLower(a.Name) < strings.ToLower(b.Name)
		}
		if a.Name != b.Name {
			return a.Name < b.Name
		}
		return a.Kind < b.Kind
	})
	return metrics
}

func orderedStatements(in []procedureir.Statement) []procedureir.Statement {
	statements := append([]procedureir.Statement(nil), in...)
	sort.SliceStable(statements, func(i, j int) bool {
		if statements[i].Range.StartByte != statements[j].Range.StartByte {
			return statements[i].Range.StartByte < statements[j].Range.StartByte
		}
		if statements[i].Range.EndByte != statements[j].Range.EndByte {
			return statements[i].Range.EndByte < statements[j].Range.EndByte
		}
		return statements[i].ID < statements[j].ID
	})
	return statements
}

func statementsFromGraph(graph cfg.Graph) []procedureir.Statement {
	statements := make([]procedureir.Statement, 0, len(graph.Blocks))
	for _, block := range graph.Blocks {
		if block.Kind == cfg.BlockStatement && block.Statement != nil {
			statements = append(statements, *block.Statement)
		}
	}
	return orderedStatements(statements)
}

func isDoCondition(statement procedureir.Statement) bool {
	return strings.EqualFold(strings.TrimSpace(statement.SyntaxKind), "do_condition")
}

func statementCount(statements []procedureir.Statement) int {
	count := 0
	for _, statement := range statements {
		if !isDoCondition(statement) {
			count++
		}
	}
	return count
}

func branchCount(statements []procedureir.Statement) int {
	count := 0
	for _, statement := range statements {
		if isDoCondition(statement) {
			continue
		}
		switch statement.Kind {
		case procedureir.StatementIf, procedureir.StatementElseIf:
			count++
		case procedureir.StatementCase:
			if !isCaseElse(statement) {
				count++
			}
		}
	}
	return count
}

type booleanControlMetricValues struct {
	booleanParameterCount         int
	optionalBooleanParameterCount int
	vagueBooleanParameterCount    int
	controlBranchCount            int
	controlledStatementCount      int
}

// booleanControlMetrics measures only source facts that are explicit in the
// procedure IR.  It deliberately does not follow aliases or interprocedural
// data flow: a Boolean parameter controls a branch only when its identifier is
// the complete condition (optionally wrapped in Not or parentheses).
func booleanControlMetrics(parameters []procedureir.Parameter, statements []procedureir.Statement) booleanControlMetricValues {
	booleanNames := make(map[string]bool)
	result := booleanControlMetricValues{}
	for _, parameter := range parameters {
		if !strings.EqualFold(strings.TrimSpace(parameter.Type), "Boolean") {
			continue
		}
		name := strings.ToLower(strings.TrimSpace(parameter.Name))
		if name == "" {
			continue
		}
		booleanNames[name] = true
		result.booleanParameterCount++
		if parameter.Optional {
			result.optionalBooleanParameterCount++
		}
		switch name {
		case "flag", "mode", "option":
			result.vagueBooleanParameterCount++
		}
	}
	if len(booleanNames) == 0 {
		return result
	}
	byID := make(map[int]procedureir.Statement, len(statements))
	for _, statement := range statements {
		if statement.ID > 0 {
			byID[statement.ID] = statement
		}
	}
	controlled := make(map[int]bool)
	for _, statement := range statements {
		if statement.Kind != procedureir.StatementIf && statement.Kind != procedureir.StatementElseIf {
			continue
		}
		name := directBooleanConditionName(statement.Condition, booleanNames)
		if name == "" {
			continue
		}
		result.controlBranchCount++
		for _, candidate := range statements {
			if candidate.ID == 0 || candidate.ID == statement.ID || isDoCondition(candidate) {
				continue
			}
			if isDescendant(candidate.ID, statement.ID, byID) {
				controlled[candidate.ID] = true
			}
		}
	}
	result.controlledStatementCount = len(controlled)
	return result
}

func directBooleanConditionName(condition *procedureir.Expression, booleanNames map[string]bool) string {
	if condition == nil {
		return ""
	}
	text := strings.TrimSpace(condition.Text)
	for {
		trimmed := strings.TrimSpace(text)
		if len(trimmed) >= 2 && strings.HasPrefix(trimmed, "(") && strings.HasSuffix(trimmed, ")") {
			text = strings.TrimSpace(trimmed[1 : len(trimmed)-1])
			continue
		}
		if len(trimmed) >= 4 && strings.EqualFold(trimmed[:4], "Not ") {
			text = strings.TrimSpace(trimmed[4:])
			continue
		}
		text = trimmed
		break
	}
	name := strings.ToLower(text)
	if booleanNames[name] {
		return name
	}
	return ""
}

func isDescendant(id, ancestor int, byID map[int]procedureir.Statement) bool {
	if id <= 0 || ancestor <= 0 || id == ancestor {
		return false
	}
	seen := map[int]bool{}
	for id > 0 && !seen[id] {
		seen[id] = true
		statement, ok := byID[id]
		if !ok {
			return false
		}
		id = statement.ParentID
		if id == ancestor {
			return true
		}
	}
	return false
}

func isCaseElse(statement procedureir.Statement) bool {
	if statement.Control != nil && statement.Control.CaseElse {
		return true
	}
	return strings.EqualFold(strings.TrimSpace(statement.Text), "case else")
}

func loopCount(statements []procedureir.Statement) int {
	count := 0
	for _, statement := range statements {
		if isDoCondition(statement) {
			continue
		}
		switch statement.Kind {
		case procedureir.StatementFor, procedureir.StatementForEach,
			procedureir.StatementWhile, procedureir.StatementDo:
			count++
		}
	}
	return count
}

func contributesNesting(kind procedureir.StatementKind) bool {
	switch kind {
	case procedureir.StatementIf, procedureir.StatementSelect,
		procedureir.StatementFor, procedureir.StatementForEach,
		procedureir.StatementWhile, procedureir.StatementDo,
		procedureir.StatementWith:
		return true
	default:
		// ElseIf, Else, and Case are alternatives inside their parent
		// construct, not additional nesting levels.
		return false
	}
}

func maxNestingDepth(statements []procedureir.Statement) int {
	byID := make(map[int]procedureir.Statement, len(statements))
	for _, statement := range statements {
		if statement.ID > 0 {
			byID[statement.ID] = statement
		}
	}
	depths := make(map[int]int, len(byID))
	visiting := make(map[int]bool, len(byID))
	var depth func(int) int
	depth = func(id int) int {
		if id <= 0 {
			return 0
		}
		if value, ok := depths[id]; ok {
			return value
		}
		if visiting[id] {
			// Malformed/recovered IR can contain a parent cycle.  Treat the
			// cycle as an unknown root rather than looping indefinitely.
			return 0
		}
		statement, ok := byID[id]
		if !ok {
			return 0
		}
		if isDoCondition(statement) {
			// The parser exposes the trailing Do condition as a synthetic
			// statement.  It contributes no level, but descendants recovered
			// under that node still belong to the enclosing loop.
			return depth(statement.ParentID)
		}
		visiting[id] = true
		value := depth(statement.ParentID)
		if contributesNesting(statement.Kind) ||
			(statement.Kind == procedureir.StatementElseIf && statement.ParentID == 0) {
			// A recovered standalone ElseIf has no enclosing If in its
			// parent chain; count it as its own conditional level.  Normal
			// ElseIf nodes are children of the outer If and remain at that
			// same level.
			value++
		}
		visiting[id] = false
		depths[id] = value
		return value
	}
	maximum := 0
	for _, statement := range statements {
		if value := depth(statement.ID); value > maximum {
			maximum = value
		}
	}
	return maximum
}

func gotoCount(statements []procedureir.Statement) int {
	count := 0
	for _, statement := range statements {
		if !isDoCondition(statement) && statement.Kind == procedureir.StatementGoTo {
			count++
		}
	}
	return count
}

func isProcedureExit(statement procedureir.Statement, kind procedureir.ProcedureKind) bool {
	if statement.Kind != procedureir.StatementExit || isDoCondition(statement) {
		return false
	}
	if statement.Control != nil {
		switch statement.Control.Transfer {
		case procedureir.TransferExitSub, procedureir.TransferExitFunction, procedureir.TransferExitProperty:
			return true
		}
	}
	// Preserve the same behavior for hand-built IR that omits control metadata.
	fields := strings.Fields(strings.ToLower(statement.Text))
	if len(fields) != 2 || fields[0] != "exit" {
		return false
	}
	switch fields[1] {
	case "sub":
		return kind == procedureir.ProcedureSub
	case "function":
		return kind == procedureir.ProcedureFunction
	case "property":
		return kind == procedureir.ProcedureProperty || kind == procedureir.ProcedurePropertyGet ||
			kind == procedureir.ProcedurePropertyLet || kind == procedureir.ProcedurePropertySet
	default:
		return false
	}
}

func exitPointCount(kind procedureir.ProcedureKind, statements []procedureir.Statement) int {
	// Reaching the matching End Sub/Function/Property is one implicit exit.
	count := 1
	for _, statement := range statements {
		if isProcedureExit(statement, kind) {
			count++
			continue
		}
		if !isDoCondition(statement) && statement.Kind == procedureir.StatementEnd {
			count++
		}
	}
	return count
}

func callFanOut(calls []procedureir.CallSite) int {
	return len(resolvedCalleeNames(calls))
}

func resolvedCalleeNames(calls []procedureir.CallSite) []string {
	seen := make(map[string]struct{}, len(calls))
	for _, call := range calls {
		if call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
			continue
		}
		candidate := call.Resolution.Candidates[0]
		key := strings.TrimSpace(candidate.QualifiedName)
		if key == "" {
			// A resolver should normally provide QualifiedName.  Keep a
			// deterministic fallback for older cached IR artifacts.
			key = fmt.Sprintf("%s:%s:%d", candidate.File, candidate.Kind, candidate.Line)
		}
		seen[strings.ToLower(key)] = struct{}{}
	}
	result := make([]string, 0, len(seen))
	for key := range seen {
		result = append(result, key)
	}
	sort.Strings(result)
	return result
}

func externalDependencyCount(calls []procedureir.CallSite) int {
	seen := map[string]bool{}
	for _, call := range calls {
		status := strings.ToLower(string(call.Resolution.Status))
		if status == string(procedureir.ResolutionExternal) || status == string(procedureir.ResolutionMemberCall) || status == string(procedureir.ResolutionBuiltinLike) {
			key := strings.ToLower(strings.TrimSpace(call.Callee.Text))
			if key != "" {
				seen[key] = true
			}
		}
	}
	return len(seen)
}

func callUncertaintyCounts(calls []procedureir.CallSite) (ambiguous, unresolved, dynamic int) {
	for _, call := range calls {
		switch call.Resolution.Status {
		case procedureir.ResolutionAmbiguous:
			ambiguous++
		case procedureir.ResolutionUnresolved, procedureir.ResolutionIncomplete:
			unresolved++
		case procedureir.ResolutionDynamic:
			dynamic++
		}
	}
	return ambiguous, unresolved, dynamic
}

func sourceLineCount(r vbast.Range) int {
	if r.StartLine <= 0 || r.EndLine < r.StartLine {
		return 0
	}
	return r.EndLine - r.StartLine + 1
}
