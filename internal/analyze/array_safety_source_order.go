package analyze

import (
	"sort"
	"strconv"
	"strings"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/constexpr"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type arraySourceOrderAllocation struct {
	line     int
	parentID int
}

type arraySourceOrderFallbackFacts struct {
	conditionalTransferLines   []int
	unconditionalTransferLines []int
	definiteExitLines          []int
	unknownFlow                bool
	parents                    map[int]procedureir.Statement
	allocations                map[string][]arraySourceOrderAllocation
	bypassTargetMin            map[int]int
	branchGroups               map[int]map[int]bool
	branchTransferBypass       map[int]map[string]int
	ambiguousTransferLines     map[int]bool
}

// buildArraySourceOrderFallbackFacts materializes the source-order facts once
// per caller. The fallback is only used for recovered CFG boundaries, but a
// caller can contain many such calls; keeping the statement and CFG scans out
// of the per-call proof avoids multiplying the recovery cost by call count.
func buildArraySourceOrderFallbackFacts(file parsedFile, proc sourceProcedure, graph *vbacfg.CFGView, variables map[string]arrayVariable, ctx analysisContext, constants map[string]int) arraySourceOrderFallbackFacts {
	facts := arraySourceOrderFallbackFacts{
		parents:                make(map[int]procedureir.Statement, proc.Statements.Len()),
		allocations:            map[string][]arraySourceOrderAllocation{},
		bypassTargetMin:        map[int]int{},
		branchGroups:           map[int]map[int]bool{},
		branchTransferBypass:   map[int]map[string]int{},
		ambiguousTransferLines: map[int]bool{},
	}
	if graph == nil {
		return facts
	}
	worklistReachable := arrayCFGWorklistReachable(graph)
	statements := make([]procedureir.Statement, 0, proc.Statements.Len())
	statementsByLine := map[int][]procedureir.Statement{}
	for statement := range proc.Statements.All() {
		statements = append(statements, statement)
		facts.parents[statement.ID] = statement
		line := statement.Range.StartLine
		if line > 0 {
			statementsByLine[line] = append(statementsByLine[line], statement)
		}
	}
	for _, statement := range statements {
		if statement.Kind != procedureir.StatementIf {
			continue
		}
		// A nested multi-line If owns an independent branch group even though
		// the IR links it to its containing If through ParentID. Register every
		// If root so allocations in all nested branches participate in the
		// source-order proof.
		facts.branchGroups[statement.ID] = map[int]bool{statement.ID: true}
	}
	hasElse := map[int]bool{}
	for _, statement := range statements {
		switch statement.Kind {
		case procedureir.StatementElseIf:
			if root, ok := arraySourceOrderIfRoot(facts.parents, statement.ID); ok {
				if branches := facts.branchGroups[root]; branches != nil {
					branches[statement.ID] = true
				}
			}
		case procedureir.StatementElse:
			if branches := facts.branchGroups[statement.ParentID]; branches != nil {
				branches[statement.ID] = true
				hasElse[statement.ParentID] = true
			}
		}
	}
	for root, branches := range facts.branchGroups {
		if !hasElse[root] {
			// An If/ElseIf chain without a final Else has an implicit path
			// that executes none of its branch bodies.
			branches[0] = true
		}
	}

	conditionalTransferLines := map[int]bool{}
	unconditionalTransferLines := map[int]bool{}
	definiteExitLines := map[int]bool{}
	for _, statement := range statements {
		line := statement.Range.StartLine
		if line <= proc.StartLine {
			continue
		}
		if statement.Kind == procedureir.StatementIf && statement.SyntaxKind == "single_line_if_statement" {
			if line >= 1 && line <= len(file.Lines) && arraySourceOrderInlineConditionalDefinitelyTerminates(normalizedCodeLine(file.Lines[line-1]), constants) {
				if block, ok := graph.BlockForStatement(statement.ID); ok && worklistReachable[block.ID] {
					definiteExitLines[line] = true
				}
			}
			for _, candidate := range statementsByLine[line] {
				if !arraySourceOrderProcedureTransfer(candidate) || !arraySourceOrderInlineConditionalTransfer(candidate, file.Lines) {
					continue
				}
				block, ok := graph.BlockForStatement(candidate.ID)
				if ok && worklistReachable[block.ID] {
					conditionalTransferLines[line] = true
					break
				}
			}
		}
		if arraySourceOrderProcedureTransfer(statement) && !arraySourceOrderInlineConditionalTransfer(statement, file.Lines) {
			unconditionalTransferLines[line] = true
		}
	}
	for line := range conditionalTransferLines {
		facts.conditionalTransferLines = append(facts.conditionalTransferLines, line)
	}
	for line := range unconditionalTransferLines {
		facts.unconditionalTransferLines = append(facts.unconditionalTransferLines, line)
	}
	for line := range definiteExitLines {
		facts.definiteExitLines = append(facts.definiteExitLines, line)
	}
	sort.Ints(facts.conditionalTransferLines)
	sort.Ints(facts.unconditionalTransferLines)
	sort.Ints(facts.definiteExitLines)

	allocationLines := map[int]bool{}
	for _, statement := range statements {
		line := statement.Range.StartLine
		if line < 1 {
			continue
		}
		block, ok := graph.BlockForStatement(statement.ID)
		if !ok || !worklistReachable[block.ID] {
			continue
		}
		for name, variable := range variables {
			if !variable.isArray || !arrayStatementAllocatesName(statement.Text, name, ctx) {
				continue
			}
			name = strings.ToLower(cleanIdentifier(name))
			facts.allocations[name] = append(facts.allocations[name], arraySourceOrderAllocation{line: line, parentID: statement.ParentID})
			allocationLines[line] = true
		}
	}
	orderedAllocationLines := make([]int, 0, len(allocationLines))
	for line := range allocationLines {
		orderedAllocationLines = append(orderedAllocationLines, line)
	}
	sort.Ints(orderedAllocationLines)
	branchAllocationLines := map[int]map[int]map[string][]int{}
	for name, allocations := range facts.allocations {
		for _, allocation := range allocations {
			root, branch, ok := facts.branchForAllocation(allocation.parentID)
			if !ok {
				continue
			}
			branches := branchAllocationLines[root]
			if branches == nil {
				branches = map[int]map[string][]int{}
			}
			names := branches[branch]
			if names == nil {
				names = map[string][]int{}
			}
			names[name] = append(names[name], allocation.line)
			branches[branch] = names
			branchAllocationLines[root] = branches
		}
	}
	for _, branches := range branchAllocationLines {
		for _, names := range branches {
			for name := range names {
				sort.Ints(names[name])
			}
		}
	}
	for line := proc.StartLine; line <= proc.EndLine && line <= len(file.Lines); line++ {
		segments := splitRangeValueSourceStatementsWithOffsets(arraySourceOrderStripComment(file.Lines[line-1]))
		hasTransfer := false
		hasAllocation := false
		for _, segment := range segments {
			if arraySourceOrderTextHasProcedureTransfer(segment.text) {
				hasTransfer = true
			}
			for name, variable := range variables {
				if variable.isArray && arrayStatementAllocatesName(segment.text, name, ctx) {
					hasAllocation = true
					break
				}
			}
			if hasTransfer && hasAllocation {
				facts.ambiguousTransferLines[line] = true
				break
			}
		}
	}

	// A source block with an edge over an allocation proves that the allocation
	// is not unconditional for calls at or after the edge target. Compute the
	// earliest such target for every candidate allocation in one graph pass;
	// individual calls then use a constant-time threshold check.
	graph.ForEachBlock(func(block vbacfg.Block) bool {
		if block.Kind != vbacfg.BlockStatement || !worklistReachable[block.ID] {
			return true
		}
		sourceLine := block.Range.StartLine
		if sourceLine <= proc.StartLine {
			return true
		}
		graph.ForEachOutgoing(block.ID, func(edge vbacfg.Edge) bool {
			target, ok := graph.BlockByID(edge.To)
			if !ok || target.Kind != vbacfg.BlockStatement {
				if block.Statement != nil && arraySourceOrderProcedureTransfer(*block.Statement) {
					facts.recordBranchTransferBypass(block, target, sourceLine, branchAllocationLines)
				}
				return true
			}
			targetLine := target.Range.StartLine
			for _, allocationLine := range orderedAllocationLines {
				if allocationLine <= sourceLine {
					continue
				}
				if allocationLine >= targetLine {
					break
				}
				if current, exists := facts.bypassTargetMin[allocationLine]; !exists || targetLine < current {
					facts.bypassTargetMin[allocationLine] = targetLine
				}
			}
			if block.Statement != nil && arraySourceOrderProcedureTransfer(*block.Statement) {
				facts.recordBranchTransferBypass(block, target, sourceLine, branchAllocationLines)
			}
			return true
		})
		return true
	})
	return facts
}

func (facts arraySourceOrderFallbackFacts) hasConditionalTransferBefore(line int) bool {
	index := sort.SearchInts(facts.conditionalTransferLines, line)
	return index > 0
}

func (facts arraySourceOrderFallbackFacts) hasDefiniteExitBefore(line int) bool {
	index := sort.SearchInts(facts.definiteExitLines, line)
	return index > 0
}

func (facts arraySourceOrderFallbackFacts) hasUnconditionalTransfer(afterLine, beforeLine int) bool {
	index := sort.Search(len(facts.unconditionalTransferLines), func(index int) bool {
		return facts.unconditionalTransferLines[index] > afterLine
	})
	return index < len(facts.unconditionalTransferLines) && facts.unconditionalTransferLines[index] < beforeLine
}

func (facts arraySourceOrderFallbackFacts) allocationDominatesCall(allocationLine, callLine int) bool {
	targetLine, bypassed := facts.bypassTargetMin[allocationLine]
	return !bypassed || targetLine > callLine
}

func (facts arraySourceOrderFallbackFacts) hasAmbiguousTransferBefore(line int) bool {
	for transferLine := range facts.ambiguousTransferLines {
		if transferLine < line {
			return true
		}
	}
	return false
}

func (facts arraySourceOrderFallbackFacts) allocationInvariant(name string, beforeLine int) bool {
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" {
		return false
	}
	if facts.hasAmbiguousTransferBefore(beforeLine) {
		return false
	}
	unconditional := false
	branchAllocations := map[int]map[int]bool{}
	for _, allocation := range facts.allocations[name] {
		if allocation.line < 1 || allocation.line >= beforeLine || facts.hasUnconditionalTransfer(allocation.line, beforeLine) {
			continue
		}
		if allocation.parentID == 0 {
			if facts.allocationDominatesCall(allocation.line, beforeLine) {
				unconditional = true
			}
			continue
		}
		root, branch, ok := facts.branchForAllocation(allocation.parentID)
		if !ok {
			continue
		}
		branches := branchAllocations[root]
		if branches == nil {
			branches = map[int]bool{}
		}
		branches[branch] = true
		branchAllocations[root] = branches
	}
	if unconditional {
		return true
	}
	for root := range facts.branchGroups {
		if facts.branchTransferBypassBefore(root, name, beforeLine) {
			return false
		}
	}
	requiredGroup := false
	for root, expected := range facts.branchGroups {
		observed := branchAllocations[root]
		if len(observed) == 0 {
			continue
		}
		requiredGroup = true
		for branch := range expected {
			if !observed[branch] {
				return false
			}
		}
	}
	return requiredGroup
}

func arraySourceOrderIfRoot(parents map[int]procedureir.Statement, statementID int) (int, bool) {
	seen := map[int]bool{}
	for statementID > 0 && !seen[statementID] {
		seen[statementID] = true
		statement, ok := parents[statementID]
		if !ok {
			return 0, false
		}
		switch statement.Kind {
		case procedureir.StatementIf:
			return statement.ID, true
		case procedureir.StatementElseIf:
			statementID = statement.ParentID
		default:
			return 0, false
		}
	}
	return 0, false
}

func (facts arraySourceOrderFallbackFacts) branchForAllocation(parentID int) (int, int, bool) {
	parent, ok := facts.parents[parentID]
	if !ok {
		return 0, 0, false
	}
	switch parent.Kind {
	case procedureir.StatementIf, procedureir.StatementElseIf:
		root, ok := arraySourceOrderIfRoot(facts.parents, parent.ID)
		return root, parent.ID, ok
	case procedureir.StatementElse:
		root, ok := arraySourceOrderIfRoot(facts.parents, parent.ParentID)
		return root, parent.ID, ok
	default:
		return 0, 0, false
	}
}

func (facts arraySourceOrderFallbackFacts) branchForStatement(statementID int) (int, int, bool) {
	seen := map[int]bool{}
	statement, ok := facts.parents[statementID]
	if !ok {
		return 0, 0, false
	}
	parentID := statement.ParentID
	for parentID > 0 && !seen[parentID] {
		seen[parentID] = true
		root, branch, branchOK := facts.branchForAllocation(parentID)
		if branchOK && facts.branchGroups[root] != nil {
			return root, branch, true
		}
		parent, parentOK := facts.parents[parentID]
		if !parentOK {
			break
		}
		parentID = parent.ParentID
	}
	return 0, 0, false
}

func (facts *arraySourceOrderFallbackFacts) recordBranchTransferBypass(block, target vbacfg.Block, sourceLine int, branchAllocationLines map[int]map[int]map[string][]int) {
	if block.Statement == nil || !arraySourceOrderProcedureTransfer(*block.Statement) {
		return
	}
	targetLine := target.Range.StartLine
	if target.Statement != nil && target.Statement.Range.StartLine > 0 {
		targetLine = target.Statement.Range.StartLine
	}
	sourceRoot, sourceBranch, sourceInBranch := facts.branchForStatement(block.Statement.ID)
	for root := range facts.branchGroups {
		rootStatement := facts.parents[root]
		if sourceLine > rootStatement.Range.EndLine {
			continue
		}
		targetAfterGroup := target.Kind != vbacfg.BlockStatement || targetLine <= 0 || targetLine > rootStatement.Range.EndLine
		if !targetAfterGroup {
			continue
		}
		if sourceLine < rootStatement.Range.StartLine || !sourceInBranch || sourceRoot != root {
			for _, names := range branchAllocationLines[root] {
				for name := range names {
					facts.saveBranchTransferBypass(root, name, sourceLine)
				}
			}
			continue
		}
		for name, lines := range branchAllocationLines[root][sourceBranch] {
			if len(lines) == 0 || lines[0] >= sourceLine {
				facts.saveBranchTransferBypass(root, name, sourceLine)
			}
		}
	}
}

func (facts *arraySourceOrderFallbackFacts) saveBranchTransferBypass(root int, name string, line int) {
	byName := facts.branchTransferBypass[root]
	if byName == nil {
		byName = map[string]int{}
		facts.branchTransferBypass[root] = byName
	}
	if previous, exists := byName[name]; !exists || line < previous {
		byName[name] = line
	}
}

func (facts arraySourceOrderFallbackFacts) branchTransferBypassBefore(root int, name string, beforeLine int) bool {
	line, ok := facts.branchTransferBypass[root][name]
	return ok && line < beforeLine
}

func arrayByRefSourceOrderFallbackApplies(file parsedFile, proc sourceProcedure, graph *vbacfg.CFGView, facts arraySourceOrderFallbackFacts, call procedureir.CallSite) bool {
	if graph == nil || facts.unknownFlow || call.Range.StartLine <= proc.StartLine || !arraySourceOrderCallLineIsSingleStatement(file.Lines, call.Range.StartLine) {
		return false
	}
	if facts.hasDefiniteExitBefore(call.Range.StartLine) {
		return false
	}
	return facts.hasConditionalTransferBefore(call.Range.StartLine)
}

func arraySourceOrderCallLineIsSingleStatement(lines []string, line int) bool {
	if line < 1 || line > len(lines) {
		return false
	}
	// The source-order fallback intentionally does not interpret statement
	// order within a colon-separated line.  Rejecting those calls keeps a
	// preceding Erase/assignment from being skipped when the recovered CFG
	// maps the whole line to one statement.
	code := strings.TrimSpace(normalizedCodeLine(arraySourceOrderStripComment(lines[line-1])))
	// VBA permits a Rem comment after a statement separator. The separator is
	// part of the comment boundary, not a second executable statement, so do
	// not let colons inside that comment disable the fallback.
	if strings.HasSuffix(code, ":") {
		code = strings.TrimSpace(strings.TrimSuffix(code, ":"))
	}
	return !arraySourceOrderHasStatementSeparator(code)
}

func arraySourceOrderHasStatementSeparator(code string) bool {
	inString := false
	for index := 0; index < len(code); index++ {
		switch code[index] {
		case '"':
			if inString && index+1 < len(code) && code[index+1] == '"' {
				index++
				continue
			}
			inString = !inString
		case ':':
			if inString {
				continue
			}
			next := index + 1
			for next < len(code) && (code[next] == ' ' || code[next] == '\t') {
				next++
			}
			if next < len(code) && code[next] == '=' {
				continue
			}
			return true
		}
	}
	return false
}

func arraySourceOrderTextHasProcedureTransfer(text string) bool {
	text = strings.TrimSpace(text)
	if _, body, ok := arrayIfThenParts(text); ok {
		for _, part := range splitRangeValueSourceStatements(body) {
			if arraySourceOrderTextHasProcedureTransfer(part) {
				return true
			}
		}
		return false
	}
	lower := strings.ToLower(text)
	if strings.HasPrefix(lower, "goto ") {
		return true
	}
	if lower == "end" || lower == "end sub" || lower == "end function" || lower == "end property" {
		return true
	}
	for _, prefix := range []string{"exit sub", "exit function", "exit property"} {
		if lower == prefix || strings.HasPrefix(lower, prefix+" ") {
			return true
		}
	}
	return false
}

func arraySourceOrderProcedureTransfer(statement procedureir.Statement) bool {
	switch statement.Kind {
	case procedureir.StatementGoTo:
		_, isGoSub := arrayLocalGoSubTargetFromStatementText(statement.Text)
		return !isGoSub
	case procedureir.StatementEnd:
		return true
	case procedureir.StatementExit:
		if statement.Control == nil {
			lower := strings.ToLower(strings.TrimSpace(statement.Text))
			return strings.HasPrefix(lower, "exit sub") || strings.HasPrefix(lower, "exit function") || strings.HasPrefix(lower, "exit property")
		}
		switch statement.Control.Transfer {
		case procedureir.TransferExitSub, procedureir.TransferExitFunction, procedureir.TransferExitProperty, procedureir.TransferTerminate:
			return true
		default:
			return false
		}
	default:
		return false
	}
}

func arraySourceOrderInlineConditionalTransfer(statement procedureir.Statement, lines []string) bool {
	line := statement.Range.StartLine
	if line < 1 || line > len(lines) {
		return false
	}
	text := strings.ToLower(strings.TrimSpace(normalizedCodeLine(lines[line-1])))
	if !strings.HasPrefix(text, "if ") {
		return false
	}
	then := strings.Index(text, " then")
	if then < 0 {
		return false
	}
	suffix := strings.TrimSpace(text[then+len(" then"):])
	if suffix == "" {
		return false
	}
	switch statement.Kind {
	case procedureir.StatementGoTo:
		return strings.Contains(suffix, "goto ")
	case procedureir.StatementEnd:
		return strings.HasPrefix(suffix, "end")
	case procedureir.StatementExit:
		return strings.Contains(suffix, "exit sub") || strings.Contains(suffix, "exit function") || strings.Contains(suffix, "exit property")
	default:
		return false
	}
}

func arraySourceOrderInlineConditionalDefinitelyTerminates(text string, constants map[string]int) bool {
	condition, body, ok := arrayIfThenParts(text)
	if !ok || strings.TrimSpace(body) == "" {
		return false
	}
	condition = strings.TrimSpace(condition)
	lowerCondition := strings.ToLower(condition)
	switch {
	case strings.HasPrefix(lowerCondition, "if "):
		condition = strings.TrimSpace(condition[len("if "):])
	case strings.HasPrefix(lowerCondition, "elseif "):
		condition = strings.TrimSpace(condition[len("elseif "):])
	default:
		return false
	}
	value, known := arraySourceOrderConstantBoolean(condition, constants)
	if !known {
		return false
	}
	thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
	selected := thenBody
	if !value {
		if !hasElse {
			return false
		}
		selected = elseBody
	}
	return arraySourceOrderInlineBodyDefinitelyTerminates(selected, constants)
}

func arraySourceOrderInlineBodyDefinitelyTerminates(text string, constants map[string]int) bool {
	for _, statement := range splitRangeValueSourceStatements(text) {
		statement = strings.TrimSpace(statement)
		if statement == "" {
			continue
		}
		if arraySourceOrderProcedureExitStatementText(statement) || arraySourceOrderInlineConditionalDefinitelyTerminates(statement, constants) {
			return true
		}
	}
	return false
}

func arraySourceOrderProcedureExitStatementText(text string) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	for _, prefix := range []string{"exit sub", "exit function", "exit property"} {
		if lower == prefix || strings.HasPrefix(lower, prefix+" ") {
			return true
		}
	}
	switch lower {
	case "end", "end sub", "end function", "end property":
		return true
	default:
		return false
	}
}

func arraySourceOrderConstantBoolean(expression string, constants map[string]int) (bool, bool) {
	values := make(map[string]constexpr.Value, len(constants))
	for name, value := range constants {
		values[name] = constexpr.Value{Kind: constexpr.ValueLongLong, Integer: int64(value)}
	}
	result := constexpr.Evaluate(expression, constexpr.NewValues(values))
	if result.Kind != constexpr.Known || result.Typed.Kind != constexpr.ValueBoolean {
		return false, false
	}
	return result.Typed.Boolean, true
}

func arraySourceOrderInlineArrayMutation(text string, names map[string]bool, ctx analysisContext) bool {
	condition, body, ok := arrayIfThenParts(text)
	if !ok || strings.TrimSpace(body) == "" {
		return false
	}
	condition = strings.TrimSpace(condition)
	lowerCondition := strings.ToLower(condition)
	switch {
	case strings.HasPrefix(lowerCondition, "if "):
		condition = strings.TrimSpace(condition[len("if "):])
	case strings.HasPrefix(lowerCondition, "elseif "):
		condition = strings.TrimSpace(condition[len("elseif "):])
	}
	thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
	branches := []string{thenBody}
	if hasElse {
		branches = append(branches, elseBody)
	}
	if value, known := arraySourceOrderConstantBoolean(condition, nil); known {
		if value {
			branches = []string{thenBody}
		} else if hasElse {
			branches = []string{elseBody}
		} else {
			branches = nil
		}
	}
	for _, branch := range branches {
		for _, statement := range splitRangeValueSourceStatements(branch) {
			statement = strings.TrimSpace(statement)
			if statement == "" {
				continue
			}
			if arraySourceOrderInlineArrayMutation(statement, names, ctx) {
				return true
			}
			if arraySourceOrderMutatesArrayStatement(statement, names, ctx) {
				return true
			}
		}
	}
	return false
}

func arraySourceOrderMutatesArrayStatement(text string, names map[string]bool, ctx analysisContext) bool {
	text = strings.TrimSpace(text)
	if match := arrayEraseRe.FindStringSubmatch(text); len(match) == 2 {
		for _, target := range splitArgs(match[1]) {
			if names[strings.ToLower(cleanIdentifier(strings.TrimSpace(target)))] {
				return true
			}
		}
	}
	if match := arrayRedimRe.FindStringSubmatch(text); len(match) > 0 {
		if strings.TrimSpace(match[1]) != "" {
			// ReDim Preserve keeps an allocated input array allocated. The
			// source-order state solver handles the separate requirement that
			// the input must already be allocated.
			return false
		}
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			name := strings.ToLower(cleanIdentifier(redim.name))
			if direct && names[name] && !arrayStatementAllocatesName(text, name, ctx) {
				return true
			}
		}
	}
	if lhs, _, indexed, ok := arrayAssignment(text); ok && !indexed {
		name := strings.ToLower(cleanIdentifier(lhs))
		return names[name] && !arrayStatementAllocatesName(text, name, ctx)
	}
	return false
}

// arraySourceOrderStripComment removes VBA apostrophe and Rem comments while
// preserving the byte offsets before the comment. The latter matters because
// recovered CallSite ranges are mapped back to colon-separated source
// segments below.
func arraySourceOrderStripComment(line string) string {
	inString := false
	for index := 0; index < len(line); index++ {
		switch line[index] {
		case '"':
			if inString && index+1 < len(line) && line[index+1] == '"' {
				index++
				continue
			}
			inString = !inString
		case '\'':
			if !inString {
				return line[:index]
			}
		default:
			if !inString && arraySourceOrderRemCommentAt(line, index) {
				return line[:index]
			}
		}
	}
	return line
}

func arraySourceOrderRemCommentAt(line string, index int) bool {
	if index+3 > len(line) || !strings.EqualFold(line[index:index+3], "Rem") {
		return false
	}
	if index > 0 {
		previous := index - 1
		for previous >= 0 {
			switch line[previous] {
			case ' ', '\t', '\r', '\n':
				previous--
				continue
			}
			break
		}
		if previous >= 0 && line[previous] != ':' {
			return false
		}
	}
	if index+3 < len(line) {
		next := line[index+3]
		if next != ':' && next != ' ' && next != '\t' && next != '\r' && next != '\n' {
			return false
		}
	}
	return true
}

func arraySourceOrderLineStartByte(source []byte, line int) int {
	if line <= 1 {
		return 0
	}
	current := 1
	for index, value := range source {
		if value != '\n' {
			continue
		}
		current++
		if current == line {
			return index + 1
		}
	}
	return -1
}

func arraySourceOrderCallsBySegment(file parsedFile, line int, calls []procedureir.CallSite) ([][]procedureir.CallSite, []procedureir.CallSite) {
	if line < 1 || line > len(file.Lines) {
		return nil, calls
	}
	segments := splitRangeValueSourceStatementsWithOffsets(arraySourceOrderStripComment(file.Lines[line-1]))
	if len(segments) == 0 {
		return nil, calls
	}
	bySegment := make([][]procedureir.CallSite, len(segments))
	unassigned := make([]procedureir.CallSite, 0)
	lineStart := arraySourceOrderLineStartByte(file.Source, line)
	for _, call := range calls {
		segmentIndex := -1
		if lineStart >= 0 {
			relative := call.Range.StartByte - lineStart
			for index, segment := range segments {
				if relative >= segment.start && (relative < segment.end || index == len(segments)-1 && relative <= segment.end) {
					segmentIndex = index
					break
				}
			}
		}
		if segmentIndex < 0 {
			unassigned = append(unassigned, call)
			continue
		}
		bySegment[segmentIndex] = append(bySegment[segmentIndex], call)
	}
	return bySegment, unassigned
}

func (a Analyzer) arrayByRefCallSourceOrderProof(file parsedFile, facts arraySourceOrderFallbackFacts, localGoSubAllocations arrayLocalGoSubAllocations, caller, target sourceProcedure, call procedureir.CallSite, initial arrayFlowState, ctx analysisContext, variables map[string]arrayVariable, constants map[string]int) (arrayFlowState, bool) {
	bindings, ok := arrayCallArgumentBindings(caller, target, call)
	if !ok {
		return nil, false
	}
	arguments := make([]string, target.Params.Len())
	bound := make(map[int]bool, len(bindings))
	for _, binding := range bindings {
		arguments[binding.parameterIndex] = binding.text
		bound[binding.parameterIndex] = true
	}
	for index, parameter := range target.Params.AllIndexed() {
		if !parameterIsByRefArray(parameter) || !bound[index] {
			continue
		}
		if directArrayArgumentName(arguments[index]) == "" && !arrayQualifiedArgumentProvenAllocated(file, caller, call, arguments[index], ctx) {
			return nil, false
		}
	}
	arrayNames := map[string]bool{}
	initiallyAllocated := map[string]bool{}
	for index, parameter := range target.Params.AllIndexed() {
		if parameterIsByRefArray(parameter) && index < len(arguments) {
			name := strings.ToLower(cleanIdentifier(directArrayArgumentName(arguments[index])))
			arrayNames[name] = true
			value, known := initial[name]
			initiallyAllocated[name] = known && value.kind == arrayAllocated && value.knownArray
		}
	}
	if call.Range.StartLine <= caller.StartLine || !arraySourceOrderCallLineIsSingleStatement(file.Lines, call.Range.StartLine) {
		return nil, false
	}
	state := cloneArrayState(initial)
	// Recovered CFG blocks can coalesce a colon-separated physical line. Apply
	// each logical segment in source order so Erase or another mutation before a
	// later call cannot be hidden by passing the whole physical line to transfer.
	for line := caller.StartLine; line < call.Range.StartLine && line <= len(file.Lines); line++ {
		segments := splitRangeValueSourceStatementsWithOffsets(arraySourceOrderStripComment(file.Lines[line-1]))
		callsBySegment, unassignedCalls := arraySourceOrderCallsBySegment(file, line, arrayCallsAtLine(caller.Calls, line))
		if len(unassignedCalls) > 0 && len(segments) > 1 {
			// A call without a trustworthy source offset cannot be placed among
			// colon-separated statements. Continuing would apply its side effect
			// after the whole line and could silently reverse an allocation and a
			// later Erase. The fallback is intentionally fail-closed here.
			return nil, false
		}
		for segmentIndex, segment := range segments {
			statement := normalizedCodeLine(segment.text)
			if strings.TrimSpace(statement) == "" {
				continue
			}
			if arraySourceOrderPreserveNeedsAllocatedInput(statement, arrayNames, state) {
				// A successful ReDim Preserve requires an already allocated
				// dynamic array. The source-order fallback has no exceptional
				// edge on which an On Error Resume Next failure could continue,
				// so an unallocated input cannot be used as an allocation proof.
				return nil, false
			}
			// Ordinary ReDim and array-factory assignments intentionally continue
			// through arrayTransfer below. Only an inline conditional mutation is
			// rejected because its branch state is not modeled by this fallback.
			if arraySourceOrderInlineArrayMutation(statement, arrayNames, ctx) {
				return nil, false
			}
			state, _ = a.arrayTransfer(file, caller, ctx, variables, state, statement, line, constants, nil)
			state = applyArrayLocalGoSubStatementEffects(state, statement, localGoSubAllocations)
			for _, previous := range callsBySegment[segmentIndex] {
				state = applyArrayModuleCallEffects(state, file, caller, previous, ctx, variables, file.moduleDecls())
				state = applyArrayUnknownModuleCallEffects(state, file, caller, previous, ctx, variables, file.moduleDecls())
				if arrayProcedureLineHasInlineConditional(file, previous.Range.StartLine) {
					state = applyArrayConditionalByRefCallEffects(state, caller, previous, ctx)
				} else {
					state = applyArrayByRefCallEffects(state, caller, previous, ctx)
				}
			}
		}
		for _, previous := range unassignedCalls {
			state = applyArrayModuleCallEffects(state, file, caller, previous, ctx, variables, file.moduleDecls())
			state = applyArrayUnknownModuleCallEffects(state, file, caller, previous, ctx, variables, file.moduleDecls())
			if arrayProcedureLineHasInlineConditional(file, previous.Range.StartLine) {
				state = applyArrayConditionalByRefCallEffects(state, caller, previous, ctx)
			} else {
				state = applyArrayByRefCallEffects(state, caller, previous, ctx)
			}
		}
	}
	for index, parameter := range target.Params.AllIndexed() {
		if !parameterIsByRefArray(parameter) || !bound[index] {
			continue
		}
		name := directArrayArgumentName(arguments[index])
		if name == "" {
			if !arrayQualifiedArgumentProvenAllocated(file, caller, call, arguments[index], ctx) {
				return nil, false
			}
			continue
		}
		value, known := state[name]
		if !known || value.kind != arrayAllocated || !value.knownArray || (!facts.allocationInvariant(name, call.Range.StartLine) && !initiallyAllocated[name]) {
			return nil, false
		}
	}
	return state, true
}

func arraySourceOrderPreserveNeedsAllocatedInput(text string, names map[string]bool, state arrayFlowState) bool {
	match := arrayRedimRe.FindStringSubmatch(strings.TrimSpace(text))
	if len(match) == 0 || strings.TrimSpace(match[1]) == "" {
		return false
	}
	for _, clause := range splitArgs(match[2]) {
		redim, direct := parseDirectArrayRedimClause(clause)
		name := strings.ToLower(cleanIdentifier(redim.name))
		if !direct || name == "" || !names[name] {
			continue
		}
		value, known := state[name]
		if !known || value.kind != arrayAllocated || !value.knownArray {
			return true
		}
	}
	return false
}

func arrayStatementAllocatesName(text, name string, ctx analysisContext) bool {
	if match := arrayRedimRe.FindStringSubmatch(strings.TrimSpace(text)); len(match) > 0 && strings.TrimSpace(match[1]) == "" {
		for _, clause := range splitArgs(match[2]) {
			redim, direct := parseDirectArrayRedimClause(clause)
			if direct && strings.EqualFold(cleanIdentifier(redim.name), name) {
				return true
			}
		}
	}
	if lhs, rhs, indexed, ok := arrayAssignment(text); ok && !indexed && strings.EqualFold(cleanIdentifier(lhs), name) {
		value, known := arrayExpressionState(rhs, arrayFlowState{}, ctx)
		return known && value.kind == arrayAllocated && value.knownArray
	}
	return false
}

func arrayByRefCallHasProvenArrayArguments(file parsedFile, target, caller sourceProcedure, call procedureir.CallSite, state arrayFlowState, ctx analysisContext) (bool, bool) {
	bindings, ok := arrayCallArgumentBindings(caller, target, call)
	if !ok {
		return false, false
	}
	foundExpression := false
	for _, binding := range bindings {
		if binding.parameterIndex < 0 || binding.parameterIndex >= target.Params.Len() {
			return false, false
		}
		parameter := target.Params.valueAt(binding.parameterIndex)
		if !parameterIsByRefArray(parameter) {
			continue
		}
		argument := binding.text
		if name := directArrayArgumentName(argument); name != "" {
			value, known := state[name]
			if !known || value.kind != arrayAllocated || !value.knownArray {
				return false, false
			}
			continue
		}
		if arrayQualifiedArgumentProvenAllocated(file, caller, call, argument, ctx) {
			foundExpression = true
			continue
		}
		if _, _, qualifiedMember := arrayQualifiedMemberParts(argument); qualifiedMember {
			return false, false
		}
		value, known := arrayExpressionState(argument, state, ctx)
		if !known || value.kind != arrayAllocated || !value.knownArray {
			return false, false
		}
		foundExpression = true
	}
	return true, foundExpression
}

func arrayByRefCallIsInnermostNested(call procedureir.CallSite, calls []arrayByRefCallCandidate) bool {
	nested := false
	for _, other := range calls {
		if arrayCallRangeContains(other.call, call) {
			nested = true
		}
		if arrayCallRangeContains(call, other.call) {
			return false
		}
	}
	return nested
}

func arrayCallRangeContains(outer, inner procedureir.CallSite) bool {
	if outer.ID == inner.ID {
		return false
	}
	if outer.Range.StartByte != 0 || outer.Range.EndByte != 0 || inner.Range.StartByte != 0 || inner.Range.EndByte != 0 {
		return outer.Range.StartByte <= inner.Range.StartByte && inner.Range.EndByte <= outer.Range.EndByte && (outer.Range.StartByte < inner.Range.StartByte || inner.Range.EndByte < outer.Range.EndByte)
	}
	return outer.Range.StartLine == inner.Range.StartLine && outer.Range.StartColumn <= inner.Range.StartColumn && inner.Range.EndColumn <= outer.Range.EndColumn && (outer.Range.StartColumn < inner.Range.StartColumn || inner.Range.EndColumn < outer.Range.EndColumn)
}

func arrayByRefCallArrayVacuouslyUnused(target sourceProcedure, arrayIndex int, arguments []string) bool {
	if arrayIndex < 0 || arrayIndex >= target.Params.Len() {
		return false
	}
	arrayName := strings.ToLower(cleanIdentifier(target.Params.valueAt(arrayIndex).Name))
	if arrayName == "" {
		return false
	}
	statements := make(map[int]procedureir.Statement, target.Statements.Len())
	for statement := range target.Statements.All() {
		statements[statement.ID] = statement
	}
	variables := map[string]arrayVariable{
		arrayName: {name: arrayName, isArray: true},
	}
	found := false
	for statement := range target.Statements.All() {
		if !arrayStatementReferencesArray(statement.Text, arrayName, variables) {
			continue
		}
		found = true
		if !arrayStatementHasFalseCallGuard(statement, statements, target, arguments) {
			return false
		}
	}
	return found
}

func arrayStatementReferencesArray(text, arrayName string, variables map[string]arrayVariable) bool {
	for _, use := range arrayIndexedUses(text, variables) {
		if strings.EqualFold(cleanIdentifier(use.name), arrayName) && len(use.args) > 0 {
			return true
		}
	}
	for _, bound := range arrayBoundCallRe.FindAllStringSubmatch(text, -1) {
		if strings.EqualFold(cleanIdentifier(bound[2]), arrayName) {
			return true
		}
	}
	return false
}

func arrayStatementHasFalseCallGuard(statement procedureir.Statement, statements map[int]procedureir.Statement, target sourceProcedure, arguments []string) bool {
	seen := map[int]bool{}
	for current := statement; current.ParentID != 0 && !seen[current.ParentID]; {
		seen[current.ParentID] = true
		parent, ok := statements[current.ParentID]
		if !ok {
			return false
		}
		if parent.Kind == procedureir.StatementIf && parent.Condition != nil && arrayConditionHasFalseCallArgument(parent.Condition.Text, target, arguments) {
			return true
		}
		current = parent
	}
	return false
}

func arrayConditionHasFalseCallArgument(condition string, target sourceProcedure, arguments []string) bool {
	for _, comparison := range arrayConditionAndRe.Split(condition, -1) {
		comparison = strings.TrimSpace(comparison)
		if index, ok := arrayProcedureParameterIndex(target, comparison); ok && index < len(arguments) && strings.EqualFold(strings.TrimSpace(arguments[index]), "false") {
			return true
		}
		lhs, operator, literal, ok := arrayCountComparison(comparison)
		if !ok {
			continue
		}
		index, ok := arrayProcedureParameterIndex(target, lhs)
		if !ok || index >= len(arguments) {
			continue
		}
		value, valueOK := integerLiteral(arguments[index])
		bound, boundOK := integerLiteral(literal)
		if valueOK && boundOK && arrayIntegerComparisonFalse(value, operator, bound) {
			return true
		}
	}
	return false
}

func arrayProcedureParameterIndex(proc sourceProcedure, name string) (int, bool) {
	name = strings.ToLower(cleanIdentifier(name))
	for index, parameter := range proc.Params.AllIndexed() {
		if strings.ToLower(cleanIdentifier(parameter.Name)) == name {
			return index, true
		}
	}
	return 0, false
}

func arrayIntegerComparisonFalse(value int, operator string, bound int) bool {
	switch operator {
	case "=":
		return value != bound
	case "<>":
		return value == bound
	case ">":
		return value <= bound
	case ">=":
		return value < bound
	case "<":
		return value >= bound
	case "<=":
		return value > bound
	default:
		return false
	}
}

func arrayConditionalEntrySource(target sourceProcedure, arguments []string, arrayParameterIndex int, source string) string {
	if source == "" {
		return ""
	}
	for index, argument := range arguments {
		if index == arrayParameterIndex || !arrayCountExpressionMatches(argument, source) {
			continue
		}
		if index < 0 || index >= target.Params.Len() || parameterIsByRefArray(target.Params.valueAt(index)) {
			continue
		}
		return strings.ToLower(target.Params.valueAt(index).Name)
	}
	return ""
}

func arrayCallArgumentTexts(proc sourceProcedure, call procedureir.CallSite) []string {
	if len(call.Arguments.ExpressionIDs) == 0 {
		return nil
	}
	facts := proc.analysisFacts()
	texts := make([]string, 0, len(call.Arguments.ExpressionIDs))
	for _, id := range call.Arguments.ExpressionIDs {
		expression, ok := facts.Expression(id)
		if !ok {
			return nil
		}
		texts = append(texts, strings.TrimSpace(expression.Text))
	}
	return texts
}

type arrayCallArgumentBinding struct {
	parameterIndex int
	text           string
}

// arrayCallArgumentBindings maps actual arguments to formal parameters while
// retaining VBA's source order for positional arguments. Named arguments are
// identified by their expression IDs, so a positional argument before a named
// argument is not mistaken for the named formal slot.
func arrayCallArgumentBindings(caller sourceProcedure, target sourceProcedure, call procedureir.CallSite) ([]arrayCallArgumentBinding, bool) {
	argumentCount := call.Arguments.Count
	// Some unparenthesized calls with an implicit member argument currently
	// expose the expression ID but leave Count at zero.  Prefer the concrete
	// expression projection when Count is absent; otherwise the ByRef transfer
	// would silently discard a real actual argument.
	if argumentCount == 0 {
		argumentCount = max(len(call.Arguments.ExpressionIDs), len(call.Arguments.Named))
	}
	if argumentCount == 0 {
		if argument, ok := arrayImplicitMemberArgument(call); ok {
			if target.Params.Len() != 1 {
				return nil, false
			}
			return []arrayCallArgumentBinding{{parameterIndex: 0, text: argument}}, true
		}
		return nil, true
	}
	if len(call.Arguments.ExpressionIDs) == 0 {
		if len(call.Arguments.Named) != argumentCount {
			return nil, false
		}
		bindings := make([]arrayCallArgumentBinding, 0, len(call.Arguments.Named))
		used := map[int]bool{}
		for _, named := range call.Arguments.Named {
			index, ok := arrayFormalParameterIndex(target, named.Name)
			if !ok || used[index] {
				return nil, false
			}
			used[index] = true
			bindings = append(bindings, arrayCallArgumentBinding{parameterIndex: index, text: strings.TrimSpace(named.ValueText)})
		}
		return bindings, true
	}
	texts := arrayCallArgumentTexts(caller, call)
	if len(texts) != argumentCount || len(call.Arguments.ExpressionIDs) != len(texts) {
		return nil, false
	}
	namedByExpressionID := make(map[int]int, len(call.Arguments.Named))
	for _, named := range call.Arguments.Named {
		if named.ExpressionID == 0 {
			return nil, false
		}
		index, ok := arrayFormalParameterIndex(target, named.Name)
		if !ok {
			return nil, false
		}
		if _, exists := namedByExpressionID[named.ExpressionID]; exists {
			return nil, false
		}
		namedByExpressionID[named.ExpressionID] = index
	}
	bindings := make([]arrayCallArgumentBinding, 0, len(texts))
	used := map[int]bool{}
	nextPositional := 0
	for actualIndex, text := range texts {
		formalIndex, named := namedByExpressionID[call.Arguments.ExpressionIDs[actualIndex]]
		if !named {
			for nextPositional < target.Params.Len() && used[nextPositional] {
				nextPositional++
			}
			if nextPositional < target.Params.Len() {
				formalIndex = nextPositional
				nextPositional++
			} else {
				last := target.Params.Len() - 1
				if last < 0 || !target.Params.valueAt(last).ParamArray {
					return nil, false
				}
				// Every extra positional argument belongs to the trailing
				// ParamArray formal. Keep one binding per actual so a direct
				// array argument among the extras still participates in the
				// ByRef effect analysis.
				formalIndex = last
			}
		}
		if used[formalIndex] && (formalIndex < 0 || formalIndex >= target.Params.Len() || !target.Params.valueAt(formalIndex).ParamArray) {
			return nil, false
		}
		used[formalIndex] = true
		bindings = append(bindings, arrayCallArgumentBinding{parameterIndex: formalIndex, text: text})
	}
	return bindings, true
}

func arrayImplicitMemberArgument(call procedureir.CallSite) (string, bool) {
	if call.Arguments.Count != 0 || len(call.Arguments.ExpressionIDs) != 0 || len(call.Arguments.Named) != 0 || call.Callee.Receiver == nil {
		return "", false
	}
	procedureName := arrayImplicitMemberArgumentCallName(call.Callee.Text)
	if procedureName == "" || !strings.EqualFold(procedureName, *call.Callee.Receiver) {
		return "", false
	}
	member := cleanIdentifier(call.Callee.Member)
	if member == "" {
		return "", false
	}
	return "." + member, true
}

func arrayCallFormalArguments(caller sourceProcedure, target sourceProcedure, call procedureir.CallSite) ([]string, bool) {
	bindings, ok := arrayCallArgumentBindings(caller, target, call)
	if !ok {
		return nil, false
	}
	arguments := make([]string, target.Params.Len())
	for _, binding := range bindings {
		if binding.parameterIndex < 0 || binding.parameterIndex >= len(arguments) {
			return nil, false
		}
		arguments[binding.parameterIndex] = binding.text
	}
	return arguments, true
}

func arrayFormalParameterIndex(target sourceProcedure, name string) (int, bool) {
	name = strings.TrimSpace(name)
	for index, parameter := range target.Params.AllIndexed() {
		if strings.EqualFold(strings.TrimSpace(parameter.Name), name) {
			return index, true
		}
	}
	return 0, false
}

func arrayCallPassesDirectArrayArgument(proc sourceProcedure, call procedureir.CallSite, name string) bool {
	name = strings.ToLower(cleanIdentifier(name))
	for _, argument := range call.Arguments.Named {
		if directArrayArgumentName(argument.ValueText) == name {
			return true
		}
	}
	for _, argument := range arrayCallArgumentTexts(proc, call) {
		if directArrayArgumentName(argument) == name {
			return true
		}
	}
	return false
}

func directArrayArgumentName(text string) string {
	text = strings.TrimSpace(text)
	if text == "" || !isIdentifierStart(text[0]) {
		return ""
	}
	for i := 1; i < len(text); i++ {
		if !isIdentifierPart(text[i]) {
			return ""
		}
	}
	return strings.ToLower(cleanIdentifier(text))
}

// arrayQualifiedArgumentProvenAllocated carries narrow allocation proofs for a
// qualified array member passed to a private ByRef array parameter. The
// ordinary array state is intentionally keyed by local identifiers, so a call
// such as `Consume holder.values` cannot otherwise reuse a dominating ReDim of
// the member. The additional proofs are limited to normal-path ReDim,
// descriptor-backed array setup, and dictionary snapshots whose non-empty
// range is established by the same caller.
func arrayQualifiedArgumentProvenAllocated(file parsedFile, caller sourceProcedure, call procedureir.CallSite, argument string, ctx analysisContext) bool {
	want := arrayQualifiedArgumentTarget(file, caller, call.Range.StartLine, argument)
	if want == "" || caller.Graph == nil {
		return false
	}
	if arrayQualifiedDescriptorArgumentProvenAllocated(file, caller, call, argument, ctx) ||
		arrayQualifiedDictionarySnapshotProvenAllocated(file, caller, call, argument) {
		return true
	}
	for statement := range caller.Statements.All() {
		line := statement.Range.StartLine
		dominates := arrayStatementDominatesCall(caller, statement.ID, statement.Range.StartLine, call)
		redim := arrayQualifiedRedimAllocatesTarget(file, caller, line, statement.Text, want)
		if line <= caller.StartLine || line >= call.Range.StartLine || !dominates {
			continue
		}
		if redim {
			return true
		}
	}
	return false
}

// arrayQualifiedDescriptorArgumentProvenAllocated recognizes the accessor
// pattern used when a VBA SAFEARRAY descriptor is projected onto a typed array
// member. The array itself is not visible to the source-level analyzer, but a
// dominating data pointer and element-count write, followed by a positive
// count guard, establish the same normal-path contract as ReDim. A projected
// `rgsabound()` array has one additional form: its caller can establish a
// successful `ReDim ... (0 To ub)` from the descriptor dimension count before
// passing the accessor through another private helper. That normal-path fact
// proves the descriptor count is positive without requiring a duplicated scalar
// count argument at every helper boundary.
func arrayQualifiedDescriptorArgumentProvenAllocated(file parsedFile, caller sourceProcedure, call procedureir.CallSite, argument string, ctx analysisContext) bool {
	receiver, member, ok := arrayQualifiedMemberParts(argument)
	if !ok || !strings.EqualFold(member, "arr") && !strings.EqualFold(member, "rgsabound") {
		return false
	}
	if receiver == "" {
		receiver = arrayWithReceiverAtLine(file, caller, call.Range.StartLine)
	}
	if receiver == "" {
		return false
	}
	arguments := arrayCallArgumentTexts(caller, call)
	if len(arguments) < 2 {
		if !strings.EqualFold(member, "rgsabound") {
			return false
		}
		normalPathNonEmpty := arrayQualifiedDescriptorHasSuccessfulShapeUse(file, caller, call, receiver)
		return arrayQualifiedDescriptorReceiverInitialized(file, caller, call, receiver, ctx, !normalPathNonEmpty, map[string]bool{})
	}
	count := canonicalArrayBoundExpression(arguments[1])
	if count != "" && arrayQualifiedDescriptorCountPositive(file, caller, call, arguments[1]) {
		wantReceiver := canonicalArrayBoundExpression(receiver)
		wantData := wantReceiver + ".sa.pvdata"
		wantBounds := wantReceiver + ".sa.rgsabound0.celements"
		hasData := false
		hasBounds := false
		for statement := range caller.Statements.All() {
			line := statement.Range.StartLine
			if line <= caller.StartLine || line >= call.Range.StartLine || !arrayStatementDominatesCall(caller, statement.ID, line, call) {
				continue
			}
			text := strings.TrimSpace(statement.Text)
			if text == "" {
				text = arrayLogicalSourceLine(file.Lines, line)
			}
			lhs, rhs, indexed, assigned := arrayAssignment(text)
			if !assigned || indexed {
				continue
			}
			target := canonicalArrayBoundExpression(lhs)
			switch target {
			case wantData:
				hasData = strings.TrimSpace(rhs) != ""
			case wantBounds:
				hasBounds = canonicalArrayBoundExpression(rhs) == count
			}
		}
		if hasData && hasBounds {
			return true
		}
	}
	if !strings.EqualFold(member, "rgsabound") || !arrayQualifiedDescriptorHasSuccessfulShapeUse(file, caller, call, receiver) {
		return false
	}
	return arrayQualifiedDescriptorReceiverInitialized(file, caller, call, receiver, ctx, false, map[string]bool{})
}

func arrayQualifiedDescriptorReceiverInitialized(file parsedFile, proc sourceProcedure, call procedureir.CallSite, receiver string, ctx analysisContext, requirePositive bool, visiting map[string]bool) bool {
	receiver = strings.ToLower(cleanIdentifier(strings.TrimSpace(receiver)))
	if receiver == "" {
		return false
	}
	parameterIndex, isParameter := arrayProcedureParameterIndex(proc, receiver)
	if !isParameter {
		return arrayQualifiedDescriptorLocalInitialized(file, proc, call, receiver, requirePositive)
	}
	visitKey := arrayProcedureKey(proc) + ":" + receiver
	if visiting[visitKey] {
		return false
	}
	visiting[visitKey] = true
	defer delete(visiting, visitKey)

	foundCaller := false
	for _, candidate := range file.Procedures {
		for nested := range candidate.Calls.All() {
			targetKey, target, ok := arrayPrivateTargetForCall(ctx, ctx.arrayPrivateTargets, nested)
			if !ok || targetKey != arrayProcedureKey(proc) {
				continue
			}
			bindings, bound := arrayCallArgumentBindings(candidate, target, nested)
			if !bound {
				return false
			}
			actual := ""
			for _, binding := range bindings {
				if binding.parameterIndex == parameterIndex {
					actual = strings.TrimSpace(binding.text)
					break
				}
			}
			if directArrayArgumentName(actual) == "" {
				return false
			}
			foundCaller = true
			if _, isParameter := arrayProcedureParameterIndex(candidate, actual); isParameter {
				if !arrayQualifiedDescriptorReceiverInitialized(file, candidate, nested, actual, ctx, requirePositive, visiting) {
					return false
				}
				continue
			}
			if !arrayQualifiedDescriptorLocalInitialized(file, candidate, nested, actual, requirePositive) {
				return false
			}
		}
	}
	return foundCaller
}

func arrayQualifiedDescriptorLocalInitialized(file parsedFile, proc sourceProcedure, call procedureir.CallSite, receiver string, requirePositive bool) bool {
	wantReceiver := canonicalArrayBoundExpression(receiver)
	wantData := wantReceiver + ".sa.pvdata"
	wantBounds := wantReceiver + ".sa.rgsabound0.celements"
	hasData := false
	count := ""
	for statement := range proc.Statements.All() {
		line := statement.Range.StartLine
		if line <= proc.StartLine || line >= call.Range.StartLine || !arrayStatementDominatesCall(proc, statement.ID, line, call) {
			continue
		}
		for _, text := range arrayQualifiedDescriptorStatementTexts(file, statement) {
			lhs, rhs, indexed, assigned := arrayAssignment(text)
			if !assigned || indexed {
				continue
			}
			switch canonicalArrayBoundExpression(lhs) {
			case wantData:
				canonical := canonicalArrayBoundExpression(rhs)
				hasData = canonical != "" && canonical != "0" && canonical != "nullptr"
			case wantBounds:
				count = canonicalArrayBoundExpression(rhs)
			}
		}
	}
	if !hasData || count == "" {
		return false
	}
	if !requirePositive {
		return true
	}
	if value, err := strconv.Atoi(count); err == nil {
		return value > 0
	}
	return arrayQualifiedDescriptorCountPositive(file, proc, call, count)
}

func arrayQualifiedDescriptorStatementTexts(file parsedFile, statement procedureir.Statement) []string {
	texts := make([]string, 0, 4)
	seen := map[string]bool{}
	add := func(source string) {
		for _, segment := range splitRangeValueSourceStatementsWithOffsets(arraySourceOrderStripComment(source)) {
			text := strings.TrimSpace(segment.text)
			if text != "" && !seen[text] {
				seen[text] = true
				texts = append(texts, text)
			}
		}
	}
	add(statement.Text)
	if statement.Range.StartLine >= 1 && statement.Range.StartLine <= len(file.Lines) {
		add(arrayLogicalSourceLine(file.Lines, statement.Range.StartLine))
	}
	return texts
}

func arrayQualifiedDescriptorHasSuccessfulShapeUse(file parsedFile, proc sourceProcedure, call procedureir.CallSite, receiver string) bool {
	wantCount := canonicalArrayBoundExpression(receiver + ".sa.rgsabound0.cElements")
	upperNames := map[string]int{}
	for statement := range proc.Statements.All() {
		line := statement.Range.StartLine
		if line <= proc.StartLine || line >= call.Range.StartLine || !arrayStatementDominatesCall(proc, statement.ID, line, call) {
			continue
		}
		for _, text := range arrayQualifiedDescriptorStatementTexts(file, statement) {
			if lhs, rhs, indexed, assigned := arrayAssignment(text); assigned && !indexed && canonicalArrayBoundExpression(rhs) == wantCount+"-1" {
				upperNames[strings.ToLower(cleanIdentifier(lhs))] = line
			}
			match := arrayRedimRe.FindStringSubmatch(strings.TrimSpace(text))
			if len(match) == 0 || strings.TrimSpace(match[1]) != "" {
				continue
			}
			for _, clause := range splitArgs(match[2]) {
				redim, direct := parseDirectArrayRedimClause(clause)
				if !direct || redim.name == "" {
					continue
				}
				dimensions := canonicalArrayBoundExpression(redim.dimensions)
				if !strings.HasPrefix(dimensions, "0to") {
					continue
				}
				upper := strings.TrimPrefix(dimensions, "0to")
				if assignmentLine, ok := upperNames[strings.ToLower(cleanIdentifier(upper))]; ok && assignmentLine < line {
					return true
				}
			}
		}
	}
	return false
}

// arrayQualifiedDescriptorCountPositive proves that the count used to build a
// descriptor-backed array is positive on the call path. It understands direct
// comparison branches and a zero guard whose then arm jumps to a label after
// the call (the normal-path shape used by parser entry points).
func arrayQualifiedDescriptorCountPositive(file parsedFile, proc sourceProcedure, call procedureir.CallSite, count string) bool {
	wantCount := canonicalArrayBoundExpression(count)
	if wantCount == "" {
		return false
	}
	for guard := range proc.Statements.All() {
		if guard.Kind != procedureir.StatementIf && guard.Kind != procedureir.StatementElseIf || !arrayStatementDominatesCall(proc, guard.ID, guard.Range.StartLine, call) {
			continue
		}
		condition := guard.Text
		if guard.Condition != nil && strings.TrimSpace(guard.Condition.Text) != "" {
			condition = guard.Condition.Text
		}
		lhs, operator, literal, ok := arrayQualifiedCountComparison(condition)
		if !ok || lhs != wantCount {
			continue
		}
		positiveBranch, positive := positiveArrayCountBranch(operator, literal)
		if !positive {
			continue
		}
		branch, underGuard := arrayQualifiedStatementBranch(proc, call.StatementID, guard.ID)
		if underGuard && arrayQualifiedBranchMatchesPositive(branch, positiveBranch) {
			return true
		}
		if operator == "=" && literal == "0" && !underGuard && arrayQualifiedZeroGuardSkipsCall(proc, guard, call) && arrayQualifiedCountHasNonNegativeOrigin(file, proc, call, count) {
			return true
		}
	}
	return false
}

func arrayQualifiedCountComparison(text string) (lhs, operator, literal string, ok bool) {
	text = strings.TrimSpace(text)
	if condition, _, hasThen := arrayIfThenParts(text); hasThen {
		text = condition
	}
	lower := strings.ToLower(text)
	if strings.HasPrefix(lower, "if ") {
		text = strings.TrimSpace(text[len("if "):])
	} else if strings.HasPrefix(lower, "elseif ") {
		text = strings.TrimSpace(text[len("elseif "):])
	}
	text = arrayQualifiedTrimOuterParens(text)
	lhs, operator, literal, ok = arrayCountComparison(text)
	if !ok {
		return "", "", "", false
	}
	lhs = canonicalArrayBoundExpression(arrayQualifiedTrimOuterParens(lhs))
	return lhs, operator, literal, lhs != ""
}

func arrayQualifiedTrimOuterParens(text string) string {
	text = strings.TrimSpace(text)
	for strings.HasPrefix(text, "(") {
		close := matchingParen(text, 0)
		if close != len(text)-1 {
			break
		}
		text = strings.TrimSpace(text[1:close])
	}
	return text
}

func arrayQualifiedStatementBranch(proc sourceProcedure, statementID, ancestorID int) (procedureir.BranchRole, bool) {
	seen := map[int]bool{}
	for statementID > 0 && !seen[statementID] {
		seen[statementID] = true
		statement, ok := arrayProcedureStatementByID(proc, statementID)
		if !ok {
			return "", false
		}
		if statement.ID == ancestorID {
			if statement.Kind == procedureir.StatementIf && statement.SyntaxKind == "single_line_if_statement" {
				return procedureir.BranchThen, true
			}
			return "", false
		}
		if statement.ParentID == ancestorID {
			if statement.Kind == procedureir.StatementElse {
				return procedureir.BranchElse, true
			}
			return procedureir.BranchThen, true
		}
		statementID = statement.ParentID
	}
	return "", false
}

func arrayQualifiedBranchMatchesPositive(branch procedureir.BranchRole, positive vbacfg.EdgeKind) bool {
	switch branch {
	case procedureir.BranchThen:
		return positive == vbacfg.EdgeBranchTrue
	case procedureir.BranchElse:
		return positive == vbacfg.EdgeBranchFalse
	default:
		return false
	}
}

func arrayQualifiedZeroGuardSkipsCall(proc sourceProcedure, guard procedureir.Statement, call procedureir.CallSite) bool {
	thenStatements := make([]procedureir.Statement, 0)
	for statement := range proc.Statements.All() {
		if statement.ParentID != guard.ID || statement.Kind == procedureir.StatementElseIf || statement.Kind == procedureir.StatementElse {
			continue
		}
		thenStatements = append(thenStatements, statement)
	}
	if len(thenStatements) == 0 {
		return false
	}
	sort.SliceStable(thenStatements, func(i, j int) bool {
		if thenStatements[i].Range.StartLine != thenStatements[j].Range.StartLine {
			return thenStatements[i].Range.StartLine < thenStatements[j].Range.StartLine
		}
		return thenStatements[i].ID < thenStatements[j].ID
	})
	for _, statement := range thenStatements[:len(thenStatements)-1] {
		if statement.Kind == procedureir.StatementIf || statement.Kind == procedureir.StatementElseIf || statement.Kind == procedureir.StatementElse || statement.Control != nil && statement.Control.Transfer != "" {
			return false
		}
	}
	return arrayQualifiedStatementLeavesBeforeCall(proc, thenStatements[len(thenStatements)-1], call)
}

func arrayQualifiedStatementLeavesBeforeCall(proc sourceProcedure, statement procedureir.Statement, call procedureir.CallSite) bool {
	if statement.Control != nil {
		switch statement.Control.Transfer {
		case procedureir.TransferExitSub, procedureir.TransferExitFunction, procedureir.TransferExitProperty, procedureir.TransferTerminate:
			return true
		case procedureir.TransferGoto:
			labelLine, ok := arrayQualifiedLabelLine(proc, statement.Control.Target)
			return ok && labelLine > call.Range.StartLine
		}
	}
	switch statement.Kind {
	case procedureir.StatementEnd:
		return true
	case procedureir.StatementExit:
		return arraySourceOrderProcedureExitStatementText(statement.Text)
	default:
		return false
	}
}

func arrayQualifiedLabelLine(proc sourceProcedure, target string) (int, bool) {
	target = strings.ToLower(strings.TrimSpace(strings.TrimSuffix(target, ":")))
	if target == "" {
		return 0, false
	}
	for statement := range proc.Statements.All() {
		if statement.Kind != procedureir.StatementLabel || !strings.EqualFold(arrayLocalGoSubLabelName(statement), target) {
			continue
		}
		return statement.Range.StartLine, true
	}
	return 0, false
}

func arrayQualifiedCountHasNonNegativeOrigin(file parsedFile, proc sourceProcedure, call procedureir.CallSite, count string) bool {
	want := canonicalArrayBoundExpression(count)
	hasAssignment := false
	hasSafeOrigin := false
	statements := make([]procedureir.Statement, 0, proc.Statements.Len())
	for statement := range proc.Statements.All() {
		statements = append(statements, statement)
	}
	sort.SliceStable(statements, func(i, j int) bool {
		if statements[i].Range.StartLine != statements[j].Range.StartLine {
			return statements[i].Range.StartLine < statements[j].Range.StartLine
		}
		return statements[i].ID < statements[j].ID
	})
	for _, statement := range statements {
		line := statement.Range.StartLine
		if line <= proc.StartLine || line >= call.Range.StartLine {
			continue
		}
		texts := []string{strings.TrimSpace(statement.Text)}
		if source := arrayLogicalSourceLine(file.Lines, line); source != "" && source != texts[0] {
			texts = append(texts, source)
		}
		matched := false
		assignmentSafe := false
		for _, text := range texts {
			lhs, rhs, indexed, assigned := arrayAssignment(text)
			if !assigned || indexed || canonicalArrayBoundExpression(lhs) != want {
				continue
			}
			matched = true
			canonical := canonicalArrayBoundExpression(rhs)
			candidateSafe := false
			switch {
			case arrayQualifiedNonNegativeCountOrigin(canonical):
				candidateSafe = true
			case canonical == want+"*2", canonical == "2*"+want:
				// A byte count multiplied by two remains non-negative after the
				// masked or LenB origin has been seen.
				candidateSafe = hasSafeOrigin
			case canonical == "0":
				// The default zero value is also non-negative.
				candidateSafe = true
			default:
				continue
			}
			assignmentSafe = assignmentSafe || candidateSafe
		}
		if !matched {
			continue
		}
		hasAssignment = true
		if !assignmentSafe {
			// An unknown later assignment invalidates an earlier safe origin;
			// otherwise a count such as `sizeB = -1` could inherit proof from a
			// preceding LenB assignment.  The source fallback above prevents a
			// truncated IR expression from hiding a safe full-line assignment.
			return false
		}
		hasSafeOrigin = true
	}
	return hasAssignment && hasSafeOrigin
}

func arrayQualifiedNonNegativeCountOrigin(canonical string) bool {
	canonical = canonicalArrayBoundExpression(canonical)
	for strings.HasPrefix(canonical, "clng(") {
		close := matchingParen(canonical, len("clng"))
		if close != len(canonical)-1 {
			break
		}
		canonical = canonical[len("clng("):close]
	}
	if strings.HasPrefix(canonical, "lenb(") && matchingParen(canonical, len("lenb")) == len(canonical)-1 {
		return true
	}
	for _, mask := range []string{"and&h7fffffff", "and&hffff&", "and&hffff", "and&hff&", "and&hff"} {
		if strings.HasSuffix(canonical, mask) {
			return true
		}
	}
	return false
}

func arrayQualifiedDictionarySnapshotProvenAllocated(file parsedFile, caller sourceProcedure, call procedureir.CallSite, argument string) bool {
	receiver, member, ok := arrayQualifiedMemberParts(argument)
	if !ok || !strings.EqualFold(member, "arrkeys") && !strings.EqualFold(member, "arritems") {
		return false
	}
	if receiver == "" {
		receiver = arrayWithReceiverAtLine(file, caller, call.Range.StartLine)
	}
	if receiver == "" {
		return false
	}
	arguments := arrayCallArgumentTexts(caller, call)
	if len(arguments) == 0 || !arrayQualifiedUpperBoundProvenNonNegative(file, caller, call, arguments[len(arguments)-1]) {
		return false
	}
	want := canonicalArrayBoundExpression(receiver + "." + member)
	for line := max(1, caller.StartLine); line < call.Range.StartLine && line <= len(file.Lines); line++ {
		source := arrayLogicalSourceLine(file.Lines, line)
		if source == "" {
			continue
		}
		if _, body, ok := arrayIfThenParts(source); ok {
			thenBody, elseBody, hasElse := arrayIfThenBodyParts(body)
			if hasElse && arrayQualifiedSnapshotAssignmentArm(file, caller, line, thenBody, want, member) && arrayQualifiedSnapshotAssignmentArm(file, caller, line, elseBody, want, member) && arrayQualifiedSourceLineDominatesCall(caller, line, call) {
				return true
			}
			continue
		}
		if arrayQualifiedSnapshotAssignmentArm(file, caller, line, source, want, member) && arrayQualifiedSourceLineDominatesCall(caller, line, call) {
			return true
		}
	}
	return false
}

func arrayQualifiedSnapshotAssignmentArm(file parsedFile, caller sourceProcedure, line int, text, want, member string) bool {
	lhs, rhs, indexed, assigned := arrayAssignment(text)
	if !assigned || indexed || arrayQualifiedArgumentTarget(file, caller, line, lhs) != want {
		return false
	}
	_, snapshotMember, ok := arrayDictionaryMemberParts(rhs)
	return ok && strings.EqualFold(snapshotMember, member[len("arr"):])
}

func arrayQualifiedSourceLineDominatesCall(proc sourceProcedure, line int, call procedureir.CallSite) bool {
	for statement := range proc.Statements.All() {
		if statement.Range.StartLine == line && arrayStatementDominatesCall(proc, statement.ID, line, call) {
			return true
		}
	}
	return false
}

func arrayQualifiedUpperBoundProvenNonNegative(file parsedFile, proc sourceProcedure, call procedureir.CallSite, argument string) bool {
	want := arrayQualifiedArgumentTarget(file, proc, call.Range.StartLine, argument)
	if want == "" {
		return false
	}
	for guard := range proc.Statements.All() {
		if guard.Kind != procedureir.StatementIf && guard.Kind != procedureir.StatementElseIf || !arrayStatementDominatesCall(proc, guard.ID, guard.Range.StartLine, call) {
			continue
		}
		condition := guard.Text
		if guard.Condition != nil && strings.TrimSpace(guard.Condition.Text) != "" {
			condition = guard.Condition.Text
		}
		lhs, operator, literal, ok := arrayQualifiedCountComparison(condition)
		if !ok {
			continue
		}
		positiveBranch, positive := positiveArrayCountBranch(operator, literal)
		if !positive {
			continue
		}
		callBranch, underGuard := arrayQualifiedStatementBranch(proc, call.StatementID, guard.ID)
		if !underGuard || !arrayQualifiedBranchMatchesPositive(callBranch, positiveBranch) {
			continue
		}
		for statement := range proc.Statements.All() {
			line := statement.Range.StartLine
			if line <= proc.StartLine || line >= call.Range.StartLine || !arrayStatementDominatesCall(proc, statement.ID, line, call) {
				continue
			}
			statementBranch, sameBranch := arrayQualifiedStatementBranch(proc, statement.ID, guard.ID)
			if !sameBranch || statementBranch != callBranch {
				continue
			}
			text := strings.TrimSpace(statement.Text)
			if text == "" {
				text = arrayLogicalSourceLine(file.Lines, line)
			}
			assignment, rhs, indexed, assigned := arrayAssignment(text)
			if assigned && !indexed && arrayQualifiedArgumentTarget(file, proc, line, assignment) == want && canonicalArrayBoundExpression(rhs) == lhs {
				return true
			}
		}
	}
	return false
}

func arrayQualifiedArgumentTarget(file parsedFile, caller sourceProcedure, line int, argument string) string {
	receiver, member, ok := arrayQualifiedMemberParts(argument)
	if !ok {
		return ""
	}
	if receiver == "" {
		receiver = arrayWithReceiverAtLine(file, caller, line)
	}
	if receiver == "" {
		return ""
	}
	return canonicalArrayBoundExpression(receiver + "." + member)
}

func arrayQualifiedMemberParts(text string) (receiver, member string, ok bool) {
	text = strings.TrimSpace(text)
	if text == "" {
		return "", "", false
	}
	dot := strings.LastIndexByte(text, '.')
	if dot < 0 || dot >= len(text)-1 {
		return "", "", false
	}
	member = cleanIdentifier(strings.TrimSpace(text[dot+1:]))
	if !arrayEraseNameRe.MatchString(member) {
		return "", "", false
	}
	receiver = strings.TrimSpace(text[:dot])
	if receiver == "" {
		if text[0] != '.' {
			return "", "", false
		}
		return "", member, true
	}
	for _, part := range strings.Split(receiver, ".") {
		if !arrayEraseNameRe.MatchString(strings.TrimSpace(part)) {
			return "", "", false
		}
	}
	return receiver, member, true
}

func arrayQualifiedRedimAllocatesTarget(file parsedFile, caller sourceProcedure, line int, text, want string) bool {
	match := arrayRedimRe.FindStringSubmatch(strings.TrimSpace(text))
	if len(match) == 0 || strings.TrimSpace(match[1]) != "" {
		return false
	}
	for _, clause := range splitArgs(match[2]) {
		clause = strings.TrimSpace(clause)
		open := firstParenOutsideString(clause)
		if open <= 0 {
			continue
		}
		close := matchingParen(clause, open)
		if close < 0 {
			continue
		}
		remainder := strings.TrimSpace(clause[close+1:])
		if remainder != "" && !arrayRedimTypeSuffixRe.MatchString(remainder) {
			continue
		}
		target := arrayQualifiedArgumentTarget(file, caller, line, strings.TrimSpace(clause[:open]))
		if target == "" || target != want {
			continue
		}
		dimensions := parseArrayDimensionsWithConstants(clause[open+1:close], arrayOptionBase(file), nil)
		if arrayDimensionsKnownNonEmpty(dimensions) {
			return true
		}
	}
	return false
}

func arrayDimensionsKnownNonEmpty(dimensions []arrayDimension) bool {
	// A plain ReDim that reaches the following statement has established an
	// allocation. Unknown bounds may make ReDim fail before that continuation,
	// but they do not make the successfully reached array empty unless the
	// bounds are provably impossible. Preserve the existing conservative check
	// for the latter case while allowing runtime bounds such as `0 To .ub`.
	return len(dimensions) > 0 && !impossibleArrayBounds(dimensions)
}

func arrayStatementDominatesCall(proc sourceProcedure, statementID, statementLine int, call procedureir.CallSite) bool {
	if proc.Graph == nil || statementID <= 0 || call.StatementID <= 0 {
		return false
	}
	statementBlock, statementOK := proc.Graph.BlockForStatement(statementID)
	callBlock, callOK := proc.Graph.BlockForStatement(call.StatementID)
	if !statementOK || !callOK {
		return false
	}
	if statementBlock.ID == callBlock.ID {
		return statementLine < call.Range.StartLine
	}
	for _, dominator := range proc.Graph.View(vbacfg.EdgeFilter{NormalOnly: true}).DominatorsOf(callBlock.ID) {
		if dominator == statementBlock.ID {
			return true
		}
	}
	return false
}

func parameterIsByRefArray(parameter parameterInfo) bool {
	// ParamArray materializes a new Variant array in the callee. Its actual
	// arguments are elements of that array, not aliases to caller arrays.
	return !parameter.ParamArray && parameterIsArray(parameter) && !strings.EqualFold(strings.TrimSpace(parameter.Passing), "ByVal")
}

func parameterIsByRefScalar(parameter parameterInfo) bool {
	return !parameterIsArray(parameter) && !strings.EqualFold(strings.TrimSpace(parameter.Passing), "ByVal")
}
