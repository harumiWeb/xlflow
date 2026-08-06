package analyze

import (
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

const (
	excelAccessRead          = "read"
	excelAccessWrite         = "write"
	excelAccessFormatting    = "formatting"
	excelAccessRangeLookup   = "range lookup"
	excelAccessWorksheetCall = "worksheet function"
)

var (
	constIntegerRe      = regexp.MustCompile(`(?i)^\s*(?:public\s+|private\s+|friend\s+|static\s+)?const\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:as\s+[A-Za-z_][A-Za-z0-9_]*)?\s*=\s*(.+?)\s*$`)
	excelCellCallRe     = regexp.MustCompile(`(?i)(^|[^a-z0-9_.])cells\s*\(`)
	directExcelLookupRe = regexp.MustCompile(`(?i)(^|[^a-z0-9_.])(?:range|worksheets|sheets)\s*\(`)
	forEachRangeRe      = regexp.MustCompile(`(?i)^\s*for\s+each\s+([A-Za-z_][A-Za-z0-9_]*)\s+in\s+(.+?)\s*$`)
	forBoundsRe         = regexp.MustCompile(`(?i)^\s*for\s+[A-Za-z_][A-Za-z0-9_]*\s*=\s*(.+?)\s+to\s+(.+?)(?:\s+step\s+(.+?))?\s*$`)
	literalRangeRe      = regexp.MustCompile(`(?i)\b(?:range|cells)\s*\(\s*"([A-Z]+)([0-9]+)(?::([A-Z]+)([0-9]+))?"\s*\)`)
)

var formattingMembers = map[string]bool{
	"bold": true, "italic": true, "underline": true, "color": true,
	"colorindex": true, "pattern": true, "patterncolor": true,
	"numberformat": true, "numberformatlocal": true, "horizontalalignment": true,
	"verticalalignment": true, "orientation": true, "wraptext": true,
	"shrinktofit": true, "columnwidth": true, "rowheight": true, "style": true,
	"linestyle": true, "lineweight": true, "themecolor": true, "tintandshade": true,
}

var rangeLookupMembers = map[string]bool{
	"cells": true, "offset": true, "range": true, "rows": true, "columns": true,
	"item": true, "worksheets": true, "sheets": true,
}

var cellValueMembers = map[string]bool{
	"value": true, "value2": true, "formula": true, "formular1c1": true,
}

var mutatingRangeMembers = map[string]bool{
	"clear": true, "clearcontents": true, "clearformats": true, "delete": true,
	"copy": true, "pastespecial": true,
}

type excelLoopAccess struct {
	Category string
	Member   string
	Line     int
	Read     bool
	Write    bool
	Helper   string
}

type excelLoopRegion struct {
	StatementID int
	Line        int
	EndLine     int
	Depth       int
	Body        map[int]bool
	Small       bool
}

type excelAccessSummary struct {
	Categories map[string]bool
	Members    map[string]bool
}

type excelRangeVariables struct {
	Range   map[string]bool
	PerCell map[string]bool
}

type excelLoopAccessIndex struct {
	Summaries map[string]excelAccessSummary
}

func buildExcelLoopAccessIndex(files []parsedFile, db *vbadb.DB, rootDir string, cfg config.Config) *excelLoopAccessIndex {
	index := &excelLoopAccessIndex{Summaries: map[string]excelAccessSummary{}}
	if db == nil {
		return index
	}
	for _, file := range files {
		procedures := sourceProceduresFromIR(file.IR, file.CFG)
		for i, proc := range procedures {
			if i >= len(file.IR.Procedures) {
				continue
			}
			key := excelProcedureKey(file.IR, file.IR.Procedures[i].Symbol)
			index.Summaries[key] = directExcelAccessSummary(file, proc, db, rootDir, cfg)
		}
	}

	// Propagate only uniquely resolved project-local summaries. The summary is
	// a finite set of categories, so repeated fixed-point passes are bounded.
	changed := true
	for changed {
		changed = false
		for _, file := range files {
			for _, procedure := range file.IR.Procedures {
				key := excelProcedureKey(file.IR, procedure.Symbol)
				current := index.Summaries[key]
				for _, call := range procedure.Calls {
					if call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
						continue
					}
					callee, ok := index.Summaries[excelCandidateKey(call.Resolution.Candidates[0])]
					if !ok {
						continue
					}
					if mergeExcelSummary(&current, callee) {
						changed = true
					}
				}
				index.Summaries[key] = current
			}
		}
	}
	return index
}

func excelProcedureKey(document procedureir.DocumentIR, symbol procedureir.ProcedureSymbol) string {
	return strings.Join([]string{
		strings.ToLower(document.Path), strings.ToLower(symbol.QualifiedName),
		strings.ToLower(string(symbol.Kind)), strconv.Itoa(symbol.DeclarationRange.StartLine),
	}, "\x00")
}

func excelCandidateKey(candidate procedureir.Candidate) string {
	return strings.Join([]string{
		strings.ToLower(candidate.File), strings.ToLower(candidate.QualifiedName),
		strings.ToLower(candidate.Kind), strconv.Itoa(candidate.Line),
	}, "\x00")
}

func mergeExcelSummary(dst *excelAccessSummary, src excelAccessSummary) bool {
	changed := false
	if dst.Categories == nil {
		dst.Categories = map[string]bool{}
	}
	if dst.Members == nil {
		dst.Members = map[string]bool{}
	}
	for category := range src.Categories {
		if !dst.Categories[category] {
			dst.Categories[category] = true
			changed = true
		}
	}
	for member := range src.Members {
		if !dst.Members[member] {
			dst.Members[member] = true
			changed = true
		}
	}
	return changed
}

func directExcelAccessSummary(file parsedFile, proc sourceProcedure, db *vbadb.DB, rootDir string, cfg config.Config) excelAccessSummary {
	summary := excelAccessSummary{Categories: map[string]bool{}, Members: map[string]bool{}}
	loopVars := rangeVariablesForProcedure(proc, file, db, rootDir, cfg)
	for _, statement := range proc.Statements {
		for _, access := range classifyExcelStatement(file, proc, statement, db, loopVars, rootDir, cfg) {
			summary.Categories[access.Category] = true
			if access.Member != "" {
				summary.Members[strings.ToLower(access.Member)] = true
			}
		}
	}
	return summary
}

func (a Analyzer) excelLoopAccessFindings(file parsedFile, proc sourceProcedure) []Finding {
	if !a.Config.Analyze.DetectExcelCellAccessInLoops || a.typeDB == nil {
		return nil
	}
	regions := excelLoopRegions(proc)
	if len(regions) == 0 {
		return nil
	}
	summaries := map[string]excelAccessSummary{}
	needHelperSummaries := a.excelLoopAccess == nil && excelProcedureHasLocalLoopCall(file, proc, regions)
	if a.excelLoopAccess != nil {
		summaries = a.excelLoopAccess.Summaries
	} else if needHelperSummaries {
		// Realtime analysis owns one document and does not have the batch
		// resolver. Build same-document summaries so local helpers still work.
		summaries = buildRealtimeExcelLoopSummaries(file, a.typeDB, a.RootDir, a.Config)
	}
	loopVars := rangeVariablesForProcedure(proc, file, a.typeDB, a.RootDir, a.Config)
	byStatement := map[int][]excelLoopAccess{}
	for _, statement := range proc.Statements {
		accesses := classifyExcelStatement(file, proc, statement, a.typeDB, loopVars, a.RootDir, a.Config)
		if len(accesses) > 0 {
			byStatement[statement.ID] = accesses
		}
	}
	for _, call := range proc.Calls {
		if a.excelLoopAccess != nil && (call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1) {
			continue
		}
		if a.excelLoopAccess == nil && !needHelperSummaries {
			continue
		}
		summary, ok := excelHelperSummary(file, call, summaries)
		if ok && len(summary.Categories) > 0 {
			memberNames := sortedStringSet(summary.Members)
			for category := range summary.Categories {
				read := category == excelAccessRead || category == excelAccessRangeLookup || category == excelAccessWorksheetCall
				write := category == excelAccessWrite || category == excelAccessFormatting
				if len(memberNames) == 0 {
					byStatement[call.StatementID] = append(byStatement[call.StatementID], excelLoopAccess{
						Category: category, Line: call.Range.StartLine, Helper: call.Callee.Text, Read: read, Write: write,
					})
					continue
				}
				for _, member := range memberNames {
					byStatement[call.StatementID] = append(byStatement[call.StatementID], excelLoopAccess{
						Category: category, Member: member, Line: call.Range.StartLine, Helper: call.Callee.Text, Read: read, Write: write,
					})
				}
			}
		}
	}

	grouped := map[int][]excelLoopAccess{}
	for statementID, accesses := range byStatement {
		line := 0
		if len(accesses) > 0 {
			line = accesses[0].Line
		}
		candidates := containingExcelLoops(regions, statementID, line)
		if len(candidates) == 0 {
			continue
		}
		selected := -1
		for i := len(candidates) - 1; i >= 0; i-- {
			if !candidates[i].Small {
				selected = i
				break
			}
		}
		if selected < 0 {
			continue
		}
		region := candidates[selected]
		grouped[region.StatementID] = append(grouped[region.StatementID], accesses...)
	}

	ordered := append([]excelLoopRegion(nil), regions...)
	sort.SliceStable(ordered, func(i, j int) bool { return ordered[i].Line < ordered[j].Line })
	var findings []Finding
	for _, region := range ordered {
		accesses := grouped[region.StatementID]
		if len(accesses) == 0 {
			continue
		}
		categories := map[string]bool{}
		members := map[string]bool{}
		helpers := map[string]bool{}
		read, write := false, false
		for _, access := range accesses {
			categories[access.Category] = true
			if access.Member != "" {
				members[strings.ToLower(access.Member)] = true
			}
			if access.Helper != "" {
				helpers[access.Helper] = true
			}
			read = read || access.Read
			write = write || access.Write
		}
		categoryNames := sortedExcelCategories(categories)
		helperNames := sortedStringSet(helpers)
		severity := "warning"
		if region.Depth >= 2 {
			severity = "error"
		}
		message := "Loop performs repeated Excel object-model access"
		if len(categoryNames) > 0 {
			message += " (" + strings.Join(categoryNames, ", ") + ")"
		}
		message += "."
		memberNames := sortedStringSet(members)
		if len(memberNames) > 0 {
			message += " Members: " + strings.Join(memberNames, ", ") + "."
		}
		if region.Depth >= 2 {
			message += " Nested loop depth: " + strconv.Itoa(region.Depth) + "."
		}
		if len(helperNames) > 0 {
			message += " Helper: " + strings.Join(helperNames, ", ") + "."
		}
		direction := ""
		if read && write {
			direction = "read and write"
		} else if read {
			direction = "read"
		} else if write {
			direction = "write"
		}
		reason := "Each iteration crosses the VBA-to-Excel COM boundary for " + direction + " work"
		if direction == "" {
			reason = "Each iteration crosses the VBA-to-Excel COM boundary"
		}
		reason += "; nested loops multiply the number of round trips."
		suggestion := "Read the range once into a Variant array with Value2, process values in memory, and write the array back once. Cache Worksheet and Range objects outside the loop, and apply formatting to the complete range."
		if categories[excelAccessWorksheetCall] {
			suggestion = "Move the WorksheetFunction call out of the per-cell loop where possible, pass a range or array to one operation, and cache the resolved Worksheet/Range objects."
		}
		finding := a.simpleFinding(file, proc, region.Line, "VBA225", severity, message, reason, suggestion)
		finding.EndLine = region.EndLine
		findings = append(findings, finding)
	}
	return findings
}

func excelProcedureHasLocalLoopCall(file parsedFile, proc sourceProcedure, regions []excelLoopRegion) bool {
	localNames := map[string]int{}
	for _, procedure := range file.IR.Procedures {
		localNames[strings.ToLower(procedure.Symbol.Name)]++
	}
	for _, call := range proc.Calls {
		if call.Callee.Receiver != nil || strings.TrimSpace(call.Callee.BaseName) == "" {
			continue
		}
		if localNames[strings.ToLower(call.Callee.BaseName)] != 1 {
			continue
		}
		for _, region := range regions {
			if region.Body[call.StatementID] || region.StatementID == call.StatementID {
				return true
			}
		}
	}
	return false
}

func buildRealtimeExcelLoopSummaries(file parsedFile, db *vbadb.DB, rootDir string, cfg config.Config) map[string]excelAccessSummary {
	summaries := map[string]excelAccessSummary{}
	procedures := sourceProceduresFromIR(file.IR, file.CFG)
	for i, candidate := range file.IR.Procedures {
		if i >= len(procedures) {
			continue
		}
		key := excelProcedureKey(file.IR, candidate.Symbol)
		summaries[key] = directExcelAccessSummary(file, procedures[i], db, rootDir, cfg)
	}
	changed := true
	for changed {
		changed = false
		for _, procedure := range file.IR.Procedures {
			key := excelProcedureKey(file.IR, procedure.Symbol)
			current := summaries[key]
			for _, call := range procedure.Calls {
				callee, ok := excelHelperProcedureKey(file, call)
				if !ok {
					continue
				}
				if summary, found := summaries[callee]; found && mergeExcelSummary(&current, summary) {
					changed = true
				}
			}
			summaries[key] = current
		}
	}
	return summaries
}

func excelHelperSummary(file parsedFile, call procedureir.CallSite, summaries map[string]excelAccessSummary) (excelAccessSummary, bool) {
	if call.Resolution.Status == procedureir.ResolutionMatched && len(call.Resolution.Candidates) == 1 {
		if summary, ok := summaries[excelCandidateKey(call.Resolution.Candidates[0])]; ok {
			return summary, true
		}
	}
	key, ok := excelHelperProcedureKey(file, call)
	if !ok {
		return excelAccessSummary{}, false
	}
	summary, ok := summaries[key]
	return summary, ok
}

func excelHelperProcedureKey(file parsedFile, call procedureir.CallSite) (string, bool) {
	if call.Callee.Receiver != nil {
		return "", false
	}
	name := strings.TrimSpace(call.Callee.BaseName)
	if name == "" {
		name = strings.TrimSpace(call.Callee.Text)
	}
	name = strings.TrimSpace(strings.TrimPrefix(name, "New "))
	if name == "" {
		return "", false
	}
	var matches []procedureir.ProcedureSymbol
	for _, procedure := range file.IR.Procedures {
		if strings.EqualFold(procedure.Symbol.Name, name) || strings.EqualFold(procedure.Symbol.QualifiedName, name) {
			matches = append(matches, procedure.Symbol)
		}
	}
	if len(matches) != 1 {
		return "", false
	}
	return excelProcedureKey(file.IR, matches[0]), true
}

func excelLoopRegions(proc sourceProcedure) []excelLoopRegion {
	constants := excelIntegerConstants(proc)
	adjacency := excelCFGAdjacency(proc.Graph)
	blockByStatement := excelCFGBlocksByStatement(proc.Graph)
	statementByBlock := excelCFGStatementsByBlock(proc.Graph)
	children := map[int][]int{}
	statements := map[int]procedureir.Statement{}
	for _, statement := range proc.Statements {
		children[statement.ParentID] = append(children[statement.ParentID], statement.ID)
		statements[statement.ID] = statement
	}
	regions := make([]excelLoopRegion, 0)
	for _, statement := range proc.Statements {
		if !isExcelLoopKind(statement.Kind) || statement.Recovered {
			continue
		}
		body := map[int]bool{}
		var visit func(int)
		visit = func(parent int) {
			for _, child := range children[parent] {
				if body[child] {
					continue
				}
				body[child] = true
				visit(child)
			}
		}
		visit(statement.ID)
		if !excelLoopCFGValid(proc, statement.ID, body, adjacency, blockByStatement) {
			continue
		}
		reachableBody, ok := excelLoopReachableStatements(proc, statement.ID, body, adjacency, blockByStatement, statementByBlock)
		if !ok {
			continue
		}
		body = reachableBody
		endLine := statement.Range.StartLine
		for id := range body {
			if child, ok := statements[id]; ok && child.Range.StartLine > endLine {
				endLine = child.Range.StartLine
			}
		}
		regions = append(regions, excelLoopRegion{
			StatementID: statement.ID,
			Line:        statement.Range.StartLine,
			EndLine:     endLine,
			Body:        body,
			Small:       excelLoopIsSmall(statement.Text, constants),
		})
	}
	for i := range regions {
		depth := 1
		for j := range regions {
			if i == j || !regions[j].Body[regions[i].StatementID] {
				continue
			}
			depth++
		}
		regions[i].Depth = depth
	}
	sort.SliceStable(regions, func(i, j int) bool {
		if regions[i].Depth != regions[j].Depth {
			return regions[i].Depth < regions[j].Depth
		}
		return regions[i].Line < regions[j].Line
	})
	return regions
}

func excelIntegerConstants(proc sourceProcedure) map[string]int {
	constants := map[string]int{}
	for _, statement := range proc.Statements {
		match := constIntegerRe.FindStringSubmatch(strings.TrimSpace(statement.Text))
		if len(match) != 3 {
			continue
		}
		if value, err := constantIntegerExpression(match[2], constants); err == nil {
			constants[strings.ToLower(match[1])] = value
		}
	}
	return constants
}

func excelCFGAdjacency(graph *vbacfg.Graph) map[vbacfg.BlockID][]vbacfg.Edge {
	adjacency := map[vbacfg.BlockID][]vbacfg.Edge{}
	if graph == nil {
		return adjacency
	}
	for _, edge := range graph.Edges {
		adjacency[edge.From] = append(adjacency[edge.From], edge)
	}
	return adjacency
}

func excelCFGBlocksByStatement(graph *vbacfg.Graph) map[int]vbacfg.BlockID {
	blocks := map[int]vbacfg.BlockID{}
	if graph == nil {
		return blocks
	}
	for _, block := range graph.Blocks {
		if block.StatementID != 0 {
			blocks[block.StatementID] = block.ID
		}
	}
	return blocks
}

func excelCFGStatementsByBlock(graph *vbacfg.Graph) map[vbacfg.BlockID]int {
	statements := map[vbacfg.BlockID]int{}
	if graph == nil {
		return statements
	}
	for _, block := range graph.Blocks {
		if block.StatementID != 0 {
			statements[block.ID] = block.StatementID
		}
	}
	return statements
}

func excelLoopCFGValid(proc sourceProcedure, loopID int, body map[int]bool, adjacency map[vbacfg.BlockID][]vbacfg.Edge, blockByStatement map[int]vbacfg.BlockID) bool {
	if proc.Graph == nil {
		return false
	}
	graph := proc.Graph
	header, ok := blockByStatement[loopID]
	if !ok {
		return false
	}
	bodyBlocks := map[vbacfg.BlockID]bool{}
	for statementID := range body {
		if block, ok := blockByStatement[statementID]; ok {
			bodyBlocks[block] = true
		}
	}
	for _, unknown := range graph.UnknownFlowSources {
		if unknown == header || bodyBlocks[unknown] {
			return false
		}
	}
	hasBody, hasBack, hasExit := false, false, false
	var loopEdges []vbacfg.Edge
	for from := range bodyBlocks {
		loopEdges = append(loopEdges, adjacency[from]...)
	}
	loopEdges = append(loopEdges, adjacency[header]...)
	for _, edge := range loopEdges {
		if edge.Uncertain {
			if edge.Kind == vbacfg.EdgeLoopBody || edge.Kind == vbacfg.EdgeLoopBack || edge.Kind == vbacfg.EdgeLoopExit {
				return false
			}
			continue
		}
		switch edge.Kind {
		case vbacfg.EdgeLoopBody:
			if edge.From == header || bodyBlocks[edge.From] {
				hasBody = true
			}
		case vbacfg.EdgeLoopBack:
			if edge.To == header || (bodyBlocks[edge.From] && bodyBlocks[edge.To]) {
				hasBack = true
			}
		case vbacfg.EdgeLoopExit:
			if edge.From == header || bodyBlocks[edge.From] {
				hasExit = true
			}
		}
	}
	if !hasBody || !hasBack {
		return false
	}
	if hasExit {
		return true
	}
	// For a nested For, the CFG builder can represent the inner exit as a
	// loop-back edge to the enclosing loop header. In that shape the back edge
	// itself closes the definite loop boundary even without a separate exit edge.
	for _, edge := range loopEdges {
		if !edge.Uncertain && edge.Kind == vbacfg.EdgeLoopBack && edge.To == header {
			return true
		}
	}
	return false
}

func excelLoopReachableStatements(proc sourceProcedure, loopID int, body map[int]bool, adjacency map[vbacfg.BlockID][]vbacfg.Edge, blockByStatement map[int]vbacfg.BlockID, statementByBlock map[vbacfg.BlockID]int) (map[int]bool, bool) {
	if proc.Graph == nil {
		return nil, false
	}
	header, ok := blockByStatement[loopID]
	if !ok {
		return nil, false
	}
	bodyBlocks := map[vbacfg.BlockID]bool{}
	for statementID := range body {
		if block, ok := blockByStatement[statementID]; ok {
			bodyBlocks[block] = true
		}
	}
	initial := map[vbacfg.BlockID]bool{}
	seenDiscovery := map[vbacfg.BlockID]bool{header: true}
	discovery := []vbacfg.BlockID{header}
	for len(discovery) > 0 {
		from := discovery[0]
		discovery = discovery[1:]
		for _, edge := range adjacency[from] {
			if edge.From != from || edge.Uncertain || edge.Class != vbacfg.EdgeNormal {
				continue
			}
			if edge.Kind == vbacfg.EdgeLoopExit || edge.Kind == vbacfg.EdgeLoopBack {
				continue
			}
			if edge.Kind == vbacfg.EdgeLoopBody {
				if bodyBlocks[edge.To] {
					initial[edge.To] = true
				}
				continue
			}
			if edge.To == header || !bodyBlocks[edge.To] || seenDiscovery[edge.To] {
				continue
			}
			seenDiscovery[edge.To] = true
			discovery = append(discovery, edge.To)
		}
	}
	if len(initial) == 0 {
		return nil, false
	}
	seen := map[vbacfg.BlockID]bool{}
	queue := make([]vbacfg.BlockID, 0, len(initial))
	for block := range initial {
		seen[block] = true
		queue = append(queue, block)
	}
	for len(queue) > 0 {
		from := queue[0]
		queue = queue[1:]
		for _, edge := range adjacency[from] {
			if edge.From != from || edge.Uncertain || edge.Class != vbacfg.EdgeNormal {
				continue
			}
			if edge.Kind == vbacfg.EdgeLoopExit || edge.To == header || !bodyBlocks[edge.To] || seen[edge.To] {
				continue
			}
			seen[edge.To] = true
			queue = append(queue, edge.To)
		}
	}
	reachable := map[int]bool{}
	for block := range seen {
		if statementID := statementByBlock[block]; statementID != 0 {
			reachable[statementID] = true
		}
	}
	return reachable, len(reachable) > 0
}

func containingExcelLoops(regions []excelLoopRegion, statementID, line int) []excelLoopRegion {
	var out []excelLoopRegion
	for _, region := range regions {
		if region.Body[statementID] || region.StatementID == statementID || line > region.Line && line <= region.EndLine {
			out = append(out, region)
		}
	}
	sort.SliceStable(out, func(i, j int) bool {
		if out[i].Depth != out[j].Depth {
			return out[i].Depth < out[j].Depth
		}
		return out[i].Line < out[j].Line
	})
	return out
}

func isExcelLoopKind(kind procedureir.StatementKind) bool {
	switch kind {
	case procedureir.StatementFor, procedureir.StatementForEach, procedureir.StatementDo, procedureir.StatementWhile:
		return true
	default:
		return false
	}
}

func excelLoopIsSmall(text string, constants map[string]int) bool {
	text = excelLoopHeaderText(text)
	match := forBoundsRe.FindStringSubmatch(strings.TrimSpace(text))
	if len(match) > 0 {
		start, err1 := constantIntegerExpression(match[1], constants)
		end, err2 := constantIntegerExpression(match[2], constants)
		step := 1
		var err3 error
		if match[3] != "" {
			step, err3 = constantIntegerExpression(match[3], constants)
		}
		if err1 == nil && err2 == nil && err3 == nil && step != 0 {
			count := 0
			if (step > 0 && start <= end) || (step < 0 && start >= end) {
				if step > 0 {
					count = (end-start)/step + 1
				} else {
					count = (start-end)/(-step) + 1
				}
			}
			return count <= 3
		}
	}
	if match := forEachRangeRe.FindStringSubmatch(strings.TrimSpace(text)); len(match) > 0 {
		count := literalRangeCellCount(match[2])
		return count > 0 && count <= 3
	}
	return false
}

type integerExpressionParser struct {
	text      string
	position  int
	constants map[string]int
}

func constantIntegerExpression(text string, constants map[string]int) (int, error) {
	parser := &integerExpressionParser{text: strings.TrimSpace(text), constants: constants}
	value, err := parser.parseExpression()
	if err != nil {
		return 0, err
	}
	parser.skipSpace()
	if parser.position != len(parser.text) {
		return 0, strconv.ErrSyntax
	}
	return value, nil
}

func (p *integerExpressionParser) parseExpression() (int, error) {
	value, err := p.parseTerm()
	if err != nil {
		return 0, err
	}
	for {
		p.skipSpace()
		if p.position >= len(p.text) {
			return value, nil
		}
		op := p.text[p.position]
		if op != '+' && op != '-' {
			return value, nil
		}
		p.position++
		right, err := p.parseTerm()
		if err != nil {
			return 0, err
		}
		if op == '+' {
			value += right
		} else {
			value -= right
		}
	}
}

func (p *integerExpressionParser) parseTerm() (int, error) {
	value, err := p.parseFactor()
	if err != nil {
		return 0, err
	}
	for {
		p.skipSpace()
		if p.position >= len(p.text) || (p.text[p.position] != '*' && p.text[p.position] != '/') {
			return value, nil
		}
		op := p.text[p.position]
		p.position++
		right, err := p.parseFactor()
		if err != nil {
			return 0, err
		}
		if op == '*' {
			value *= right
		} else {
			if right == 0 {
				return 0, strconv.ErrSyntax
			}
			value /= right
		}
	}
}

func (p *integerExpressionParser) parseFactor() (int, error) {
	p.skipSpace()
	if p.position >= len(p.text) {
		return 0, strconv.ErrSyntax
	}
	if p.text[p.position] == '+' || p.text[p.position] == '-' {
		negative := p.text[p.position] == '-'
		p.position++
		value, err := p.parseFactor()
		if negative {
			value = -value
		}
		return value, err
	}
	if p.text[p.position] == '(' {
		p.position++
		value, err := p.parseExpression()
		p.skipSpace()
		if err != nil || p.position >= len(p.text) || p.text[p.position] != ')' {
			return 0, strconv.ErrSyntax
		}
		p.position++
		return value, nil
	}
	start := p.position
	for p.position < len(p.text) && ((p.text[p.position] >= '0' && p.text[p.position] <= '9') || (p.text[p.position] >= 'A' && p.text[p.position] <= 'Z') || (p.text[p.position] >= 'a' && p.text[p.position] <= 'z') || p.text[p.position] == '_') {
		p.position++
	}
	if start == p.position {
		return 0, strconv.ErrSyntax
	}
	token := p.text[start:p.position]
	if value, err := strconv.Atoi(token); err == nil {
		return value, nil
	}
	if value, ok := p.constants[strings.ToLower(token)]; ok {
		return value, nil
	}
	return 0, strconv.ErrSyntax
}

func (p *integerExpressionParser) skipSpace() {
	for p.position < len(p.text) && (p.text[p.position] == ' ' || p.text[p.position] == '\t') {
		p.position++
	}
}

func excelLoopHeaderText(text string) string {
	if newline := strings.IndexAny(text, "\r\n"); newline >= 0 {
		return text[:newline]
	}
	return text
}

func literalRangeCellCount(expr string) int {
	match := literalRangeRe.FindStringSubmatch(expr)
	if len(match) == 0 {
		return 0
	}
	startCol := excelColumnNumber(match[1])
	startRow, _ := strconv.Atoi(match[2])
	endCol, endRow := startCol, startRow
	if match[3] != "" {
		endCol = excelColumnNumber(match[3])
		endRow, _ = strconv.Atoi(match[4])
	}
	if startCol <= 0 || endCol < startCol || endRow < startRow {
		return 0
	}
	return (endCol - startCol + 1) * (endRow - startRow + 1)
}

func excelColumnNumber(value string) int {
	total := 0
	for _, r := range strings.ToUpper(value) {
		if r < 'A' || r > 'Z' {
			return 0
		}
		total = total*26 + int(r-'A'+1)
	}
	return total
}

func rangeVariablesForProcedure(proc sourceProcedure, file parsedFile, db *vbadb.DB, rootDir string, cfg config.Config) excelRangeVariables {
	vars := excelRangeVariables{Range: map[string]bool{}, PerCell: map[string]bool{}}
	for _, declaration := range proc.Declarations {
		if isExcelRangeType(declaration.Type) {
			vars.Range[strings.ToLower(declaration.Name)] = true
		}
	}
	for _, statement := range proc.Statements {
		match := forEachRangeRe.FindStringSubmatch(strings.TrimSpace(excelLoopHeaderText(statement.Text)))
		if len(match) == 0 || !strings.Contains(strings.ToLower(match[2]), ".cells") {
			continue
		}
		receiver := strings.TrimSpace(strings.Split(strings.ToLower(match[2]), ".cells")[0])
		if isExcelRangeExpression(file, db, receiver, statement.Range.StartLine-1, rootDir, cfg) {
			name := strings.ToLower(match[1])
			vars.Range[name] = true
			vars.PerCell[name] = true
		}
	}
	return vars
}

func classifyExcelStatement(file parsedFile, proc sourceProcedure, statement procedureir.Statement, db *vbadb.DB, rangeVars excelRangeVariables, rootDir string, cfg config.Config) []excelLoopAccess {
	line := statement.Range.StartLine
	rawText := statement.Text
	text := rawText
	if isExcelLoopKind(statement.Kind) {
		text = excelLoopHeaderText(text)
	}
	text = sanitizeExcelStatementText(text)
	lower := strings.ToLower(text)
	var out []excelLoopAccess
	for _, expression := range statementMemberExpressions(statement, proc.Expressions) {
		receiver, member, memberEnd, ok := splitExcelMemberAccess(expression.Text)
		if !ok {
			continue
		}
		memberLower := strings.ToLower(member)
		if strings.EqualFold(receiver, "Application") && memberLower == "worksheetfunction" {
			out = append(out, excelLoopAccess{Category: excelAccessWorksheetCall, Member: member, Line: line, Read: true})
			continue
		}
		typ, resolved := resolveExcelExpressionType(file, db, receiver, line-1, rootDir, cfg)
		perCell := isPerCellExcelExpression(file, db, receiver, rangeVars, line-1, rootDir, cfg)
		if !resolved && !perCell {
			continue
		}
		if rangeVars.Range[strings.ToLower(strings.TrimSpace(receiver))] {
			typ = "Excel.Range"
		}
		if perCell {
			typ = "Excel.Range"
		}
		assignmentEnd := expressionMemberEnd(statement, expression, memberEnd, len(text))
		if cellValueMembers[memberLower] && isExcelRangeType(typ) {
			access := excelLoopAccess{Category: excelAccessRead, Member: member, Line: line, Read: true}
			if assignmentOperatorAfter(text, assignmentEnd) {
				access.Read = false
				access.Write = true
				access.Category = excelAccessWrite
			}
			if perCell {
				out = append(out, access)
			}
			continue
		}
		if formattingMembers[memberLower] && perCell {
			out = append(out, excelLoopAccess{Category: excelAccessFormatting, Member: member, Line: line, Write: assignmentOperatorAfter(text, assignmentEnd)})
			continue
		}
		if rangeLookupMembers[memberLower] && isExcelWorksheetOrRangeType(typ) {
			if memberLower == "range" && !perCell && literalRangeCellCount(receiver) > 1 {
				continue
			}
			out = append(out, excelLoopAccess{Category: excelAccessRangeLookup, Member: member, Line: line, Read: true})
			continue
		}
		if mutatingRangeMembers[memberLower] && isExcelRangeType(typ) {
			out = append(out, excelLoopAccess{Category: excelAccessWrite, Member: member, Line: line, Write: true})
		}
	}
	if strings.Contains(lower, "worksheetfunction.") || strings.Contains(lower, ".evaluate(") {
		out = append(out, excelLoopAccess{Category: excelAccessWorksheetCall, Line: line, Read: true})
	}
	if excelCellCallRe.MatchString(text) {
		out = append(out, excelLoopAccess{Category: excelAccessRangeLookup, Member: "Cells", Line: line, Read: true})
	}
	directRangeLookup := directExcelLookupRe.MatchString(text)
	if directRangeLookup && literalRangeCellCount(rawText) <= 1 {
		out = append(out, excelLoopAccess{Category: excelAccessRangeLookup, Member: "Range/Worksheet", Line: line, Read: true})
	}
	if statement.Kind == procedureir.StatementForEach && strings.Contains(lower, ".cells") {
		out = append(out, excelLoopAccess{Category: excelAccessRangeLookup, Member: "Cells", Line: line, Read: true})
	}
	return deduplicateExcelAccesses(out)
}

func resolveExcelExpressionType(file parsedFile, db *vbadb.DB, expression string, line int, rootDir string, cfg config.Config) (string, bool) {
	if db == nil || strings.TrimSpace(expression) == "" {
		return "", false
	}
	return (intel.Analyzer{RootDir: rootDir, Config: cfg, DB: db}).ResolveDocumentExpressionTypeAt(file.intelDocument(), expression, line)
}

func isExcelRangeExpression(file parsedFile, db *vbadb.DB, expression string, line int, rootDir string, cfg config.Config) bool {
	typ, ok := resolveExcelExpressionType(file, db, expression, line, rootDir, cfg)
	return ok && isExcelRangeType(typ)
}

func isExcelRangeType(typ string) bool {
	lower := strings.ToLower(strings.TrimSpace(typ))
	return lower == "range" || lower == "excel.range" || strings.HasSuffix(lower, ".range")
}

func isExcelWorksheetOrRangeType(typ string) bool {
	lower := strings.ToLower(strings.TrimSpace(typ))
	return isExcelRangeType(lower) || strings.Contains(lower, "worksheet") || strings.Contains(lower, "worksheets") || strings.Contains(lower, "excel.application")
}

func isPerCellExcelExpression(file parsedFile, db *vbadb.DB, expression string, rangeVars excelRangeVariables, line int, rootDir string, cfg config.Config) bool {
	lower := strings.ToLower(strings.TrimSpace(expression))
	if rangeVars.PerCell[lower] {
		return true
	}
	if strings.HasPrefix(lower, "cells(") {
		return true
	}
	if strings.HasPrefix(lower, "range(") || strings.Contains(lower, ".range(") {
		count := literalRangeCellCount(lower)
		return count == 0 || count == 1
	}
	root := lower
	if dot := strings.IndexByte(root, '.'); dot >= 0 {
		root = strings.TrimSpace(root[:dot])
	}
	if rangeVars.PerCell[root] && containsPerCellFormattingMember(lower) {
		return true
	}
	if !strings.Contains(lower, ".cells(") && !strings.Contains(lower, ".offset(") {
		return false
	}
	if rangeVars.PerCell[root] || rangeVars.Range[root] {
		return true
	}
	typ, ok := resolveExcelExpressionType(file, db, root, line, rootDir, cfg)
	return ok && isExcelWorksheetOrRangeType(typ)
}

func statementMemberExpressions(statement procedureir.Statement, expressions []procedureir.Expression) []procedureir.Expression {
	byID := make(map[int]procedureir.Expression, len(expressions))
	for _, expression := range expressions {
		byID[expression.ID] = expression
	}
	seen := map[int]bool{}
	var out []procedureir.Expression
	var visit func(int)
	visit = func(id int) {
		if seen[id] {
			return
		}
		expression, ok := byID[id]
		if !ok {
			return
		}
		seen[id] = true
		if !expression.Recovered && expression.Kind == procedureir.ExpressionMember {
			out = append(out, expression)
		}
		for _, childID := range expression.Children {
			visit(childID)
		}
	}
	for _, id := range statement.ExpressionIDs {
		visit(id)
	}
	if len(out) == 0 {
		for _, expression := range expressions {
			if expression.StatementID == statement.ID && !expression.Recovered && expression.Kind == procedureir.ExpressionMember {
				out = append(out, expression)
			}
		}
	}
	sort.SliceStable(out, func(i, j int) bool {
		return out[i].Range.StartByte < out[j].Range.StartByte
	})
	return out
}

func splitExcelMemberAccess(text string) (string, string, int, bool) {
	lastDot := -1
	depth := 0
	inString := false
	for i := 0; i < len(text); i++ {
		switch text[i] {
		case '"':
			if inString && i+1 < len(text) && text[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '(':
			if !inString {
				depth++
			}
		case ')':
			if !inString && depth > 0 {
				depth--
			}
		case '.':
			if !inString && depth == 0 {
				lastDot = i
			}
		}
	}
	if lastDot < 0 {
		return "", "", 0, false
	}
	memberStart := lastDot + 1
	for memberStart < len(text) && (text[memberStart] == ' ' || text[memberStart] == '\t') {
		memberStart++
	}
	memberEnd := memberStart
	if memberEnd >= len(text) || !isVBAIdentifierStart(text[memberEnd]) {
		return "", "", 0, false
	}
	for memberEnd < len(text) && isVBAIdentifierPart(text[memberEnd]) {
		memberEnd++
	}
	return strings.TrimSpace(text[:lastDot]), text[memberStart:memberEnd], memberEnd, true
}

func expressionMemberEnd(statement procedureir.Statement, expression procedureir.Expression, memberEnd, textLength int) int {
	offset := expression.Range.StartByte - statement.Range.StartByte
	if offset < 0 || offset+memberEnd > textLength {
		return textLength
	}
	return offset + memberEnd
}

func sanitizeExcelStatementText(text string) string {
	bytes := []byte(text)
	inString := false
	for i := 0; i < len(bytes); i++ {
		if inString {
			bytes[i] = ' '
			if text[i] == '"' {
				if i+1 < len(bytes) && text[i+1] == '"' {
					bytes[i+1] = ' '
					i++
				} else {
					inString = false
				}
			}
			continue
		}
		if text[i] == '"' {
			bytes[i] = ' '
			inString = true
			continue
		}
		if text[i] == '\'' {
			for j := i; j < len(bytes); j++ {
				bytes[j] = ' '
			}
			break
		}
	}
	trimmed := strings.TrimSpace(string(bytes))
	if len(trimmed) >= 4 && strings.EqualFold(trimmed[:4], "rem ") {
		return strings.Repeat(" ", len(text))
	}
	return string(bytes)
}

func containsPerCellFormattingMember(expression string) bool {
	for _, member := range []string{".font", ".interior", ".borders", ".numberformat", ".style", ".alignment"} {
		if strings.Contains(expression, member) {
			return true
		}
	}
	return false
}

func assignmentOperatorAfter(text string, end int) bool {
	if end < 0 || end > len(text) {
		return false
	}
	segmentEnd := len(text)
	if colon := strings.IndexByte(text[end:], ':'); colon >= 0 {
		segmentEnd = end + colon
	}
	lower := strings.ToLower(text)
	then := vbaKeywordPosition(lower, "then", end, segmentEnd)
	conditionalHeader := hasVBAKeywordPrefix(lower, "if") || hasVBAKeywordPrefix(lower, "elseif") || hasVBAKeywordPrefix(lower, "while") || hasVBAKeywordPrefix(lower, "do while") || hasVBAKeywordPrefix(lower, "do until")
	if conditionalHeader && (then < 0 || end < then) {
		return false
	}
	for i := end; i < segmentEnd; i++ {
		if text[i] != '=' {
			continue
		}
		if i > 0 && (text[i-1] == '<' || text[i-1] == '>' || text[i-1] == '=') {
			continue
		}
		if i+1 < len(text) && text[i+1] == '=' {
			continue
		}
		return true
	}
	return false
}

func hasVBAKeywordPrefix(text, keyword string) bool {
	text = strings.TrimSpace(text)
	if !strings.HasPrefix(text, keyword) {
		return false
	}
	return len(text) == len(keyword) || !isVBAIdentifierPart(text[len(keyword)])
}

func vbaKeywordPosition(text, keyword string, start, end int) int {
	for i := start; i+len(keyword) <= end; i++ {
		if !strings.EqualFold(text[i:i+len(keyword)], keyword) {
			continue
		}
		beforeOK := i == 0 || !isVBAIdentifierPart(text[i-1])
		after := i + len(keyword)
		afterOK := after >= len(text) || !isVBAIdentifierPart(text[after])
		if beforeOK && afterOK {
			return i
		}
	}
	return -1
}

func deduplicateExcelAccesses(accesses []excelLoopAccess) []excelLoopAccess {
	seen := map[string]bool{}
	out := make([]excelLoopAccess, 0, len(accesses))
	for _, access := range accesses {
		key := strings.Join([]string{access.Category, strings.ToLower(access.Member), strconv.Itoa(access.Line), strconv.FormatBool(access.Read), strconv.FormatBool(access.Write)}, "\x00")
		if seen[key] {
			continue
		}
		seen[key] = true
		out = append(out, access)
	}
	return out
}

func sortedExcelCategories(categories map[string]bool) []string {
	order := []string{excelAccessRead, excelAccessWrite, excelAccessFormatting, excelAccessRangeLookup, excelAccessWorksheetCall}
	var out []string
	for _, category := range order {
		if categories[category] {
			out = append(out, category)
		}
	}
	return out
}

func sortedStringSet(values map[string]bool) []string {
	out := make([]string, 0, len(values))
	for value := range values {
		out = append(out, value)
	}
	sort.Strings(out)
	return out
}
