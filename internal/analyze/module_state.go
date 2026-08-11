package analyze

import (
	"fmt"
	"path/filepath"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// moduleStateAnalysis is deliberately kept as an analyzer-internal model. The
// public JSON projection is a map so that adding fields to the exploratory
// metrics does not change the Finding contract.
type moduleStateAnalysis struct {
	Findings []Finding
	Metrics  map[string]any
}

type moduleStateProcedure struct {
	Key          string
	File         string
	DisplayFile  string
	Module       string
	ModuleKind   string
	Name         string
	Qualified    string
	Kind         procedureir.ProcedureKind
	ParamCount   int
	Visibility   string
	EventHandler bool
}

type moduleStateProcedureAccess struct {
	Procedure moduleStateProcedure
	Reads     map[string]string
	Writes    map[string]string
	Mutators  map[string]string
}

type moduleStateField struct {
	Key          string
	File         string
	DisplayFile  string
	Module       string
	ModuleKind   string
	Name         string
	Type         string
	Visibility   string
	Kind         string
	Scope        procedureir.SymbolScope
	Line         int
	IsObject     bool
	IsCollection bool
	IsExcel      bool
	Readers      map[string]moduleStateProcedure
	Writers      map[string]moduleStateProcedure
	Mutators     map[string]moduleStateProcedure
	ReadRoots    map[string]bool
	WriteRoots   map[string]bool
	Roots        map[string]moduleStateProcedure
	EventReads   map[string]bool
	EventWrites  map[string]bool
	InCycle      bool
	CycleRead    bool
	CycleWrite   bool
}

func buildModuleStateAnalysis(rootDir string, cfg config.Config, files []parsedFile) moduleStateAnalysis {
	procedures, byKey, byCandidate := moduleStateProcedures(files)
	fields, fieldsByFileName, fieldsByName := moduleStateFields(files)
	procedureAccesses := moduleStateProcedureAccesses(procedures)
	if len(fields) == 0 {
		return moduleStateAnalysis{Metrics: moduleStateMetricsProjection(fields, procedureAccesses)}
	}

	edges := map[string]map[string]bool{}
	for _, procedure := range procedures {
		edges[procedure.Key] = map[string]bool{}
	}
	for _, file := range files {
		for _, proc := range file.IR.Procedures {
			caller := moduleStateProcedureKey(file.IR.Path, file.IR.ModuleName, proc.Symbol)
			if _, ok := byKey[caller]; !ok {
				continue
			}
			for _, call := range proc.Calls {
				if call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
					continue
				}
				callee := moduleStateCandidateKey(call.Resolution.Candidates[0], byCandidate)
				if callee != "" {
					edges[caller][callee] = true
				}
			}
		}
	}

	cycles := moduleStateCycleNodes(edges)
	roots := moduleStateRoots(cfg, procedures)
	eventRoots := moduleStateEventRoots(procedures)
	rootSets := moduleStateRootReachability(edges, roots)

	// Classify collection fields before attributing mutator calls. Otherwise a
	// mutator in a procedure that appears before the initializer is skipped
	// while the field still has its declaration-time type classification.
	for _, file := range files {
		for _, proc := range file.IR.Procedures {
			procedure, ok := byKey[moduleStateProcedureKey(file.IR.Path, file.IR.ModuleName, proc.Symbol)]
			if !ok {
				continue
			}
			for _, access := range proc.Accesses {
				field := moduleStateResolveField(rootDir, access, procedure, fieldsByFileName, fieldsByName)
				if field == nil || field.IsCollection {
					continue
				}
				if moduleStateCollectionInitializer(proc, access, field.Name) {
					field.IsCollection = true
				}
			}
		}
	}

	for _, file := range files {
		for _, proc := range file.IR.Procedures {
			procedure, ok := byKey[moduleStateProcedureKey(file.IR.Path, file.IR.ModuleName, proc.Symbol)]
			if !ok || procedureAccesses[procedure.Key] == nil {
				continue
			}
			for _, access := range proc.Accesses {
				field := moduleStateResolveField(rootDir, access, procedure, fieldsByFileName, fieldsByName)
				if field == nil {
					continue
				}
				if access.Mode == procedureir.AccessRead || access.Mode == procedureir.AccessReadWrite {
					field.Readers[procedure.Key] = procedure
					if cycles[procedure.Key] {
						field.InCycle = true
						field.CycleRead = true
					}
					procedureAccesses[procedure.Key].Reads[field.Key] = moduleStateFieldDisplayName(field)
					for root := range rootSets[procedure.Key] {
						field.ReadRoots[root] = true
						if rootProcedure, ok := byKey[root]; ok {
							field.Roots[root] = rootProcedure
						}
					}
					if procedure.EventHandler {
						field.EventReads[procedure.Key] = true
					}
					for root := range rootSets[procedure.Key] {
						if eventRoots[root] {
							field.EventReads[root] = true
						}
					}
				}
				if access.Mode == procedureir.AccessWrite || access.Mode == procedureir.AccessReadWrite {
					field.Writers[procedure.Key] = procedure
					if cycles[procedure.Key] {
						field.InCycle = true
						field.CycleWrite = true
					}
					procedureAccesses[procedure.Key].Writes[field.Key] = moduleStateFieldDisplayName(field)
					for root := range rootSets[procedure.Key] {
						field.WriteRoots[root] = true
						if rootProcedure, ok := byKey[root]; ok {
							field.Roots[root] = rootProcedure
						}
					}
					if procedure.EventHandler {
						field.EventWrites[procedure.Key] = true
					}
					for root := range rootSets[procedure.Key] {
						if eventRoots[root] {
							field.EventWrites[root] = true
						}
					}
				}
			}
			for _, call := range proc.Calls {
				member := strings.ToLower(strings.TrimSpace(call.Callee.Member))
				if !moduleStateCollectionMutator(proc, call, member) {
					continue
				}
				name := moduleStateReceiverName(call)
				if name == "" {
					continue
				}
				if moduleStateProcedureDeclaresName(proc, name) {
					continue
				}
				field := moduleStateResolveName(name, procedure, fieldsByFileName, fieldsByName)
				if field == nil {
					continue
				}
				if !field.IsCollection {
					continue
				}
				field.Mutators[procedure.Key] = procedure
				field.Writers[procedure.Key] = procedure
				if cycles[procedure.Key] {
					field.InCycle = true
					field.CycleWrite = true
				}
				procedureAccesses[procedure.Key].Mutators[field.Key] = moduleStateFieldDisplayName(field)
				procedureAccesses[procedure.Key].Writes[field.Key] = moduleStateFieldDisplayName(field)
				for root := range rootSets[procedure.Key] {
					field.WriteRoots[root] = true
					if rootProcedure, ok := byKey[root]; ok {
						field.Roots[root] = rootProcedure
					}
				}
				if procedure.EventHandler {
					field.EventWrites[procedure.Key] = true
				}
				for root := range rootSets[procedure.Key] {
					if eventRoots[root] {
						field.EventWrites[root] = true
					}
				}
			}
		}
	}

	metrics := moduleStateMetricsProjection(fields, procedureAccesses)
	findings := moduleStateFindings(rootDir, cfg, files, fields)
	return moduleStateAnalysis{Findings: findings, Metrics: metrics}
}

func moduleStateProcedures(files []parsedFile) ([]moduleStateProcedure, map[string]moduleStateProcedure, map[string]string) {
	all := []moduleStateProcedure{}
	byKey := map[string]moduleStateProcedure{}
	byCandidate := map[string]string{}
	for _, file := range files {
		for _, proc := range file.IR.Procedures {
			item := moduleStateProcedure{
				Key:          moduleStateProcedureKey(file.IR.Path, file.IR.ModuleName, proc.Symbol),
				File:         file.Path,
				DisplayFile:  file.IR.Path,
				Module:       file.IR.ModuleName,
				ModuleKind:   file.IR.ModuleKind,
				Name:         proc.Symbol.Name,
				Qualified:    proc.Symbol.QualifiedName,
				Kind:         proc.Symbol.Kind,
				ParamCount:   len(proc.Symbol.Parameters),
				Visibility:   proc.Symbol.Visibility,
				EventHandler: proc.Symbol.IsEventHandler,
			}
			all = append(all, item)
			byKey[item.Key] = item
			candidate := moduleStateCandidateIdentity(file.IR.Path, proc.Symbol.QualifiedName, proc.Symbol.Kind, proc.Symbol.DeclarationRange.StartLine)
			byCandidate[candidate] = item.Key
		}
	}
	return all, byKey, byCandidate
}

func moduleStateProcedureAccesses(procedures []moduleStateProcedure) map[string]*moduleStateProcedureAccess {
	out := make(map[string]*moduleStateProcedureAccess, len(procedures))
	for _, procedure := range procedures {
		out[procedure.Key] = &moduleStateProcedureAccess{
			Procedure: procedure,
			Reads:     map[string]string{},
			Writes:    map[string]string{},
			Mutators:  map[string]string{},
		}
	}
	return out
}

func moduleStateFields(files []parsedFile) ([]*moduleStateField, map[string]*moduleStateField, map[string][]*moduleStateField) {
	fields := []*moduleStateField{}
	byFileName := map[string]*moduleStateField{}
	byName := map[string][]*moduleStateField{}
	for _, file := range files {
		textDeclarations := moduleDeclarations(file.Lines, sourceProceduresFromIR(file.IR, file.CFG))
		for _, declaration := range file.IR.Declarations {
			if declaration.Scope != procedureir.ScopeModule && declaration.Scope != procedureir.ScopeProject {
				continue
			}
			if declaration.Kind != "variable" && declaration.Kind != "const" {
				continue
			}
			line := declaration.Range.StartLine
			if textDeclaration, ok := textDeclarations[strings.ToLower(strings.TrimSpace(declaration.Name))]; ok && textDeclaration.Line > 0 {
				line = textDeclaration.Line
			}
			field := &moduleStateField{
				Key:          moduleStateFieldKey(file.IR.Path, file.IR.ModuleName, declaration.Name),
				File:         file.Path,
				DisplayFile:  file.IR.Path,
				Module:       file.IR.ModuleName,
				ModuleKind:   file.IR.ModuleKind,
				Name:         declaration.Name,
				Type:         declaration.Type,
				Visibility:   declaration.Visibility,
				Kind:         declaration.Kind,
				Scope:        declaration.Scope,
				Line:         line,
				IsObject:     declaration.IsObject,
				IsCollection: moduleStateCollectionType(declaration.Type),
				IsExcel:      declaration.IsObject && moduleStateExcelType(declaration.Type),
				Readers:      map[string]moduleStateProcedure{},
				Writers:      map[string]moduleStateProcedure{},
				Mutators:     map[string]moduleStateProcedure{},
				ReadRoots:    map[string]bool{},
				WriteRoots:   map[string]bool{},
				Roots:        map[string]moduleStateProcedure{},
				EventReads:   map[string]bool{},
				EventWrites:  map[string]bool{},
			}
			fields = append(fields, field)
			byFileName[moduleStateFieldKey(file.IR.Path, file.IR.ModuleName, declaration.Name)] = field
			byName[strings.ToLower(strings.TrimSpace(declaration.Name))] = append(byName[strings.ToLower(strings.TrimSpace(declaration.Name))], field)
		}
	}
	return fields, byFileName, byName
}

func moduleStateResolveField(rootDir string, access procedureir.VariableAccess, procedure moduleStateProcedure, byFileName map[string]*moduleStateField, byName map[string][]*moduleStateField) *moduleStateField {
	if access.Scope != procedureir.ScopeModule && access.Scope != procedureir.ScopeProject {
		return nil
	}
	if access.Resolution.Scope == procedureir.ScopeProject && len(access.Resolution.Candidates) == 1 {
		candidate := access.Resolution.Candidates[0]
		candidateFile := candidate.File
		if candidateFile != "" && !filepath.IsAbs(candidateFile) && rootDir != "" {
			candidateFile = filepath.Join(rootDir, filepath.FromSlash(candidateFile))
		}
		for _, field := range byName[strings.ToLower(strings.TrimSpace(access.Name))] {
			if strings.EqualFold(moduleStatePathKey(field.File), moduleStatePathKey(candidateFile)) && field.Line == candidate.Line {
				return field
			}
		}
	}
	return moduleStateResolveName(access.Name, procedure, byFileName, byName)
}

func moduleStateResolveName(name string, procedure moduleStateProcedure, byFileName map[string]*moduleStateField, byName map[string][]*moduleStateField) *moduleStateField {
	name = strings.ToLower(strings.TrimSpace(name))
	if name == "" {
		return nil
	}
	if field := byFileName[moduleStateFieldKey(procedure.File, procedure.Module, name)]; field != nil {
		return field
	}
	candidates := byName[name]
	for _, field := range candidates {
		if strings.EqualFold(moduleStatePathKey(field.File), moduleStatePathKey(procedure.File)) && strings.EqualFold(field.Module, procedure.Module) {
			return field
		}
	}
	for _, field := range candidates {
		if field.Scope == procedureir.ScopeProject && strings.EqualFold(field.ModuleKind, "standard") {
			return field
		}
	}
	return nil
}

func moduleStateCandidateKey(candidate procedureir.Candidate, byCandidate map[string]string) string {
	return byCandidate[moduleStateCandidateIdentity(candidate.File, candidate.QualifiedName, procedureir.ProcedureKind(candidate.Kind), candidate.Line)]
}

func moduleStateCandidateIdentity(file, qualified string, kind procedureir.ProcedureKind, line int) string {
	return strings.Join([]string{strings.ToLower(filepath.ToSlash(filepath.Clean(file))), strings.ToLower(strings.TrimSpace(qualified)), strings.ToLower(string(kind)), fmt.Sprint(line)}, "\x00")
}

func moduleStateProcedureKey(file, module string, symbol procedureir.ProcedureSymbol) string {
	return moduleStateCandidateIdentity(file, module+"."+symbol.Name, symbol.Kind, symbol.DeclarationRange.StartLine)
}

func moduleStateFieldKey(file, module, name string) string {
	return strings.ToLower(filepath.ToSlash(filepath.Clean(file)) + "\x00" + strings.TrimSpace(module) + "\x00" + strings.TrimSpace(name))
}

func moduleStateReceiverName(call procedureir.CallSite) string {
	if call.Callee.Receiver != nil {
		return cleanIdentifier(*call.Callee.Receiver)
	}
	text := strings.TrimSpace(call.Callee.Text)
	if dot := strings.LastIndex(text, "."); dot >= 0 {
		return cleanIdentifier(text[:dot])
	}
	return ""
}

func moduleStateCollectionMutator(proc procedureir.ProcedureIR, call procedureir.CallSite, member string) bool {
	switch member {
	case "add", "remove", "removeall", "clear", "delete":
		return true
	case "item", "comparemode":
		return moduleStatePropertySetCall(proc, call)
	default:
		return false
	}
}

func moduleStatePropertySetCall(proc procedureir.ProcedureIR, call procedureir.CallSite) bool {
	for _, statement := range proc.Statements {
		if statement.ID != call.StatementID {
			continue
		}
		if statement.Kind != procedureir.StatementSet && statement.Kind != procedureir.StatementAssignment {
			return false
		}
		return moduleStateExpressionContains(proc.Expressions, statement.TargetID, call.ExpressionID)
	}
	return false
}

func moduleStateExpressionContains(expressions []procedureir.Expression, rootID, childID int) bool {
	if rootID <= 0 || childID <= 0 {
		return false
	}
	for current := childID; current > 0 && current <= len(expressions); {
		if current == rootID {
			return true
		}
		current = expressions[current-1].ParentID
	}
	return false
}

func moduleStateCollectionType(typ string) bool {
	switch moduleStateUnqualifiedType(typ) {
	case "collection", "dictionary":
		return true
	default:
		return false
	}
}

func moduleStateUnqualifiedType(typ string) string {
	lower := strings.ToLower(strings.TrimSpace(typ))
	for _, qualifier := range []string{"excel.", "vba.", "scripting."} {
		if strings.HasPrefix(lower, qualifier) {
			return strings.TrimSpace(strings.TrimPrefix(lower, qualifier))
		}
	}
	return lower
}

func moduleStateCollectionInitializer(proc procedureir.ProcedureIR, access procedureir.VariableAccess, fieldName string) bool {
	if !strings.EqualFold(strings.TrimSpace(access.Name), strings.TrimSpace(fieldName)) {
		return false
	}
	for _, statement := range proc.Statements {
		if access.StatementID > 0 && statement.ID != access.StatementID {
			continue
		}
		text := strings.ToLower(statement.Text)
		if fieldName != "" && !strings.Contains(text, strings.ToLower(strings.TrimSpace(fieldName))) {
			continue
		}
		if strings.Contains(text, "new collection") || strings.Contains(text, "scripting.dictionary") {
			return true
		}
	}
	return false
}

func moduleStateProcedureDeclaresName(proc procedureir.ProcedureIR, name string) bool {
	name = strings.TrimSpace(name)
	if name == "" {
		return false
	}
	for _, parameter := range proc.Symbol.Parameters {
		if strings.EqualFold(strings.TrimSpace(parameter.Name), name) {
			return true
		}
	}
	for _, declaration := range proc.Declarations {
		if !strings.EqualFold(strings.TrimSpace(declaration.Name), name) {
			continue
		}
		if declaration.Scope == procedureir.ScopeLocal || declaration.Scope == procedureir.ScopeParameter {
			return true
		}
	}
	return false
}

func moduleStateExcelType(typ string) bool {
	switch moduleStateUnqualifiedType(typ) {
	case "application", "workbook", "worksheet", "range", "chart", "pivottable", "listobject", "window":
		return true
	default:
		return false
	}
}

func moduleStateRoots(cfg config.Config, procedures []moduleStateProcedure) map[string]bool {
	roots := map[string]bool{}
	entry := strings.ToLower(strings.TrimSpace(cfg.Project.Entry))
	for _, procedure := range procedures {
		qualified := strings.ToLower(strings.TrimSpace(procedure.Qualified))
		public := !strings.EqualFold(strings.TrimSpace(procedure.Visibility), "Private")
		if entry != "" && qualified == entry {
			roots[procedure.Key] = true
		}
		if strings.EqualFold(procedure.ModuleKind, "standard") && public && procedure.Kind == procedureir.ProcedureSub && procedure.ParamCount == 0 {
			roots[procedure.Key] = true
		}
		if procedure.EventHandler {
			roots[procedure.Key] = true
		}
	}
	return roots
}

func moduleStateEventRoots(procedures []moduleStateProcedure) map[string]bool {
	roots := map[string]bool{}
	for _, procedure := range procedures {
		if procedure.EventHandler {
			roots[procedure.Key] = true
		}
	}
	return roots
}

func moduleStateRootReachability(edges map[string]map[string]bool, roots map[string]bool) map[string]map[string]bool {
	sets := map[string]map[string]bool{}
	for node := range edges {
		sets[node] = map[string]bool{}
	}
	queue := []string{}
	for root := range roots {
		sets[root][root] = true
		queue = append(queue, root)
	}
	sort.Strings(queue)
	for len(queue) > 0 {
		node := queue[0]
		queue = queue[1:]
		for next := range edges[node] {
			changed := false
			for root := range sets[node] {
				if !sets[next][root] {
					sets[next][root] = true
					changed = true
				}
			}
			if changed {
				queue = append(queue, next)
			}
		}
	}
	return sets
}

func moduleStateCycleNodes(edges map[string]map[string]bool) map[string]bool {
	// A node is in a cycle when it can reach itself through a non-empty edge.
	// The project graphs are small compared with source parsing, and this
	// deterministic DFS avoids exposing a second graph abstraction publicly.
	cycles := map[string]bool{}
	for start := range edges {
		seen := map[string]bool{}
		stack := []string{start}
		for len(stack) > 0 {
			node := stack[len(stack)-1]
			stack = stack[:len(stack)-1]
			for next := range edges[node] {
				if next == start {
					cycles[start] = true
					continue
				}
				if !seen[next] {
					seen[next] = true
					stack = append(stack, next)
				}
			}
		}
	}
	return cycles
}

func moduleStateMetricsProjection(fields []*moduleStateField, accesses map[string]*moduleStateProcedureAccess) map[string]any {
	items := make([]map[string]any, 0, len(fields))
	for _, field := range fields {
		classification := "mutable"
		if field.Kind == "const" {
			classification = "constant"
		} else if !field.IsCollection && len(field.Writers) == 0 && len(field.Mutators) == 0 {
			classification = "read_only_configuration"
		}
		item := map[string]any{
			"file": moduleStateMetricsFile(field.DisplayFile, field.File), "module": field.Module, "module_kind": field.ModuleKind,
			"name": field.Name, "type": field.Type, "visibility": field.Visibility,
			"kind": field.Kind, "scope": string(field.Scope), "line": field.Line,
			"classification": classification, "reader_count": len(field.Readers),
			"writer_count": len(field.Writers), "mutator_count": len(field.Mutators),
			"root_count":         len(moduleStateUnion(field.ReadRoots, field.WriteRoots)),
			"roots":              moduleStateProcedureNames(field.Roots),
			"event_reader_count": len(field.EventReads), "event_writer_count": len(field.EventWrites),
			"in_call_cycle": field.InCycle, "cached_excel_reference": field.IsExcel,
			"collection_or_dictionary": field.IsCollection,
			"readers":                  moduleStateProcedureNames(field.Readers), "writers": moduleStateProcedureNames(field.Writers),
		}
		items = append(items, item)
	}
	sort.SliceStable(items, func(i, j int) bool {
		leftFile, rightFile := fmt.Sprint(items[i]["file"]), fmt.Sprint(items[j]["file"])
		if leftFile != rightFile {
			return leftFile < rightFile
		}
		leftLine, rightLine := items[i]["line"].(int), items[j]["line"].(int)
		if leftLine != rightLine {
			return leftLine < rightLine
		}
		return fmt.Sprint(items[i]["name"]) < fmt.Sprint(items[j]["name"])
	})
	state := map[string]any{"fields": items, "field_count": len(items)}
	state["procedures"] = moduleStateProcedureMetrics(accesses)
	return map[string]any{"module_state": state}
}

func moduleStateProcedureMetrics(accesses map[string]*moduleStateProcedureAccess) []map[string]any {
	keys := make([]string, 0, len(accesses))
	for key := range accesses {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	items := make([]map[string]any, 0, len(accesses))
	for _, key := range keys {
		access := accesses[key]
		if access == nil {
			continue
		}
		items = append(items, map[string]any{
			"file": moduleStateMetricsFile(access.Procedure.DisplayFile, access.Procedure.File), "module": access.Procedure.Module,
			"module_kind": access.Procedure.ModuleKind, "name": access.Procedure.Name,
			"qualified": access.Procedure.Qualified, "reads": moduleStateFieldNames(access.Reads),
			"writes": moduleStateFieldNames(access.Writes), "mutators": moduleStateFieldNames(access.Mutators),
		})
	}
	sort.SliceStable(items, func(i, j int) bool {
		left := fmt.Sprintf("%s\x00%s\x00%s", items[i]["file"], items[i]["module"], items[i]["name"])
		right := fmt.Sprintf("%s\x00%s\x00%s", items[j]["file"], items[j]["module"], items[j]["name"])
		return left < right
	})
	return items
}

func moduleStateFieldNames(fields map[string]string) []string {
	seen := make(map[string]bool, len(fields))
	out := make([]string, 0, len(fields))
	for _, field := range fields {
		if field == "" || seen[field] {
			continue
		}
		seen[field] = true
		out = append(out, field)
	}
	sort.Strings(out)
	return out
}

func moduleStateFieldDisplayName(field *moduleStateField) string {
	if field == nil {
		return ""
	}
	return strings.Trim(strings.Join([]string{field.Module, field.Name}, "."), ".")
}

func moduleStateMetricsFile(display, fallback string) string {
	if strings.TrimSpace(display) != "" {
		return filepath.ToSlash(filepath.Clean(display))
	}
	return filepath.ToSlash(filepath.Clean(fallback))
}

func moduleStateProcedureNames(items map[string]moduleStateProcedure) []string {
	out := make([]string, 0, len(items))
	for _, item := range items {
		out = append(out, item.Qualified)
	}
	sort.Strings(out)
	return out
}

func moduleStateUnion(left, right map[string]bool) map[string]bool {
	out := map[string]bool{}
	for key := range left {
		out[key] = true
	}
	for key := range right {
		out[key] = true
	}
	return out
}

func moduleStateFindings(rootDir string, cfg config.Config, files []parsedFile, fields []*moduleStateField) []Finding {
	if !cfg.Analyze.DetectRiskyModuleState {
		return nil
	}
	byFile := map[string]parsedFile{}
	for _, file := range files {
		byFile[moduleStatePathKey(file.Path)] = file
	}
	findings := []Finding{}
	for _, field := range fields {
		if field.Kind == "const" {
			continue
		}
		reasons := []string{}
		roots := moduleStateUnion(field.ReadRoots, field.WriteRoots)
		if len(field.WriteRoots) > 0 && len(roots) > 1 {
			reasons = append(reasons, "multiple entry points share mutable state")
		}
		if len(field.EventReads) > 0 && len(field.EventWrites) > 0 {
			reasons = append(reasons, "state is retained across event invocations")
		}
		if field.CycleRead && field.CycleWrite {
			reasons = append(reasons, "mutable state participates in a call cycle")
		}
		if field.IsExcel && len(reasons) > 0 {
			reasons = append(reasons, "cached Excel object reference may outlive its workbook context")
		}
		if field.IsCollection && len(field.WriteRoots) > 0 && len(roots) > 1 {
			reasons = append(reasons, "Collection or Dictionary mutation crosses entry points")
		}
		if len(reasons) == 0 {
			continue
		}
		file := byFile[moduleStatePathKey(field.File)]
		line := field.Line
		rel, err := filepath.Rel(rootDir, field.File)
		if err != nil {
			rel = field.File
		}
		message := fmt.Sprintf("Module-level mutable field %q has lifecycle coupling (%d readers, %d writers, %d roots).", field.Name, len(field.Readers), len(field.Writers), len(roots))
		reason := strings.Join(reasons, "; ") + "."
		suggestion := "Prefer explicit parameters or return values; otherwise centralize initialization and mutation at one lifecycle boundary."
		if field.IsExcel {
			suggestion = "Reacquire the Excel object when needed, validate it is not Nothing and still belongs to the expected workbook, or pass it explicitly."
		} else if field.IsCollection {
			suggestion = "Keep Collection or Dictionary ownership in one procedure/module and expose operations through explicit parameters or return values."
		}
		finding := Finding{
			Code: "VBA240", Severity: "warning", File: filepath.ToSlash(rel), Module: file.Module,
			Line: line, Message: message, Reason: reason, Suggestion: suggestion,
			NearbyCode: nearby(file.Lines, line, 2),
		}
		findings = append(findings, finding)
	}
	return findings
}

func moduleStatePathKey(path string) string {
	return strings.ToLower(filepath.ToSlash(filepath.Clean(path)))
}
