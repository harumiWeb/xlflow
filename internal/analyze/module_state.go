package analyze

import (
	"fmt"
	"path/filepath"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/modulestate"
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
	Callees      []string
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
	documents := make([]procedureir.DocumentIR, len(files))
	for index := range files {
		documents[index] = files[index].IR
	}
	facts := modulestate.CollectWithOptions(documents, modulestate.Options{RootDir: rootDir})
	procedures, byKey := moduleStateProcedures(files, facts.Procedures)
	fields, fieldsByID := moduleStateFieldsFromFacts(files, facts.Fields, byKey)
	procedureAccesses := moduleStateProcedureAccesses(procedures)
	moduleStatePopulateProcedureAccesses(facts.Procedures, fieldsByID, procedureAccesses)
	if len(fields) == 0 {
		return moduleStateAnalysis{Metrics: moduleStateMetricsProjection(fields, procedureAccesses)}
	}

	edges := map[string]map[string]bool{}
	for _, procedure := range procedures {
		edges[procedure.Key] = map[string]bool{}
		for _, callee := range procedure.Callees {
			if _, ok := byKey[callee]; ok {
				edges[procedure.Key][callee] = true
			}
		}
	}

	cycles := moduleStateCycleNodes(edges)
	roots := moduleStateRoots(cfg, procedures)
	eventRoots := moduleStateEventRoots(procedures)
	rootSets := moduleStateRootReachability(edges, roots)

	for _, field := range fields {
		for key, procedure := range field.Readers {
			moduleStateAttributeAccess(field, procedure, key, true, cycles, rootSets, eventRoots, byKey)
		}
		for key, procedure := range field.Writers {
			moduleStateAttributeAccess(field, procedure, key, false, cycles, rootSets, eventRoots, byKey)
		}
	}

	metrics := moduleStateMetricsProjection(fields, procedureAccesses)
	findings := moduleStateFindings(rootDir, cfg, files, fields)
	return moduleStateAnalysis{Findings: findings, Metrics: metrics}
}

func moduleStateProcedures(files []parsedFile, facts []modulestate.Procedure) ([]moduleStateProcedure, map[string]moduleStateProcedure) {
	all := []moduleStateProcedure{}
	byKey := map[string]moduleStateProcedure{}
	byFile := make(map[string]parsedFile, len(files))
	for _, file := range files {
		byFile[moduleStatePathKey(file.IR.Path)] = file
	}
	for _, fact := range facts {
		file := byFile[moduleStatePathKey(fact.Path)]
		item := moduleStateProcedure{
			Key: fact.ID, File: file.Path, DisplayFile: fact.Path,
			Module: fact.Module, ModuleKind: fact.ModuleKind, Name: fact.Name,
			Qualified: fact.QualifiedName, Kind: fact.Kind,
			ParamCount: fact.ParameterCount, Visibility: fact.Visibility,
			EventHandler: fact.EventHandler, Callees: fact.Callees,
		}
		if item.File == "" {
			item.File = fact.Path
		}
		all = append(all, item)
		byKey[item.Key] = item
	}
	return all, byKey
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

func moduleStatePopulateProcedureAccesses(facts []modulestate.Procedure, fields map[string]*moduleStateField, accesses map[string]*moduleStateProcedureAccess) {
	for _, fact := range facts {
		access := accesses[fact.ID]
		if access == nil {
			continue
		}
		for _, key := range fact.Reads {
			if field := fields[key]; field != nil {
				access.Reads[key] = moduleStateFieldDisplayName(field)
			}
		}
		for _, key := range fact.Writes {
			if field := fields[key]; field != nil {
				access.Writes[key] = moduleStateFieldDisplayName(field)
			}
		}
		for _, key := range fact.Mutators {
			if field := fields[key]; field != nil {
				access.Mutators[key] = moduleStateFieldDisplayName(field)
			}
		}
	}
}

func moduleStateFieldsFromFacts(files []parsedFile, facts []modulestate.Field, procedures map[string]moduleStateProcedure) ([]*moduleStateField, map[string]*moduleStateField) {
	fields := []*moduleStateField{}
	byID := map[string]*moduleStateField{}
	byFile := make(map[string]parsedFile, len(files))
	declarationsByFile := make(map[string]map[string]sourceDeclaration, len(files))
	for _, file := range files {
		path := moduleStatePathKey(file.IR.Path)
		byFile[path] = file
		declarationsByFile[path] = file.moduleDecls()
	}
	for _, fact := range facts {
		path := moduleStatePathKey(fact.Path)
		file := byFile[path]
		line := fact.DeclarationLine
		if file.Path != "" {
			if textDeclaration, ok := declarationsByFile[path][strings.ToLower(strings.TrimSpace(fact.Name))]; ok && textDeclaration.Line > 0 {
				line = textDeclaration.Line
			}
		}
		field := &moduleStateField{
			Key: fact.ID, File: file.Path, DisplayFile: fact.Path,
			Module: fact.Module, ModuleKind: fact.ModuleKind, Name: fact.Name,
			Type: fact.Type, Visibility: fact.Visibility, Kind: fact.Kind,
			Scope: fact.Scope, Line: line, IsObject: fact.IsObject,
			IsCollection: fact.IsCollection, IsExcel: fact.IsExcel,
			Readers: map[string]moduleStateProcedure{}, Writers: map[string]moduleStateProcedure{},
			Mutators: map[string]moduleStateProcedure{}, ReadRoots: map[string]bool{},
			WriteRoots: map[string]bool{}, Roots: map[string]moduleStateProcedure{},
			EventReads: map[string]bool{}, EventWrites: map[string]bool{},
		}
		if field.File == "" {
			field.File = fact.Path
		}
		for _, key := range fact.Readers {
			if procedure, ok := procedures[key]; ok {
				field.Readers[key] = procedure
			}
		}
		for _, key := range fact.Writers {
			if procedure, ok := procedures[key]; ok {
				field.Writers[key] = procedure
			}
		}
		for _, key := range fact.Mutators {
			if procedure, ok := procedures[key]; ok {
				field.Mutators[key] = procedure
			}
		}
		fields = append(fields, field)
		byID[field.Key] = field
	}
	return fields, byID
}

func moduleStateAttributeAccess(field *moduleStateField, procedure moduleStateProcedure, key string, read bool, cycles map[string]bool, rootSets map[string]map[string]bool, eventRoots map[string]bool, byKey map[string]moduleStateProcedure) {
	if cycles[key] {
		field.InCycle = true
		if read {
			field.CycleRead = true
		} else {
			field.CycleWrite = true
		}
	}
	for root := range rootSets[key] {
		if read {
			field.ReadRoots[root] = true
		} else {
			field.WriteRoots[root] = true
		}
		if rootProcedure, ok := byKey[root]; ok {
			field.Roots[root] = rootProcedure
		}
	}
	if procedure.EventHandler {
		if read {
			field.EventReads[key] = true
		} else {
			field.EventWrites[key] = true
		}
	}
	for root := range rootSets[key] {
		if !eventRoots[root] {
			continue
		}
		if read {
			field.EventReads[root] = true
		} else {
			field.EventWrites[root] = true
		}
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
