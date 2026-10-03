package architecture

import (
	"cmp"
	"context"
	"fmt"
	"os"
	"path/filepath"
	"slices"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/callgraph"
	"github.com/harumiWeb/xlflow/internal/vba/calls"
	"github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/effects"
	"github.com/harumiWeb/xlflow/internal/vba/hotspots"
	"github.com/harumiWeb/xlflow/internal/vba/modulestate"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/proceduremetrics"
	"github.com/harumiWeb/xlflow/internal/vba/reachability"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

type collectionHooks struct {
	afterSourceRead func(string)
	afterParse      func(string)
	afterIRBuild    func(string)
	afterCFG        func(string)
	loadTypeDB      func() (typedb.LoadResult, error)
}

type sourceSnapshot struct {
	path       string
	physical   string
	moduleKind string
	symbols    symbols.FileResult
	ir         procedureir.DocumentIR
	cfg        cfg.Document
	resolved   procedureir.DocumentIR
}

func collectContext(ctx context.Context, root string, cfgValue config.Config, options Options) (Report, []map[string]any, error) {
	return collectContextWithHooks(ctx, root, cfgValue, options, nil)
}

func collectContextWithHooks(ctx context.Context, root string, cfgValue config.Config, options Options, hooks *collectionHooks) (Report, []map[string]any, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return Report{}, []map[string]any{}, err
	}
	options.Module = strings.TrimSpace(options.Module)
	if hooks == nil {
		hooks = &collectionHooks{}
	}
	loadTypeDB := hooks.loadTypeDB
	if loadTypeDB == nil {
		loadTypeDB = func() (typedb.LoadResult, error) { return typedb.LoadForRuntime("") }
	}
	rootAbs, err := filepath.Abs(root)
	if err != nil {
		return Report{}, []map[string]any{}, err
	}
	scopePath, scopePhysical, err := validatePathScope(rootAbs, options.Path)
	if err != nil {
		return Report{}, []map[string]any{}, err
	}
	options.Path = scopePath
	files, err := symbols.DiscoverProjectSourceFilesContext(ctx, rootAbs, cfgValue)
	if err != nil {
		return Report{}, []map[string]any{}, err
	}
	if err := ctx.Err(); err != nil {
		return Report{}, []map[string]any{}, err
	}
	if options.Path != "" && !discoveryContainsPath(files, scopePhysical) {
		info, err := os.Stat(scopePhysical)
		if err != nil {
			return Report{}, []map[string]any{}, fmt.Errorf("inspect architecture scope %q: %w", options.Path, err)
		}
		if !info.IsDir() {
			return Report{}, []map[string]any{}, fmt.Errorf("%w: path %q is not a discovered VBA source", ErrInvalidScope, options.Path)
		}
	}

	snapshots := make([]sourceSnapshot, 0, len(files))
	capturedSources := make(map[string][]byte, len(files)+1)
	allSymbols := make([]symbols.Symbol, 0)
	moduleKinds := make(map[string]string, len(files))
	for _, sourceFile := range files {
		if err := ctx.Err(); err != nil {
			return Report{}, []map[string]any{}, err
		}
		path, err := filepath.Rel(rootAbs, sourceFile.Path)
		if err != nil {
			return Report{}, []map[string]any{}, err
		}
		path = filepath.ToSlash(filepath.Clean(path))
		body, err := os.ReadFile(sourceFile.Path)
		if err != nil {
			return Report{}, []map[string]any{}, fmt.Errorf("read VBA source %q: %w", path, err)
		}
		if hooks.afterSourceRead != nil {
			hooks.afterSourceRead(path)
		}
		if err := ctx.Err(); err != nil {
			return Report{}, []map[string]any{}, err
		}
		capturedSources[path] = body
		parsed, err := ast.ParseDocumentContext(ctx, path, body)
		if err != nil {
			if ctx.Err() != nil {
				return Report{}, []map[string]any{}, ctx.Err()
			}
			return Report{}, []map[string]any{}, fmt.Errorf("%w: %s: %w", ErrParse, path, err)
		}
		if hooks.afterParse != nil {
			hooks.afterParse(path)
		}
		symbolFile, err := symbols.InspectParsedContext(ctx, symbols.SourceOptions{
			RootDir: rootAbs, Path: path, ModuleKind: sourceFile.ModuleKind, IncludePrivate: true,
		}, parsed)
		if err != nil {
			parsed.Close()
			return Report{}, []map[string]any{}, err
		}
		ir, err := procedureir.BuildParsedContext(ctx, procedureir.BuildOptions{
			RootDir: rootAbs, Path: path, ModuleName: symbolFile.ModuleName, ModuleKind: sourceFile.ModuleKind,
		}, parsed)
		if err != nil {
			parsed.Close()
			return Report{}, []map[string]any{}, err
		}
		if ir.Parse.HasError || ir.Parse.HasMissing {
			parsed.Close()
			if ctx.Err() != nil {
				return Report{}, []map[string]any{}, ctx.Err()
			}
			return Report{}, []map[string]any{}, fmt.Errorf("%w: %s contains parser errors or missing syntax", ErrParse, path)
		}
		parsed.Close()
		if hooks.afterIRBuild != nil {
			hooks.afterIRBuild(path)
		}
		graph, err := cfg.BuildDocumentContext(ctx, ir)
		if err != nil {
			return Report{}, []map[string]any{}, err
		}
		if hooks.afterCFG != nil {
			hooks.afterCFG(path)
		}
		allSymbols = append(allSymbols, symbolFile.Symbols...)
		moduleKinds[symbolFile.Path] = symbolFile.ModuleKind
		snapshots = append(snapshots, sourceSnapshot{
			path: path, physical: sourceFile.Path, moduleKind: sourceFile.ModuleKind,
			symbols: symbolFile, ir: ir, cfg: graph,
		})
		if err := ctx.Err(); err != nil {
			return Report{}, []map[string]any{}, err
		}
	}
	if module := strings.TrimSpace(options.Module); module != "" {
		found := false
		for _, snapshot := range snapshots {
			if strings.EqualFold(snapshot.ir.ModuleName, module) {
				found = true
				break
			}
		}
		if !found {
			return Report{}, []map[string]any{}, fmt.Errorf("%w: module %q does not exist", ErrInvalidScope, module)
		}
	}
	typeDBResult, typeDBErr := loadTypeDB()
	warnings := make([]map[string]any, 0, len(typeDBResult.Warnings)+1)
	typeDBWarnings := slices.Clone(typeDBResult.Warnings)
	if typeDBErr != nil {
		typeDBResult.Complete = false
		typeDBWarnings = append(typeDBWarnings, typeDBErr.Error())
	}
	if typeDBResult.DB == nil {
		typeDBResult.Complete = false
		fallbackDB, fallbackErr := vbadb.LoadBuiltin()
		if fallbackErr == nil {
			typeDBResult.DB = fallbackDB
		} else {
			typeDBResult.Complete = false
			typeDBWarnings = append(typeDBWarnings, fallbackErr.Error())
		}
	}
	for _, warning := range typeDBWarnings {
		warnings = append(warnings, map[string]any{"code": "type_database_incomplete", "message": warning})
	}
	externalSymbols := typeDBResolverSymbols(typeDBResult.DB)

	resolverDocuments := make([]procedureir.DocumentIR, len(snapshots))
	for index := range snapshots {
		resolverDocuments[index] = snapshots[index].ir
	}
	canonicalResolver := procedureir.BuildProjectResolver(resolverDocuments, externalSymbols, typeDBResult.Complete)

	callResult := calls.Result{
		Root: filepath.ToSlash(rootAbs), Calls: []calls.Call{}, Symbols: allSymbols,
		ModuleKinds: moduleKinds, TypeReferences: []calls.TypeReference{}, DynamicReferences: []calls.DynamicReference{},
	}
	resolvedDocuments := make([]procedureir.DocumentIR, 0, len(snapshots))
	effectDocuments := make([]effects.Document, 0, len(snapshots))
	procedureMetrics := make([]proceduremetrics.ProcedureMetrics, 0)
	for index := range snapshots {
		if err := ctx.Err(); err != nil {
			return Report{}, []map[string]any{}, err
		}
		snapshot := &snapshots[index]
		view := procedureir.ResolveView(snapshot.ir, canonicalResolver)
		snapshot.resolved = view.Materialize()
		resolvedDocuments = append(resolvedDocuments, snapshot.resolved)
		effectDocuments = append(effectDocuments, effects.Document{IR: snapshot.ir, Resolution: view, CFG: snapshot.cfg})
		procedureMetrics = append(procedureMetrics, proceduremetrics.CollectDocument(snapshot.resolved, snapshot.cfg)...)
		callResult.TypeReferences = append(callResult.TypeReferences, calls.TypeReferencesFromIR(snapshot.ir)...)
		callResult.Calls = append(callResult.Calls, calls.FromResolvedIR(snapshot.resolved)...)
		// Implicit Application APIs need resolution to distinguish callbacks
		// from project procedures and non-callable locals with the same name.
		for _, procedure := range snapshot.resolved.Procedures {
			for _, site := range procedure.Calls {
				callResult.DynamicReferences = append(callResult.DynamicReferences, calls.DynamicReferencesForIR(site, procedure.Expressions, symbols.ParseSummary{
					HasError: snapshot.ir.Parse.HasError, HasMissing: snapshot.ir.Parse.HasMissing,
				})...)
			}
		}
	}
	callResult.Summary.Files = len(snapshots)
	callResult.Summary.Calls = len(callResult.Calls)
	for _, call := range callResult.Calls {
		switch call.Resolution.Status {
		case "matched":
			callResult.Summary.Matched++
		case "ambiguous":
			callResult.Summary.Ambiguous++
		case "unresolved", "external", "member_call", "dynamic", "incomplete", "non_callable":
			callResult.Summary.Unresolved++
		}
	}

	graphSnapshot := callgraph.SnapshotFromResult(&callResult)
	dependencyResult := callgraph.Dependencies(graphSnapshot, callgraph.DependencyRequest{})
	cyclicComponents, err := callgraph.FindCyclicComponentsContext(ctx, graphSnapshot)
	if err != nil {
		return Report{}, []map[string]any{}, err
	}
	reachabilityResult, err := reachability.AnalyzeContext(ctx, reachability.Options{
		RootDir: rootAbs, Config: cfgValue,
		Symbols: &symbols.Result{Root: filepath.ToSlash(rootAbs), Files: symbolFiles(snapshots)},
		Calls:   &callResult, Sources: capturedSources,
	})
	if err != nil {
		return Report{}, []map[string]any{}, err
	}
	stateFacts, err := modulestate.CollectContext(ctx, resolvedDocuments, modulestate.Options{RootDir: rootAbs})
	if err != nil {
		return Report{}, []map[string]any{}, err
	}
	if err := ctx.Err(); err != nil {
		return Report{}, []map[string]any{}, err
	}
	effectSummary, _ := effects.BuildWithStats(effectDocuments)
	hotspotReport, _, err := hotspots.BuildFromMetrics(ctx, procedureMetrics)
	if err != nil {
		return Report{}, []map[string]any{}, err
	}
	if err := ctx.Err(); err != nil {
		return Report{}, []map[string]any{}, err
	}

	full := assembleFullReport(options, snapshots, callResult, dependencyResult, cyclicComponents, reachabilityResult, stateFacts, effectSummary, hotspotReport, procedureMetrics, typeDBResult.DB != nil, typeDBResult.Complete, typeDBWarnings)
	full.ProjectSummary = summarize(full, snapshots, callResult, reachabilityResult)
	full.Summary = full.ProjectSummary
	if options.Path == "" && strings.TrimSpace(options.Module) == "" {
		return full, warnings, nil
	}
	filtered, ok := scopeReport(full, snapshots, callResult, reachabilityResult, options, scopePhysical)
	if !ok {
		return Report{}, []map[string]any{}, fmt.Errorf("%w: no source matches the requested path and module", ErrInvalidScope)
	}
	return filtered, warnings, nil
}

func symbolFiles(snapshots []sourceSnapshot) []symbols.FileResult {
	out := make([]symbols.FileResult, 0, len(snapshots))
	for _, snapshot := range snapshots {
		out = append(out, snapshot.symbols)
	}
	return out
}

func validatePathScope(root, requested string) (string, string, error) {
	requested = strings.TrimSpace(requested)
	if requested == "" {
		return "", "", nil
	}
	physical := requested
	if !filepath.IsAbs(physical) {
		physical = filepath.Join(root, filepath.FromSlash(physical))
	}
	physical, err := filepath.Abs(physical)
	if err != nil {
		return "", "", fmt.Errorf("%w: %v", ErrInvalidScope, err)
	}
	physical = filepath.Clean(physical)
	relative, err := filepath.Rel(root, physical)
	if err != nil || relative == ".." || strings.HasPrefix(relative, ".."+string(filepath.Separator)) {
		return "", "", fmt.Errorf("%w: path must be inside the project", ErrInvalidScope)
	}
	if _, err := os.Stat(physical); err != nil {
		if !os.IsNotExist(err) {
			return "", "", fmt.Errorf("inspect architecture scope %q: %w", requested, err)
		}
		return "", "", fmt.Errorf("%w: path %q: %v", ErrInvalidScope, requested, err)
	}
	return filepath.ToSlash(filepath.Clean(relative)), physical, nil
}

func discoveryContainsPath(files []symbols.SourceFile, requested string) bool {
	for _, file := range files {
		if samePhysicalPath(file.Path, requested) || pathWithin(file.Path, requested) {
			return true
		}
	}
	return false
}

func samePhysicalPath(left, right string) bool {
	return strings.EqualFold(filepath.Clean(left), filepath.Clean(right))
}

func pathWithin(file, directory string) bool {
	rel, err := filepath.Rel(directory, file)
	return err == nil && rel != ".." && !strings.HasPrefix(rel, ".."+string(filepath.Separator))
}

func assembleFullReport(
	options Options, snapshots []sourceSnapshot, callResult calls.Result,
	dependencies callgraph.DependencyResult, cycles []callgraph.CyclicComponent, reachable reachability.Result,
	state modulestate.Facts, effectSummary effects.ProjectSummary, hotspotReport hotspots.Report,
	procedureMetrics []proceduremetrics.ProcedureMetrics, typeDBLoaded, typeDBComplete bool, typeDBWarnings []string,
) Report {
	dependencyModules := make(map[string]string)
	dependencyProcedures := make(map[string]string)
	for _, node := range dependencies.Nodes {
		if node.Kind == "module" {
			dependencyModules[moduleKey(node.Module, node.File)] = node.ID
		}
		if strings.HasPrefix(node.ID, "procedure|") {
			dependencyProcedures[strings.TrimPrefix(node.ID, "procedure|")] = node.ID
		}
	}
	procedureIDs := make(map[string]string)
	semanticProcedureIDs := make(map[string]string)
	callerIDs := make(map[string]string)
	metricsByProcedure := make(map[string]proceduremetrics.Metrics, len(procedureMetrics))
	for _, metric := range procedureMetrics {
		metricsByProcedure[metricIdentityKey(metric.File, metric.Module, metric.Name, string(metric.Kind), metric.DeclarationRange.StartByte)] = metric.Metrics
	}
	reachabilityByProcedure := make(map[string]string, len(reachable.Confirmed)+len(reachable.Possible)+len(reachable.Unreachable))
	for _, node := range reachable.Confirmed {
		reachabilityByProcedure[node.ID.String()] = "confirmed"
	}
	for _, node := range reachable.Possible {
		reachabilityByProcedure[node.ID.String()] = "possible"
	}
	for _, node := range reachable.Unreachable {
		reachabilityByProcedure[node.ID.String()] = "unreachable"
	}
	procedures := make([]Procedure, 0)
	modules := make([]Module, 0, len(snapshots))
	for _, snapshot := range snapshots {
		moduleID := dependencyModules[moduleKey(snapshot.ir.ModuleName, snapshot.path)]
		modules = append(modules, Module{ID: moduleID, Name: snapshot.ir.ModuleName, Kind: snapshot.moduleKind, File: snapshot.path, ProcedureCount: len(snapshot.ir.Procedures)})
		for _, proc := range snapshot.ir.Procedures {
			rng := proc.Symbol.DeclarationRange
			id := callgraph.ID{
				Module: snapshot.ir.ModuleName, QualifiedName: proc.Symbol.QualifiedName,
				Kind: string(proc.Symbol.Kind), File: snapshot.path, Line: rng.StartLine, Column: rng.StartColumn,
			}.String()
			procedureDependencyID := dependencyProcedures[id]
			if procedureDependencyID == "" {
				procedureDependencyID = "procedure|" + id
			}
			procedureIDs[effectsIdentityKey(effects.ProcedureIdentity{
				File: snapshot.path, Module: snapshot.ir.ModuleName, Name: proc.Symbol.Name,
				QualifiedName: proc.Symbol.QualifiedName, Kind: proc.Symbol.Kind,
				DeclarationLine: rng.StartLine,
			})] = procedureDependencyID
			semanticProcedureIDs[semanticProcedureKey(snapshot.path, snapshot.ir.ModuleName, proc.Symbol.Name, string(proc.Symbol.Kind), rng.StartLine)] = procedureDependencyID
			callerKey := callerIdentityKey(snapshot.path, snapshot.ir.ModuleName, proc.Symbol.QualifiedName, string(proc.Symbol.Kind))
			if previous, exists := callerIDs[callerKey]; exists && previous != procedureDependencyID {
				callerIDs[callerKey] = ""
			} else {
				callerIDs[callerKey] = procedureDependencyID
			}
			procedures = append(procedures, Procedure{
				ID: procedureDependencyID, CallgraphID: id, ModuleID: moduleID, Name: proc.Symbol.Name, QualifiedName: proc.Symbol.QualifiedName,
				Kind: string(proc.Symbol.Kind), Module: snapshot.ir.ModuleName, ModuleKind: snapshot.moduleKind,
				File: snapshot.path, Location: callgraph.Location{File: snapshot.path, StartLine: rng.StartLine, StartColumn: rng.StartColumn, EndLine: rng.EndLine, EndColumn: rng.EndColumn},
				Visibility: proc.Symbol.Visibility, Reachability: reachabilityByProcedure[id],
				DeclarationRange: sourceRange(rng),
				Metrics:          metricsByProcedure[metricIdentityKey(snapshot.path, snapshot.ir.ModuleName, proc.Symbol.Name, string(proc.Symbol.Kind), rng.StartByte)],
			})
		}
	}
	stateProcedureIDs := make(map[string]string, len(state.Procedures))
	for _, procedure := range state.Procedures {
		key := semanticProcedureKey(procedure.Path, procedure.Module, procedure.Name, string(procedure.Kind), procedure.DeclarationLine)
		stateProcedureIDs[procedure.ID] = semanticProcedureIDs[key]
	}
	modules = slices.Clip(modules)
	slices.SortFunc(modules, func(a, b Module) int { return cmp.Compare(a.ID, b.ID) })
	slices.SortFunc(procedures, func(a, b Procedure) int { return cmp.Compare(a.ID, b.ID) })
	entryPoints := make([]EntryPoint, 0, len(reachable.EntryPoints))
	for _, resolved := range reachable.EntryPoints {
		entry := EntryPoint{
			Target: resolved.Root.Target, Status: resolved.Status, Confidence: string(resolved.Confidence),
			Reason: resolved.Root.Reason, Candidates: []string{},
		}
		for _, node := range resolved.Nodes {
			nodeID := node.ID.String()
			if dependencyID := dependencyProcedures[nodeID]; dependencyID != "" {
				nodeID = dependencyID
			}
			entry.Candidates = append(entry.Candidates, nodeID)
		}
		if len(entry.Candidates) == 1 {
			entry.NodeID = entry.Candidates[0]
		}
		slices.Sort(entry.Candidates)
		entryPoints = append(entryPoints, entry)
	}
	slices.SortFunc(entryPoints, func(a, b EntryPoint) int {
		if order := cmp.Compare(strings.ToLower(a.Target), strings.ToLower(b.Target)); order != 0 {
			return order
		}
		if order := cmp.Compare(a.Reason, b.Reason); order != 0 {
			return order
		}
		return cmp.Compare(a.Status, b.Status)
	})
	stateReport := projectModuleState(state, dependencyModules, stateProcedureIDs)
	effectReport := projectExcelEffects(effectSummary, procedureIDs)
	dynamicReferences := projectDynamicReferences(callResult.DynamicReferences, callerIDs)
	uncertaintyEvidence := projectCallUncertainty(callResult, callerIDs)
	external := make([]callgraph.UncertainEdge, 0)
	uncertainty := Uncertainty{Reasons: []string{}}
	if !typeDBLoaded || !typeDBComplete {
		uncertainty.Reasons = append(uncertainty.Reasons, "optional type metadata is incomplete; type-dependent analysis may be incomplete")
	}
	uncertainty.Reasons = append(uncertainty.Reasons, typeDBWarnings...)
	for _, call := range callResult.Calls {
		switch call.Resolution.Status {
		case "ambiguous":
			uncertainty.AmbiguousCallCount++
		case "unresolved", "incomplete", "non_callable":
			uncertainty.UnresolvedCallCount++
		case "dynamic", "member_call":
			uncertainty.DynamicCallCount++
		}
	}
	for _, edge := range dependencies.UncertainEdges {
		if edge.Status == "external" {
			external = append(external, edge)
		}
	}
	return Report{
		SchemaVersion: 1, Scope: Scope{Path: options.Path, Module: strings.TrimSpace(options.Module)},
		Modules: modules, Procedures: procedures, EntryPoints: entryPoints, DynamicReferences: dynamicReferences,
		Dependencies: dependencies, DependencyBoundaryNodeIDs: []string{}, Cycles: cycles, Unreachable: slices.Clone(reachable.Unreachable),
		PossibleReachability: slices.Clone(reachable.Possible), Hotspots: hotspotReport,
		ModuleState: stateReport, ExcelEffects: effectReport, ExternalDependencies: external,
		Uncertainty: Uncertainty{
			UnresolvedCallCount: uncertainty.UnresolvedCallCount,
			AmbiguousCallCount:  uncertainty.AmbiguousCallCount,
			DynamicCallCount:    uncertainty.DynamicCallCount,
			TypeDatabaseLoaded:  typeDBLoaded, TypeDatabaseComplete: typeDBComplete,
			Reasons: uncertainty.Reasons, Evidence: uncertaintyEvidence,
		},
	}
}

func metricIdentityKey(file, module, name, kind string, startByte int) string {
	return strings.ToLower(filepath.ToSlash(filepath.Clean(file)) + "\x00" + module + "\x00" + name + "\x00" + kind + "\x00" + fmt.Sprint(startByte))
}

func semanticProcedureKey(file, module, name, kind string, line int) string {
	return strings.ToLower(filepath.ToSlash(filepath.Clean(file)) + "\x00" + module + "\x00" + name + "\x00" + kind + "\x00" + fmt.Sprint(line))
}

func callerIdentityKey(file, module, qualifiedName, kind string) string {
	return strings.ToLower(filepath.ToSlash(filepath.Clean(file)) + "\x00" + module + "\x00" + qualifiedName + "\x00" + kind)
}

func sourceRange(rng ast.Range) SourceRange {
	return SourceRange{StartLine: rng.StartLine, StartColumn: rng.StartColumn, EndLine: rng.EndLine, EndColumn: rng.EndColumn, StartByte: rng.StartByte, EndByte: rng.EndByte}
}

func projectDynamicReferences(references []calls.DynamicReference, callerIDs map[string]string) []DynamicReference {
	out := make([]DynamicReference, 0, len(references))
	for _, reference := range references {
		callerID := ""
		if reference.Caller != nil {
			callerID = callerIDs[callerIdentityKey(reference.File, reference.Module, reference.Caller.QualifiedName, reference.Caller.Kind)]
		}
		out = append(out, DynamicReference{
			File: filepath.ToSlash(reference.File), Module: reference.Module, CallerID: callerID,
			API: reference.API, ArgumentIndex: reference.ArgumentIndex, ArgumentName: reference.ArgumentName,
			Expression: reference.Expression, Target: reference.Target, Kind: reference.Kind, Range: sourceRange(reference.Range),
		})
	}
	slices.SortFunc(out, func(a, b DynamicReference) int {
		if order := cmp.Compare(a.File, b.File); order != 0 {
			return order
		}
		if order := cmp.Compare(a.Range.StartByte, b.Range.StartByte); order != 0 {
			return order
		}
		if order := cmp.Compare(a.API, b.API); order != 0 {
			return order
		}
		return cmp.Compare(a.Target, b.Target)
	})
	return out
}

func projectCallUncertainty(result calls.Result, callerIDs map[string]string) []CallUncertaintyEvidence {
	out := make([]CallUncertaintyEvidence, 0)
	for _, call := range result.Calls {
		if call.Resolution.Status == "matched" {
			continue
		}
		callerID := ""
		if call.Caller != nil {
			callerID = callerIDs[callerIdentityKey(call.File, call.Module, call.Caller.QualifiedName, call.Caller.Kind)]
		}
		candidates := make([]UncertaintyCandidate, 0, len(call.Resolution.Candidates))
		for _, candidate := range call.Resolution.Candidates {
			candidates = append(candidates, UncertaintyCandidate{QualifiedName: candidate.QualifiedName, Kind: candidate.Kind, File: filepath.ToSlash(candidate.File), Line: candidate.Line})
		}
		slices.SortFunc(candidates, func(a, b UncertaintyCandidate) int {
			if order := cmp.Compare(a.QualifiedName, b.QualifiedName); order != 0 {
				return order
			}
			if order := cmp.Compare(a.File, b.File); order != 0 {
				return order
			}
			return cmp.Compare(a.Line, b.Line)
		})
		out = append(out, CallUncertaintyEvidence{
			Status: call.Resolution.Status, CallerID: callerID, File: filepath.ToSlash(call.File), Module: call.Module,
			Target: call.Callee.Text, Candidates: candidates,
			Range: sourceRange(call.Range),
		})
	}
	slices.SortFunc(out, func(a, b CallUncertaintyEvidence) int {
		if order := cmp.Compare(a.File, b.File); order != 0 {
			return order
		}
		if order := cmp.Compare(a.Range.StartByte, b.Range.StartByte); order != 0 {
			return order
		}
		return cmp.Compare(a.Status, b.Status)
	})
	return out
}

func projectModuleState(facts modulestate.Facts, moduleIDs, procedureIDs map[string]string) ModuleState {
	fields := make([]ModuleStateField, 0, len(facts.Fields))
	summary := ModuleStateSummary{}
	for _, field := range facts.Fields {
		readers := projectProcedureIDs(field.Readers, procedureIDs)
		writers := projectProcedureIDs(field.Writers, procedureIDs)
		mutators := projectProcedureIDs(field.Mutators, procedureIDs)
		summary.MutableStateReads += len(readers)
		summary.MutableStateWrites += len(writers)
		summary.MutableStateMutations += len(mutators)
		fields = append(fields, ModuleStateField{
			ID: field.ID, File: filepath.ToSlash(field.Path), ModuleID: moduleIDs[moduleKey(field.Module, filepath.ToSlash(field.Path))],
			Module: field.Module, Name: field.Name, Type: field.Type, Visibility: field.Visibility,
			Kind: field.Kind, Scope: string(field.Scope), DeclarationLine: field.DeclarationLine,
			IsObject: field.IsObject, IsCollection: field.IsCollection, IsExcel: field.IsExcel,
			Readers: readers, Writers: writers, Mutators: mutators,
		})
	}
	return ModuleState{Summary: summary, Fields: fields}
}

func projectProcedureIDs(source []string, index map[string]string) []string {
	out := make([]string, 0, len(source))
	for _, id := range source {
		if projected := index[id]; projected != "" {
			out = append(out, projected)
		}
	}
	slices.Sort(out)
	return out
}

func projectExcelEffects(summary effects.ProjectSummary, procedureIDs map[string]string) ExcelEffects {
	allowed := map[effects.EffectKind]bool{
		effects.WritesCells: true, effects.ChangesWorkbook: true, effects.OpensWorkbook: true,
		effects.ClosesWorkbook: true, effects.DisablesEvents: true, effects.RestoresEvents: true,
		effects.ChangesCalculation: true, effects.Recalculates: true, effects.ChangesSelection: true,
		effects.ChangesControls: true, effects.ShowsDialog: true, effects.ChangesApplicationState: true,
		effects.RestoresApplicationState: true,
	}
	out := ExcelEffects{Evidence: []EffectEvidence{}}
	for _, procedure := range summary.AllDirect() {
		for _, evidence := range procedure.Direct {
			if !allowed[evidence.Effect] {
				continue
			}
			target := evidence.Target
			if target == "" {
				target = evidence.Value
			}
			out.Evidence = append(out.Evidence, EffectEvidence{
				ProcedureID:   procedureIDs[effectsIdentityKey(procedure.Identity)],
				ProcedureName: procedure.Identity.QualifiedName, Kind: string(evidence.Effect), Target: target,
				File: filepath.ToSlash(procedure.Identity.File), Line: evidence.Range.StartLine, Evidence: target,
				Value: evidence.Value, Range: sourceRange(evidence.Range), StatementID: evidence.StatementID, CallID: evidence.CallID,
			})
		}
	}
	slices.SortFunc(out.Evidence, func(a, b EffectEvidence) int {
		if order := cmp.Compare(a.File, b.File); order != 0 {
			return order
		}
		if order := cmp.Compare(a.Line, b.Line); order != 0 {
			return order
		}
		if order := cmp.Compare(a.ProcedureID, b.ProcedureID); order != 0 {
			return order
		}
		return cmp.Compare(a.Kind, b.Kind)
	})
	out.DirectEffectCount = len(out.Evidence)
	return out
}

func effectsIdentityKey(identity effects.ProcedureIdentity) string {
	return strings.ToLower(strings.Join([]string{
		filepath.ToSlash(filepath.Clean(identity.File)), identity.Module, identity.QualifiedName,
		string(identity.Kind), fmt.Sprint(identity.DeclarationLine),
	}, "\x00"))
}

func moduleKey(module, file string) string {
	return strings.ToLower(strings.TrimSpace(module)) + "\x00" + filepath.ToSlash(filepath.Clean(file))
}

func summarize(report Report, snapshots []sourceSnapshot, callsResult calls.Result, reachable reachability.Result) Summary {
	summary := Summary{
		Files: len(snapshots), Modules: len(report.Modules), Procedures: len(report.Procedures), Calls: len(callsResult.Calls),
		DependencyCycles: len(report.Cycles), ExternalDependencies: len(report.ExternalDependencies),
		ProcedureHotspots: len(report.Hotspots.Procedures), ModuleHotspots: len(report.Hotspots.Modules),
	}
	for _, edge := range report.Dependencies.Edges {
		if edge.Kind == "calls" && strings.HasPrefix(edge.From, "procedure|") {
			summary.ConfirmedCallEdges++
		}
	}
	for _, edge := range report.Dependencies.UncertainEdges {
		if edge.Evidence.Kind == "call" {
			summary.UncertainCallEdges++
		}
	}
	selected := selectedCallgraphIDs(report.Procedures)
	for _, node := range reachable.Confirmed {
		if selected[node.ID.String()] {
			summary.ConfirmedReachable++
		}
	}
	for _, node := range reachable.Possible {
		if selected[node.ID.String()] {
			summary.PossiblyReachable++
		}
	}
	for _, node := range reachable.Unreachable {
		if selected[node.ID.String()] {
			summary.Unreachable++
		}
	}
	return summary
}

func scopeReport(full Report, snapshots []sourceSnapshot, callResult calls.Result, reachable reachability.Result, options Options, physicalPath string) (Report, bool) {
	selectedFiles := make(map[string]bool)
	for _, snapshot := range snapshots {
		if options.Path != "" && !samePhysicalPath(snapshot.physical, physicalPath) && !pathWithin(snapshot.physical, physicalPath) {
			continue
		}
		if options.Module != "" && !strings.EqualFold(snapshot.ir.ModuleName, strings.TrimSpace(options.Module)) {
			continue
		}
		selectedFiles[snapshot.path] = true
	}
	out := full
	out.Scope = Scope{Path: options.Path, Module: strings.TrimSpace(options.Module)}
	out.Modules = filterModules(full.Modules, selectedFiles)
	out.Procedures = filterProcedures(full.Procedures, selectedFiles)
	out.EntryPoints = filterEntryPoints(full.EntryPoints, out.Procedures)
	out.Dependencies = filterDependencyResult(full.Dependencies, selectedFiles)
	out.Cycles = filterCycles(full.Cycles, out.Procedures)
	// Related SCCs retain every member, including nonadjacent boundary nodes.
	retained := make(map[string]bool, len(out.Dependencies.Nodes))
	for _, node := range out.Dependencies.Nodes {
		retained[node.ID] = true
	}
	for _, component := range out.Cycles {
		for _, member := range component.Nodes {
			retained["procedure|"+member.String()] = true
		}
	}
	out.Dependencies.Nodes = make([]callgraph.DependencyNode, 0, len(retained))
	for _, node := range full.Dependencies.Nodes {
		if retained[node.ID] {
			out.Dependencies.Nodes = append(out.Dependencies.Nodes, node)
		}
	}
	selectedDependencyNodes := selectedDependencyNodeIDs(full.Dependencies, selectedFiles)
	boundaryIDs := make([]string, 0)
	for _, node := range out.Dependencies.Nodes {
		if !selectedDependencyNodes[node.ID] {
			boundaryIDs = append(boundaryIDs, node.ID)
		}
	}
	slices.Sort(boundaryIDs)
	out.DependencyBoundaryNodeIDs = boundaryIDs
	out.Unreachable = filterReachabilityNodes(reachable.Unreachable, out.Procedures)
	out.PossibleReachability = filterReachabilityNodes(reachable.Possible, out.Procedures)
	out.DynamicReferences = filterDynamicReferences(full.DynamicReferences, selectedFiles)
	out.Hotspots = filterHotspots(full.Hotspots, selectedFiles)
	out.ModuleState = filterModuleState(full.ModuleState, selectedFiles)
	out.ExcelEffects = filterExcelEffects(full.ExcelEffects, selectedFiles)
	out.ExternalDependencies = filterExternalDependencies(full.ExternalDependencies, out.Dependencies)
	out.Uncertainty = uncertaintyForScope(full.Uncertainty, callResult, selectedFiles)
	out.Summary = summarize(out, snapshotsForScope(snapshots, selectedFiles), callsForScope(callResult, selectedFiles), reachable)
	return out, true
}

func filterEntryPoints(input []EntryPoint, procedures []Procedure) []EntryPoint {
	selected := make(map[string]bool, len(procedures))
	for _, procedure := range procedures {
		selected[procedure.ID] = true
	}
	out := make([]EntryPoint, 0)
	for _, entry := range input {
		candidates := make([]string, 0)
		for _, id := range entry.Candidates {
			if selected[id] {
				candidates = append(candidates, id)
			}
		}
		if len(entry.Candidates) != 0 && len(candidates) == 0 {
			continue
		}
		entry.Candidates = candidates
		// Filtering does not upgrade an ambiguous root to a resolved one.
		if !selected[entry.NodeID] {
			entry.NodeID = ""
		}
		out = append(out, entry)
	}
	return out
}

func filterModules(input []Module, files map[string]bool) []Module {
	out := make([]Module, 0)
	for _, module := range input {
		if files[module.File] {
			out = append(out, module)
		}
	}
	return out
}

func filterProcedures(input []Procedure, files map[string]bool) []Procedure {
	out := make([]Procedure, 0)
	for _, procedure := range input {
		if files[procedure.File] {
			out = append(out, procedure)
		}
	}
	return out
}

func filterDependencyResult(input callgraph.DependencyResult, files map[string]bool) callgraph.DependencyResult {
	selected := selectedDependencyNodeIDs(input, files)
	retained := make(map[string]bool, len(selected))
	for id := range selected {
		retained[id] = true
	}
	edges := make([]callgraph.DependencyEdge, 0)
	for _, edge := range input.Edges {
		if selected[edge.From] || selected[edge.To] {
			edges = append(edges, edge)
			retained[edge.From], retained[edge.To] = true, true
		}
	}
	uncertain := make([]callgraph.UncertainEdge, 0)
	for _, edge := range input.UncertainEdges {
		if selected[edge.From] {
			uncertain = append(uncertain, edge)
		}
	}
	nodes := make([]callgraph.DependencyNode, 0)
	for _, node := range input.Nodes {
		if retained[node.ID] {
			nodes = append(nodes, node)
		}
	}
	return callgraph.DependencyResult{Target: input.Target, Nodes: nodes, Edges: edges, UncertainEdges: uncertain}
}

func selectedDependencyNodeIDs(input callgraph.DependencyResult, files map[string]bool) map[string]bool {
	selected := make(map[string]bool)
	for _, node := range input.Nodes {
		if files[node.File] {
			selected[node.ID] = true
		}
	}
	return selected
}

func filterCycles(input []callgraph.CyclicComponent, procedures []Procedure) []callgraph.CyclicComponent {
	selected := selectedCallgraphIDs(procedures)
	out := make([]callgraph.CyclicComponent, 0)
	for _, component := range input {
		for _, node := range component.Nodes {
			if selected[node.String()] {
				out = append(out, component)
				break
			}
		}
	}
	return out
}

func filterReachabilityNodes(input []callgraph.Node, procedures []Procedure) []callgraph.Node {
	selected := selectedCallgraphIDs(procedures)
	out := make([]callgraph.Node, 0)
	for _, node := range input {
		if selected[node.ID.String()] {
			out = append(out, node)
		}
	}
	return out
}

func typeDBResolverSymbols(db *vbadb.DB) []procedureir.ResolverSymbol {
	if db == nil {
		return []procedureir.ResolverSymbol{}
	}
	constants := db.AllConstantsList()
	out := make([]procedureir.ResolverSymbol, 0, len(constants))
	for _, constant := range constants {
		if strings.TrimSpace(constant.EnumGroup) == "" {
			continue
		}
		out = append(out, procedureir.ResolverSymbol{
			Name: constant.Name, Parent: constant.EnumGroup, Module: constant.Library,
			ModuleKind: "external", Kind: "enum_member", Visibility: "Public",
			File: "<typelib>" + constant.Library, Line: 0,
			IsConst: true, ValueShape: procedureir.ValueShapeScalar,
		})
	}
	slices.SortFunc(out, func(a, b procedureir.ResolverSymbol) int {
		if order := cmp.Compare(strings.ToLower(a.Module), strings.ToLower(b.Module)); order != 0 {
			return order
		}
		if order := cmp.Compare(strings.ToLower(a.Parent), strings.ToLower(b.Parent)); order != 0 {
			return order
		}
		return cmp.Compare(strings.ToLower(a.Name), strings.ToLower(b.Name))
	})
	return out
}

func selectedCallgraphIDs(procedures []Procedure) map[string]bool {
	selected := make(map[string]bool, len(procedures))
	for _, procedure := range procedures {
		id := procedure.CallgraphID
		if id == "" {
			id = strings.TrimPrefix(procedure.ID, "procedure|")
		}
		selected[id] = true
	}
	return selected
}

func filterHotspots(input hotspots.Report, files map[string]bool) hotspots.Report {
	out := hotspots.Report{SchemaVersion: input.SchemaVersion, ScoreModel: input.ScoreModel, Procedures: []hotspots.Entity{}, Modules: []hotspots.Entity{}}
	for _, entity := range input.Procedures {
		if files[entity.File] {
			out.Procedures = append(out.Procedures, entity)
		}
	}
	for _, entity := range input.Modules {
		if files[entity.File] {
			out.Modules = append(out.Modules, entity)
		}
	}
	return out
}

func filterModuleState(input ModuleState, files map[string]bool) ModuleState {
	out := ModuleState{Summary: ModuleStateSummary{}, Fields: []ModuleStateField{}}
	for _, field := range input.Fields {
		if !files[field.File] {
			continue
		}
		out.Fields = append(out.Fields, field)
		out.Summary.MutableStateReads += len(field.Readers)
		out.Summary.MutableStateWrites += len(field.Writers)
		out.Summary.MutableStateMutations += len(field.Mutators)
	}
	return out
}

func filterExcelEffects(input ExcelEffects, files map[string]bool) ExcelEffects {
	out := ExcelEffects{Evidence: []EffectEvidence{}}
	for _, evidence := range input.Evidence {
		if files[evidence.File] {
			out.Evidence = append(out.Evidence, evidence)
		}
	}
	out.DirectEffectCount = len(out.Evidence)
	return out
}

func filterExternalDependencies(input []callgraph.UncertainEdge, dependencies callgraph.DependencyResult) []callgraph.UncertainEdge {
	selected := make(map[string]bool)
	for _, edge := range dependencies.UncertainEdges {
		selected[uncertainEdgeKey(edge)] = true
	}
	out := make([]callgraph.UncertainEdge, 0)
	for _, edge := range input {
		if selected[uncertainEdgeKey(edge)] {
			out = append(out, edge)
		}
	}
	return out
}

func uncertainEdgeKey(edge callgraph.UncertainEdge) string {
	return strings.Join([]string{edge.Status, edge.From, edge.Callee, edge.Evidence.File, fmt.Sprint(edge.Evidence.Location.StartLine), fmt.Sprint(edge.Evidence.Location.StartColumn)}, "\x00")
}

func uncertaintyForScope(full Uncertainty, callsResult calls.Result, files map[string]bool) Uncertainty {
	out := Uncertainty{
		TypeDatabaseLoaded: full.TypeDatabaseLoaded, TypeDatabaseComplete: full.TypeDatabaseComplete,
		Reasons: slices.Clone(full.Reasons), Evidence: make([]CallUncertaintyEvidence, 0),
	}
	for _, evidence := range full.Evidence {
		if files[evidence.File] {
			out.Evidence = append(out.Evidence, evidence)
		}
	}
	for _, call := range callsResult.Calls {
		if call.Caller == nil || !files[call.File] {
			continue
		}
		switch call.Resolution.Status {
		case "ambiguous":
			out.AmbiguousCallCount++
		case "unresolved", "incomplete", "non_callable":
			out.UnresolvedCallCount++
		case "dynamic", "member_call":
			out.DynamicCallCount++
		}
	}
	return out
}

func filterDynamicReferences(input []DynamicReference, files map[string]bool) []DynamicReference {
	out := make([]DynamicReference, 0)
	for _, reference := range input {
		if files[reference.File] {
			out = append(out, reference)
		}
	}
	return out
}

func snapshotsForScope(snapshots []sourceSnapshot, files map[string]bool) []sourceSnapshot {
	out := make([]sourceSnapshot, 0)
	for _, snapshot := range snapshots {
		if files[snapshot.path] {
			out = append(out, snapshot)
		}
	}
	return out
}

func callsForScope(input calls.Result, files map[string]bool) calls.Result {
	out := calls.Result{Calls: []calls.Call{}}
	for _, call := range input.Calls {
		if files[call.File] {
			out.Calls = append(out.Calls, call)
		}
	}
	return out
}
