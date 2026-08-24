package analyze

import (
	"context"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/gui"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// Dictionary and Collection safety rules intentionally share one flow model.
// Object identity, rather than the spelling of a local variable, is the unit
// of state so direct aliases receive the same key and mutation facts.
type dictionaryCollectionKind uint8

const (
	dcUnknown dictionaryCollectionKind = iota
	dcDictionary
	dcCollection
)

type dcObjectState struct {
	Kind           dictionaryCollectionKind
	LateBound      bool
	Empty          int // 1 definitely empty, -1 definitely non-empty, 0 unknown
	CompareMode    string
	Keys           map[string]int // 1 present, -1 absent; missing is unknown
	Normalizations map[string]string
}

type dcKeyExpr struct {
	Base string
	Norm string
}

type dcFlowState struct {
	Bindings map[string]string
	Objects  map[string]dcObjectState
	Scalars  map[string]dcKeyExpr
}

type dcHelperEffect struct {
	Param  int
	Member string
	KeyArg int
}

type dcHelperSummary struct {
	Name          string
	Factory       dictionaryCollectionKind
	FactoryLate   bool
	Effects       []dcHelperEffect
	ExistsObject  int
	ExistsKey     int
	Materializes  bool
	SourceFile    string
	DeclarationAt int
}

type dictionaryCollectionIndex struct {
	Helpers map[string]dcHelperSummary
}

func dictionaryCollectionAnalysisEnabled(cfg config.AnalyzeConfig) bool {
	return cfg.DetectDictionaryCollectionGuard ||
		cfg.DetectDictionaryIterationValueUsage ||
		cfg.DetectDictionaryCompareModeOrder ||
		cfg.DetectDictionaryLoopMaterialization ||
		cfg.DetectDictionaryKeyNormalization ||
		cfg.DetectLateBoundDictionaryConstants ||
		cfg.DetectCollectionIterationMutation ||
		cfg.DetectCollectionIndexOrigin
}

func buildDictionaryCollectionIndex(files []parsedFile) *dictionaryCollectionIndex {
	index := &dictionaryCollectionIndex{Helpers: map[string]dcHelperSummary{}}
	duplicates := map[string]bool{}
	type source struct {
		file parsedFile
		proc sourceProcedure
	}
	sources := map[string]source{}
	for _, file := range files {
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			if proc.Name == "" {
				continue
			}
			summary := directDictionaryCollectionSummary(file, proc)
			key := strings.ToLower(proc.Name)
			if _, exists := index.Helpers[key]; exists {
				duplicates[key] = true
				continue
			}
			index.Helpers[key] = summary
			sources[key] = source{file: file, proc: proc}
		}
	}
	for key := range duplicates {
		delete(index.Helpers, key)
		delete(sources, key)
	}
	// Propagate uniquely resolved local effects to a fixed point. The domain is
	// finite (factory kind, exists forwarding, materialization, and parameter
	// effects), so at most one pass per helper is sufficient for convergence.
	for pass := 0; pass < len(sources); pass++ {
		changed := false
		for key, src := range sources {
			summary := index.Helpers[key]
			if dcPropagateHelperSummary(src.proc, &summary, index.Helpers) {
				index.Helpers[key] = summary
				changed = true
			}
		}
		if !changed {
			break
		}
	}
	return index
}

func dcPropagateHelperSummary(proc sourceProcedure, summary *dcHelperSummary, helpers map[string]dcHelperSummary) bool {
	params := map[string]int{}
	for i, param := range proc.Params.AllIndexed() {
		params[strings.ToLower(param.Name)] = i
	}
	changed := false
	for statement := range proc.Statements.All() {
		text := dcStatementSource(statement)
		if assignment := dcSetAssignment(text); assignment != nil && strings.EqualFold(assignment.target, proc.Name) {
			if name, args, ok := dcSimpleFunctionCall(assignment.rhs); ok {
				callee, found := helpers[strings.ToLower(lastName(name))]
				if found && summary.Factory == dcUnknown && callee.Factory != dcUnknown {
					summary.Factory, summary.FactoryLate = callee.Factory, callee.FactoryLate
					changed = true
				}
				if found && summary.ExistsObject < 0 && callee.ExistsObject >= 0 && callee.ExistsKey >= 0 && callee.ExistsObject < len(args) && callee.ExistsKey < len(args) {
					objectParam, objectOK := params[strings.ToLower(cleanIdentifier(args[callee.ExistsObject]))]
					keyParam, keyOK := params[strings.ToLower(cleanIdentifier(args[callee.ExistsKey]))]
					if objectOK && keyOK {
						summary.ExistsObject, summary.ExistsKey = objectParam, keyParam
						changed = true
					}
				}
			}
		}
		name, args, ok := parseSimpleCall(text)
		if !ok {
			continue
		}
		callee, found := helpers[strings.ToLower(lastName(name))]
		if !found {
			continue
		}
		if callee.Materializes && !summary.Materializes {
			summary.Materializes = true
			changed = true
		}
		for _, effect := range callee.Effects {
			if effect.Param < 0 || effect.Param >= len(args) {
				continue
			}
			param, found := params[strings.ToLower(cleanIdentifier(args[effect.Param]))]
			if !found {
				continue
			}
			keyParam := -1
			if effect.KeyArg >= 0 && effect.KeyArg < len(args) {
				if mapped, found := params[strings.ToLower(cleanIdentifier(args[effect.KeyArg]))]; found {
					keyParam = mapped
				}
			}
			candidate := dcHelperEffect{Param: param, Member: effect.Member, KeyArg: keyParam}
			if !dcHasHelperEffect(summary.Effects, candidate) {
				summary.Effects = append(summary.Effects, candidate)
				changed = true
			}
		}
	}
	return changed
}

func dcHasHelperEffect(effects []dcHelperEffect, candidate dcHelperEffect) bool {
	for _, effect := range effects {
		if effect == candidate {
			return true
		}
	}
	return false
}

func directDictionaryCollectionSummary(file parsedFile, proc sourceProcedure) dcHelperSummary {
	summary := dcHelperSummary{Name: proc.Name, ExistsObject: -1, ExistsKey: -1, SourceFile: file.Path, DeclarationAt: proc.StartLine}
	params := map[string]int{}
	for i, param := range proc.Params.AllIndexed() {
		params[strings.ToLower(param.Name)] = i
	}
	returnName := strings.ToLower(proc.Name)
	for statement := range proc.Statements.All() {
		text := dcStatementSource(statement)
		lower := strings.ToLower(text)
		if match := dcSetAssignment(text); match != nil && strings.EqualFold(match.target, returnName) {
			if kind, late, ok := dcConstruction(match.rhs); ok {
				summary.Factory, summary.FactoryLate = kind, late
			}
		}
		if receiver, member, args, ok := dcMemberCall(text); ok {
			if param, exists := params[strings.ToLower(receiver)]; exists {
				switch strings.ToLower(member) {
				case "add", "remove", "removeall", "comparemode":
					keyArg := -1
					if strings.EqualFold(member, "add") && dcDeclaredKind(proc, receiver) == dcCollection {
						if len(args) > 1 {
							if candidate, found := params[strings.ToLower(cleanIdentifier(args[1]))]; found {
								keyArg = candidate
							}
						}
					} else if len(args) > 0 {
						if candidate, found := params[strings.ToLower(cleanIdentifier(args[0]))]; found {
							keyArg = candidate
						}
					}
					summary.Effects = append(summary.Effects, dcHelperEffect{Param: param, Member: strings.ToLower(member), KeyArg: keyArg})
				case "keys", "items":
					if dcDeclaredKind(proc, receiver) == dcDictionary {
						summary.Materializes = true
					}
				}
			}
		}
		if strings.Contains(lower, returnName+" = ") {
			for name, objectParam := range params {
				needle := name + ".exists("
				at := strings.Index(lower, needle)
				if at < 0 {
					continue
				}
				arg := dcFirstArgument(text[at+len(needle):])
				if keyParam, ok := params[strings.ToLower(cleanIdentifier(arg))]; ok {
					summary.ExistsObject, summary.ExistsKey = objectParam, keyParam
				}
			}
		}
	}
	return summary
}

func dcDeclaredKind(proc sourceProcedure, name string) dictionaryCollectionKind {
	for decl := range proc.Declarations.All() {
		if strings.EqualFold(decl.Name, name) {
			return dcKindFromType(decl.Type)
		}
	}
	for param := range proc.Params.All() {
		if strings.EqualFold(param.Name, name) {
			return dcKindFromType(param.Type)
		}
	}
	return dcUnknown
}

func (a Analyzer) dictionaryCollectionSafetyFindings(cancelCtx context.Context, file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) ([]Finding, error) {
	if err := cancelCtx.Err(); err != nil {
		return nil, err
	}
	if proc.Graph == nil {
		return nil, nil
	}
	initial := dcInitialState(file, proc, moduleDecls)
	in := map[vbacfg.BlockID]*dcFlowState{proc.Graph.Entry: initial}
	queue := []vbacfg.BlockID{proc.Graph.Entry}
	queued := map[vbacfg.BlockID]bool{proc.Graph.Entry: true}
	blocks := map[vbacfg.BlockID]vbacfg.Block{}
	edges := map[vbacfg.BlockID][]vbacfg.Edge{}
	for _, block := range proc.Graph.Blocks {
		blocks[block.ID] = block
	}
	for _, edge := range proc.Graph.Edges {
		if edge.Class == vbacfg.EdgeNormal {
			edges[edge.From] = append(edges[edge.From], edge)
		}
	}
	for len(queue) > 0 {
		if err := cancelCtx.Err(); err != nil {
			return nil, err
		}
		id := queue[0]
		queue = queue[1:]
		queued[id] = false
		state := in[id]
		if state == nil {
			continue
		}
		out := state.clone()
		if block := blocks[id]; block.Statement != nil {
			a.dcTransfer(file, proc, *block.Statement, out)
		}
		for _, edge := range edges[id] {
			candidate := out.clone()
			if block := blocks[id]; block.Statement != nil {
				a.dcApplyGuard(*block.Statement, edge, candidate)
			}
			merged, changed := dcMergeState(in[edge.To], candidate)
			if !changed {
				continue
			}
			in[edge.To] = merged
			if !queued[edge.To] {
				queue = append(queue, edge.To)
				queued[edge.To] = true
			}
		}
	}

	narrowProbes := dcNarrowProbeStatements(proc)
	loops := excelLoopRegions(proc)
	// Normal batch/realtime setup attaches module facts before rule execution.
	// Read the field directly here so standalone callers retain the cheap source
	// fallback instead of rebuilding the complete module index for each file.
	constFacts := file.ModuleFacts
	var constNames map[string]bool
	if constFacts == nil {
		// Keep standalone package-level callers (and focused tests that build a
		// parsedFile directly) on the historical source fallback, filtered to
		// module constants and the current procedure's local constants.
		constNames = dcConstantNamesForProcedure(file, proc)
	}
	seen := map[string]bool{}
	var findings []Finding
	linear := initial.clone()
	for statement := range proc.Statements.All() {
		if err := cancelCtx.Err(); err != nil {
			return nil, err
		}
		block, ok := proc.Graph.BlockForStatement(statement.ID)
		if !ok {
			continue
		}
		state := dcOverlayState(in[block.ID], linear)
		findings = append(findings, a.dcStatementFindings(file, proc, statement, state, loops, narrowProbes, constFacts, constNames, seen)...)
		a.dcTransfer(file, proc, statement, linear)
	}
	return findings, nil
}

func dcInitialState(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) *dcFlowState {
	state := &dcFlowState{Bindings: map[string]string{}, Objects: map[string]dcObjectState{}, Scalars: map[string]dcKeyExpr{}}
	add := func(key string, decl sourceDeclaration, local bool) {
		kind := dcKindFromType(decl.Type)
		if kind == dcUnknown {
			return
		}
		id := "module:" + strings.ToLower(file.Module) + ":" + key
		empty := 0
		if local {
			id = "local:" + strings.ToLower(proc.Name) + ":" + key
			if decl.NewExpression {
				empty = 1
			}
		}
		state.Bindings[key] = id
		state.Objects[id] = dcObjectState{Kind: kind, Empty: empty, CompareMode: dcDefaultCompareMode(kind, empty == 1), Keys: map[string]int{}, Normalizations: map[string]string{}}
	}
	for key, decl := range moduleDecls {
		add(key, decl, false)
	}
	for key, decl := range file.procedureDeclarationsFor(proc) {
		add(key, decl, true)
	}
	for param := range proc.Params.All() {
		kind := dcKindFromType(param.Type)
		if kind == dcUnknown {
			continue
		}
		key := strings.ToLower(param.Name)
		id := "param:" + strings.ToLower(proc.Name) + ":" + key
		state.Bindings[key] = id
		state.Objects[id] = dcObjectState{Kind: kind, Keys: map[string]int{}, Normalizations: map[string]string{}}
	}
	return state
}

func dcOverlayState(flow, linear *dcFlowState) *dcFlowState {
	if flow == nil {
		return linear.clone()
	}
	out := flow.clone()
	for key, id := range linear.Bindings {
		if _, exists := out.Bindings[key]; !exists {
			out.Bindings[key] = id
		}
		if object, ok := linear.Objects[id]; ok {
			if flowObject, exists := out.Objects[id]; exists {
				object.Empty = flowObject.Empty
				object.Keys = dcCloneIntMap(flowObject.Keys)
				object.Normalizations = dcCloneStringMap(flowObject.Normalizations)
			}
			out.Objects[id] = object
		}
	}
	for key, value := range linear.Scalars {
		if _, exists := out.Scalars[key]; !exists {
			out.Scalars[key] = value
		}
	}
	return out
}

func dcKindFromType(typ string) dictionaryCollectionKind {
	switch strings.ToLower(cleanIdentifier(typ)) {
	case "dictionary", "scripting.dictionary":
		return dcDictionary
	case "collection", "vba.collection":
		return dcCollection
	default:
		return dcUnknown
	}
}

func dcDefaultCompareMode(kind dictionaryCollectionKind, fresh bool) string {
	if kind == dcDictionary && fresh {
		return "binary"
	}
	return ""
}

type dcAssignment struct{ target, rhs string }

func dcSetAssignment(text string) *dcAssignment {
	lower := strings.ToLower(strings.TrimSpace(text))
	if !strings.HasPrefix(lower, "set ") {
		return nil
	}
	parts := strings.SplitN(strings.TrimSpace(text)[4:], "=", 2)
	if len(parts) != 2 {
		return nil
	}
	return &dcAssignment{target: cleanIdentifier(strings.TrimSpace(parts[0])), rhs: strings.TrimSpace(parts[1])}
}

func dcValueAssignment(text string) *dcAssignment {
	trimmed := strings.TrimSpace(text)
	lower := strings.ToLower(trimmed)
	if strings.HasPrefix(lower, "set ") || strings.HasPrefix(lower, "if ") || strings.HasPrefix(lower, "elseif ") {
		return nil
	}
	if strings.HasPrefix(lower, "let ") {
		trimmed = strings.TrimSpace(trimmed[4:])
	}
	parts := strings.SplitN(trimmed, "=", 2)
	if len(parts) != 2 || strings.Contains(parts[0], ".") || strings.Contains(parts[0], "(") {
		return nil
	}
	target := cleanIdentifier(strings.TrimSpace(parts[0]))
	if target == "" {
		return nil
	}
	return &dcAssignment{target: target, rhs: strings.TrimSpace(parts[1])}
}

func dcConstruction(rhs string) (dictionaryCollectionKind, bool, bool) {
	lower := strings.ToLower(strings.Join(strings.Fields(rhs), " "))
	switch {
	case lower == "new collection", lower == "new vba.collection":
		return dcCollection, false, true
	case lower == "new dictionary", lower == "new scripting.dictionary":
		return dcDictionary, false, true
	case dictionaryCreateRe.MatchString(rhs):
		return dcDictionary, true, true
	default:
		return dcUnknown, false, false
	}
}

func (a Analyzer) dcTransfer(file parsedFile, proc sourceProcedure, statement procedureir.Statement, state *dcFlowState) {
	text := dcStatementSource(statement)
	if text == "" {
		return
	}
	if assignment := dcSetAssignment(text); assignment != nil {
		target := strings.ToLower(assignment.target)
		if kind, late, ok := dcConstruction(assignment.rhs); ok {
			id := "construct:" + strings.ToLower(file.Path) + ":" + strconv.Itoa(statement.ID) + ":" + target
			state.Bindings[target] = id
			state.Objects[id] = dcObjectState{Kind: kind, LateBound: late, Empty: 1, CompareMode: dcDefaultCompareMode(kind, true), Keys: map[string]int{}, Normalizations: map[string]string{}}
			return
		}
		if source := strings.ToLower(cleanIdentifier(assignment.rhs)); state.Bindings[source] != "" && dcBareIdentifier(assignment.rhs) {
			state.Bindings[target] = state.Bindings[source]
			return
		}
		if name, args, ok := dcSimpleFunctionCall(assignment.rhs); ok {
			if summary, found := a.dcHelper(name); found && summary.Factory != dcUnknown {
				id := "factory:" + strings.ToLower(file.Path) + ":" + strconv.Itoa(statement.ID) + ":" + target
				state.Bindings[target] = id
				state.Objects[id] = dcObjectState{Kind: summary.Factory, LateBound: summary.FactoryLate, Empty: 1, CompareMode: dcDefaultCompareMode(summary.Factory, true), Keys: map[string]int{}, Normalizations: map[string]string{}}
				a.dcApplyHelperEffects(summary, args, state)
				return
			}
		}
		delete(state.Bindings, target)
		return
	}
	if assignment := dcValueAssignment(text); assignment != nil {
		key := strings.ToLower(assignment.target)
		if expr, ok := dcParseKeyExpr(assignment.rhs, state.Scalars); ok {
			state.Scalars[key] = expr
		} else {
			delete(state.Scalars, key)
		}
	}
	if receiver, member, args, ok := dcMemberCall(text); ok {
		a.dcApplyMember(strings.ToLower(receiver), strings.ToLower(member), args, state)
	}
	if receiver, keyText, ok := dcDefaultAssignment(text); ok {
		if id := state.Bindings[strings.ToLower(receiver)]; id != "" {
			object := state.Objects[id]
			if object.Kind == dcDictionary {
				expr, parsed := dcParseKeyExpr(keyText, state.Scalars)
				if parsed {
					object.Keys[dcKeyIdentity(expr)] = 1
					dcObserveNormalization(&object, expr)
				}
				object.Empty = -1
				state.Objects[id] = object
			}
		}
	}
	if name, args, ok := parseSimpleCall(text); ok {
		if dcCallTargetUnqualified(text) {
			if summary, found := a.dcHelper(name); found {
				a.dcApplyHelperEffects(summary, args, state)
				return
			}
		}
		for _, arg := range args {
			if id := state.Bindings[strings.ToLower(cleanIdentifier(arg))]; id != "" && dcBareIdentifier(arg) {
				object := state.Objects[id]
				object.Empty, object.CompareMode = 0, ""
				object.Keys = map[string]int{}
				object.Normalizations = map[string]string{}
				state.Objects[id] = object
			}
		}
	}
}

func (a Analyzer) dcApplyMember(receiver, member string, args []string, state *dcFlowState) {
	id := state.Bindings[receiver]
	if id == "" {
		return
	}
	object := state.Objects[id]
	switch member {
	case "add":
		keyArg := 0
		if object.Kind == dcCollection {
			keyArg = 1
		}
		if keyArg < len(args) && strings.TrimSpace(args[keyArg]) != "" {
			if expr, ok := dcParseKeyExpr(args[keyArg], state.Scalars); ok {
				object.Keys[dcKeyIdentity(expr)] = 1
				dcObserveNormalization(&object, expr)
			}
		}
		object.Empty = -1
	case "remove":
		if len(args) > 0 {
			if expr, ok := dcParseKeyExpr(args[0], state.Scalars); ok {
				object.Keys[dcKeyIdentity(expr)] = -1
			}
		}
		object.Empty = 0
	case "removeall":
		object.Empty = 1
		object.Keys = map[string]int{}
		object.Normalizations = map[string]string{}
	case "comparemode":
		if len(args) > 0 && object.Empty != -1 {
			object.CompareMode = dcCompareMode(args[0])
		}
	}
	state.Objects[id] = object
}

func (a Analyzer) dcApplyHelperEffects(summary dcHelperSummary, args []string, state *dcFlowState) {
	for _, effect := range summary.Effects {
		if effect.Param < 0 || effect.Param >= len(args) {
			continue
		}
		receiver := strings.ToLower(cleanIdentifier(args[effect.Param]))
		id := state.Bindings[receiver]
		if id == "" || !dcBareIdentifier(args[effect.Param]) {
			continue
		}
		object := state.Objects[id]
		switch effect.Member {
		case "add":
			object.Empty = -1
			if effect.KeyArg >= 0 && effect.KeyArg < len(args) {
				if expr, ok := dcParseKeyExpr(args[effect.KeyArg], state.Scalars); ok {
					object.Keys[dcKeyIdentity(expr)] = 1
					dcObserveNormalization(&object, expr)
				}
			}
		case "remove":
			object.Empty = 0
			if effect.KeyArg >= 0 && effect.KeyArg < len(args) {
				if expr, ok := dcParseKeyExpr(args[effect.KeyArg], state.Scalars); ok {
					object.Keys[dcKeyIdentity(expr)] = -1
				}
			}
		case "removeall":
			object.Empty = 1
			object.Keys = map[string]int{}
			object.Normalizations = map[string]string{}
		case "comparemode":
			object.CompareMode = ""
		}
		state.Objects[id] = object
	}
}

func (a Analyzer) dcHelper(name string) (dcHelperSummary, bool) {
	if a.dictionaryCollection == nil {
		return dcHelperSummary{}, false
	}
	summary, ok := a.dictionaryCollection.Helpers[strings.ToLower(lastName(name))]
	return summary, ok
}

func (a Analyzer) dcApplyGuard(statement procedureir.Statement, edge vbacfg.Edge, state *dcFlowState) {
	if statement.Kind != procedureir.StatementIf && statement.Kind != procedureir.StatementElseIf {
		return
	}
	receiver, keyText, negated, ok := dcExistsCondition(statement.Text)
	if !ok {
		if name, args, neg, helperOK := dcHelperCondition(statement.Text); helperOK {
			summary, found := a.dcHelper(name)
			if !found || summary.ExistsObject < 0 || summary.ExistsKey < 0 || summary.ExistsObject >= len(args) || summary.ExistsKey >= len(args) {
				return
			}
			receiver, keyText, negated = cleanIdentifier(args[summary.ExistsObject]), args[summary.ExistsKey], neg
		} else {
			return
		}
	}
	present := edge.Kind == vbacfg.EdgeBranchTrue
	if negated {
		present = !present
	}
	id := state.Bindings[strings.ToLower(receiver)]
	if id == "" {
		return
	}
	object := state.Objects[id]
	expr, parsed := dcParseKeyExpr(keyText, state.Scalars)
	if !parsed {
		return
	}
	if present {
		object.Keys[dcKeyIdentity(expr)] = 1
		object.Empty = -1
	} else {
		object.Keys[dcKeyIdentity(expr)] = -1
	}
	state.Objects[id] = object
}

func (a Analyzer) dcStatementFindings(file parsedFile, proc sourceProcedure, statement procedureir.Statement, state *dcFlowState, loops []excelLoopRegion, narrow map[int]bool, constFacts *moduleAnalysisFacts, constNames map[string]bool, seen map[string]bool) []Finding {
	text := maskStringLiterals(dcStatementSource(statement))
	line := statement.Range.StartLine
	if text == "" {
		return nil
	}
	var findings []Finding
	receiver, member, args, memberCall := dcMemberCall(text)
	if memberCall {
		_, object := dcObjectForReceiver(state, receiver)
		if object.Kind == dcUnknown {
			receiver, member, args, memberCall = dcKnownMemberCall(text, state)
		}
	}
	if memberCall {
		id, object := dcObjectForReceiver(state, receiver)
		lowerMember := strings.ToLower(member)
		if a.Config.Analyze.DetectDictionaryCollectionGuard && !narrow[statement.ID] && (lowerMember == "item" || lowerMember == "remove") && len(args) > 0 && !dcAccessGuarded(proc, statement, receiver, args[0]) {
			findings = append(findings, a.dcGuardFinding(file, proc, line, receiver, args[0], object, state, seen)...)
		}
		if a.Config.Analyze.DetectDictionaryCompareModeOrder && object.Kind == dcDictionary && lowerMember == "comparemode" && object.Empty == -1 {
			findings = dcAppendFinding(findings, seen, a.simpleFinding(file, proc, line, "VBA230", "warning", receiver+".CompareMode is changed after entries were added.", "Scripting.Dictionary rejects CompareMode changes while it contains entries.", "Set CompareMode immediately after construction, or call RemoveAll before changing it."))
		}
		if a.Config.Analyze.DetectLateBoundDictionaryConstants && object.Kind == dcDictionary && object.LateBound && lowerMember == "comparemode" && len(args) > 0 {
			name := strings.ToLower(cleanIdentifier(args[0]))
			declared := false
			if constFacts != nil {
				declared = constFacts.hasConstantForProcedure(name, proc)
			} else {
				declared = constNames[strings.ToLower(strings.TrimSpace(name))]
			}
			if dcLateCompareConstants[name] && !declared {
				findings = dcAppendFinding(findings, seen, a.simpleFinding(file, proc, line, "VBA233", "warning", args[0]+" is an undefined Scripting enum constant in late-bound Dictionary code.", "Late binding should not depend on enum names supplied only by a Scripting Runtime reference.", "Use vbBinaryCompare, vbTextCompare, or vbDatabaseCompare, or declare an explicit project Const."))
			}
		}
		if a.Config.Analyze.DetectDictionaryKeyNormalization && object.Kind == dcDictionary && (lowerMember == "add" || lowerMember == "exists" || lowerMember == "item" || lowerMember == "remove") {
			if len(args) > 0 {
				findings = append(findings, a.dcNormalizationFinding(file, proc, line, receiver, args[0], object, state, seen)...)
			}
		}
		if a.Config.Analyze.DetectDictionaryLoopMaterialization && object.Kind == dcDictionary && (lowerMember == "keys" || lowerMember == "items") && dcStatementInRepeatedLoop(statement, loops) && !dcLoopHeaderMaterialization(statement, loops) {
			findings = dcAppendFinding(findings, seen, a.simpleFinding(file, proc, line, "VBA231", "warning", receiver+"."+member+" is evaluated inside a loop.", "Dictionary Keys and Items create a new array on each evaluation.", "Cache "+receiver+"."+member+" before the loop and reuse the array."))
		}
		if a.Config.Analyze.DetectCollectionIndexOrigin && object.Kind == dcCollection && lowerMember == "item" && len(args) > 0 && dcZeroIndex(args[0], statement, proc) {
			findings = dcAppendFinding(findings, seen, a.simpleFinding(file, proc, line, "VBA235", "warning", receiver+" uses a zero-based index for a Collection item.", "VBA Collection indexes are one-based even when a related array or loop is zero-based.", "Use indexes from 1 To "+receiver+".Count, or add 1 when converting a zero-based array index."))
		}
		if a.Config.Analyze.DetectCollectionIterationMutation && object.Kind == dcCollection && (lowerMember == "add" || lowerMember == "remove") && dcMutatesEnumeratedCollection(statement, id, state, proc, loops) {
			findings = dcAppendFinding(findings, seen, a.simpleFinding(file, proc, line, "VBA234", "warning", receiver+" is mutated while it is being enumerated.", "Changing a Collection during For Each can skip elements or fail at runtime.", "Collect pending changes and apply them after iteration, or use a deliberate reverse index loop for removals."))
		}
	}
	if a.Config.Analyze.DetectDictionaryCollectionGuard && !narrow[statement.ID] {
		if defaultReceiver, keyText, ok := dcDefaultAccess(text); ok {
			if !dcAccessGuarded(proc, statement, defaultReceiver, keyText) && !dcCollectionIndexGuarded(proc, statement, defaultReceiver, keyText) {
				_, object := dcObjectForReceiver(state, defaultReceiver)
				findings = append(findings, a.dcGuardFinding(file, proc, line, defaultReceiver, keyText, object, state, seen)...)
			}
		}
	}
	if a.Config.Analyze.DetectDictionaryKeyNormalization {
		if defaultReceiver, keyText, ok := dcDefaultAccess(text); ok {
			_, object := dcObjectForReceiver(state, defaultReceiver)
			if object.Kind == dcDictionary {
				findings = append(findings, a.dcNormalizationFinding(file, proc, line, defaultReceiver, keyText, object, state, seen)...)
			}
		}
	}
	if a.Config.Analyze.DetectCollectionIndexOrigin {
		for _, access := range dcKnownDefaultAccesses(text, state) {
			defaultReceiver, keyText := access[0], access[1]
			_, object := dcObjectForReceiver(state, defaultReceiver)
			if object.Kind == dcCollection && dcZeroIndex(keyText, statement, proc) {
				findings = dcAppendFinding(findings, seen, a.simpleFinding(file, proc, line, "VBA235", "warning", defaultReceiver+" uses a zero-based index for a Collection item.", "VBA Collection indexes are one-based even when a related array or loop is zero-based.", "Use indexes from 1 To "+defaultReceiver+".Count, or add 1 when converting a zero-based array index."))
			}
		}
	}
	if name, callArgs, ok := parseSimpleCall(text); ok && dcCallTargetUnqualified(text) {
		summary, found := a.dcHelper(name)
		if found && summary.Materializes && a.Config.Analyze.DetectDictionaryLoopMaterialization && dcStatementInRepeatedLoop(statement, loops) {
			findings = dcAppendFinding(findings, seen, a.simpleFinding(file, proc, line, "VBA231", "warning", name+" repeatedly materializes Dictionary Keys or Items inside a loop.", "The resolved helper evaluates a Dictionary snapshot on each invocation.", "Move the helper call before the loop or pass a cached Keys/Items array."))
		}
		if found && a.Config.Analyze.DetectCollectionIterationMutation {
			for _, effect := range summary.Effects {
				if effect.Member != "add" && effect.Member != "remove" || effect.Param < 0 || effect.Param >= len(callArgs) {
					continue
				}
				receiver := cleanIdentifier(callArgs[effect.Param])
				id, object := dcObjectForReceiver(state, receiver)
				if object.Kind == dcCollection && dcMutatesEnumeratedCollection(statement, id, state, proc, loops) {
					findings = dcAppendFinding(findings, seen, a.simpleFinding(file, proc, line, "VBA234", "warning", receiver+" is mutated by "+name+" while it is being enumerated.", "The resolved helper changes the same Collection used by the enclosing For Each.", "Apply the helper after iteration, or pass a different Collection."))
					break
				}
			}
		}
	}
	return findings
}

func dcAccessGuarded(proc sourceProcedure, statement procedureir.Statement, receiver, keyText string) bool {
	want, ok := dcParseKeyExpr(keyText, map[string]dcKeyExpr{})
	if !ok {
		return false
	}
	statements := map[int]procedureir.Statement{}
	for candidate := range proc.Statements.All() {
		statements[candidate.ID] = candidate
	}
	for current := statement; current.ID != 0; current = statements[current.ParentID] {
		guardReceiver, guardKey, negated, found := dcExistsCondition(current.Text)
		if found && !negated && strings.EqualFold(guardReceiver, receiver) {
			if got, parsed := dcParseKeyExpr(guardKey, map[string]dcKeyExpr{}); parsed && got == want {
				return true
			}
		}
		if current.ParentID == 0 {
			break
		}
	}
	return false
}

func dcCollectionIndexGuarded(proc sourceProcedure, statement procedureir.Statement, receiver, keyText string) bool {
	key := strings.ToLower(cleanIdentifier(keyText))
	if !dcBareIdentifier(key) {
		return false
	}
	statements := map[int]procedureir.Statement{}
	for candidate := range proc.Statements.All() {
		statements[candidate.ID] = candidate
	}
	for current := statement; current.ID != 0; current = statements[current.ParentID] {
		header := strings.ToLower(strings.Join(strings.Fields(dcStatementSource(current)), " "))
		if strings.HasPrefix(header, "for "+key+" = 1 to ") && strings.Contains(header, strings.ToLower(receiver)+".count") {
			return true
		}
		if current.ParentID == 0 {
			break
		}
	}
	return false
}

func dcStatementSource(statement procedureir.Statement) string {
	text := statement.Text
	if newline := strings.IndexAny(text, "\r\n"); newline >= 0 {
		text = text[:newline]
	}
	return strings.Join(strings.Fields(gui.StripComment(text)), " ")
}

func (a Analyzer) dcGuardFinding(file parsedFile, proc sourceProcedure, line int, receiver, keyText string, object dcObjectState, state *dcFlowState, seen map[string]bool) []Finding {
	if object.Kind == dcUnknown {
		return nil
	}
	expr, ok := dcParseKeyExpr(keyText, state.Scalars)
	presence := 0
	if ok {
		presence = object.Keys[dcKeyIdentity(expr)]
	}
	if presence == 1 {
		return nil
	}
	severity := "information"
	message := receiver + " item/key existence cannot be proven before access."
	reason := "The key may be absent on one or more reachable paths."
	if presence == -1 || object.Empty == 1 {
		severity = "warning"
		message = receiver + " accesses or removes a key that is definitely absent."
		reason = "The tracked object is empty or the same key was removed on every reachable path."
	}
	finding := a.simpleFinding(file, proc, line, "VBA207", severity, message, reason, "Guard Dictionary access with Exists; for Collection, use a Count check or one narrow error probe followed immediately by On Error GoTo 0.")
	var out []Finding
	out = dcAppendFinding(out, seen, finding)
	return out
}

func dcObjectForReceiver(state *dcFlowState, receiver string) (string, dcObjectState) {
	key := strings.ToLower(cleanIdentifier(receiver))
	if id := state.Bindings[key]; id != "" {
		return id, state.Objects[id]
	}
	suffix := ":" + key
	match := ""
	var object dcObjectState
	for id, candidate := range state.Objects {
		if !strings.HasSuffix(id, suffix) || candidate.Kind == dcUnknown {
			continue
		}
		if match != "" {
			return "", dcObjectState{}
		}
		match, object = id, candidate
	}
	return match, object
}

func (a Analyzer) dcNormalizationFinding(file parsedFile, proc sourceProcedure, line int, receiver, keyText string, object dcObjectState, state *dcFlowState, seen map[string]bool) []Finding {
	expr, ok := dcParseKeyExpr(keyText, state.Scalars)
	if !ok || expr.Base == "" {
		return nil
	}
	previous, found := object.Normalizations[expr.Base]
	if !found || previous == expr.Norm || previous == "" && expr.Norm == "" {
		return nil
	}
	if object.CompareMode == "text" && dcCaseOnlyNormalization(previous) && dcCaseOnlyNormalization(expr.Norm) {
		return nil
	}
	finding := a.simpleFinding(file, proc, line, "VBA232", "warning", receiver+" uses inconsistent normalization for key "+expr.Base+".", "The same key source is used with both "+dcNormLabel(previous)+" and "+dcNormLabel(expr.Norm)+", which can create or miss distinct entries.", "Normalize the key once with a consistent LCase$/UCase$/Trim$ pipeline before Add, Exists, Item, and Remove.")
	var out []Finding
	out = dcAppendFinding(out, seen, finding)
	return out
}

func (s *dcFlowState) clone() *dcFlowState {
	if s == nil {
		return nil
	}
	out := &dcFlowState{Bindings: map[string]string{}, Objects: map[string]dcObjectState{}, Scalars: map[string]dcKeyExpr{}}
	for key, value := range s.Bindings {
		out.Bindings[key] = value
	}
	for key, value := range s.Objects {
		value.Keys = dcCloneIntMap(value.Keys)
		value.Normalizations = dcCloneStringMap(value.Normalizations)
		out.Objects[key] = value
	}
	for key, value := range s.Scalars {
		out.Scalars[key] = value
	}
	return out
}

func dcMergeState(current, incoming *dcFlowState) (*dcFlowState, bool) {
	if current == nil {
		return incoming.clone(), true
	}
	merged := &dcFlowState{Bindings: map[string]string{}, Objects: map[string]dcObjectState{}, Scalars: map[string]dcKeyExpr{}}
	for key, value := range current.Bindings {
		if incoming.Bindings[key] == value {
			merged.Bindings[key] = value
			continue
		}
		left, leftOK := current.Objects[value]
		rightID := incoming.Bindings[key]
		right, rightOK := incoming.Objects[rightID]
		if leftOK && !rightOK && left.Kind != dcUnknown {
			merged.Bindings[key] = value
			merged.Objects[value] = dcObjectState{Kind: left.Kind, LateBound: left.LateBound, Keys: map[string]int{}, Normalizations: map[string]string{}}
			continue
		}
		if leftOK && rightOK && left.Kind != dcUnknown && left.Kind == right.Kind {
			phi := "phi:" + key
			merged.Bindings[key] = phi
			merged.Objects[phi] = dcJoinObject(left, right)
		}
	}
	for key, value := range incoming.Bindings {
		if _, exists := current.Bindings[key]; exists {
			continue
		}
		if object, ok := incoming.Objects[value]; ok && object.Kind != dcUnknown {
			merged.Bindings[key] = value
			merged.Objects[value] = dcObjectState{Kind: object.Kind, LateBound: object.LateBound, Keys: map[string]int{}, Normalizations: map[string]string{}}
		}
	}
	for key, left := range current.Objects {
		right, ok := incoming.Objects[key]
		if !ok || left.Kind != right.Kind || left.LateBound != right.LateBound {
			continue
		}
		merged.Objects[key] = dcJoinObject(left, right)
	}
	for key, value := range current.Scalars {
		if incoming.Scalars[key] == value {
			merged.Scalars[key] = value
		}
	}
	if dcStatesEqual(current, merged) {
		return current, false
	}
	return merged, true
}

func dcJoinObject(left, right dcObjectState) dcObjectState {
	joined := dcObjectState{Kind: left.Kind, LateBound: left.LateBound && right.LateBound, Keys: map[string]int{}, Normalizations: map[string]string{}}
	if left.Empty == right.Empty {
		joined.Empty = left.Empty
	}
	if left.CompareMode == right.CompareMode {
		joined.CompareMode = left.CompareMode
	}
	for fact, value := range left.Keys {
		if right.Keys[fact] == value {
			joined.Keys[fact] = value
		}
	}
	for base, value := range left.Normalizations {
		if right.Normalizations[base] == value {
			joined.Normalizations[base] = value
		}
	}
	return joined
}

func dcStatesEqual(left, right *dcFlowState) bool {
	if len(left.Bindings) != len(right.Bindings) || len(left.Objects) != len(right.Objects) || len(left.Scalars) != len(right.Scalars) {
		return false
	}
	for key, value := range left.Bindings {
		if right.Bindings[key] != value {
			return false
		}
	}
	for key, value := range left.Scalars {
		if right.Scalars[key] != value {
			return false
		}
	}
	for key, leftObject := range left.Objects {
		rightObject, ok := right.Objects[key]
		if !ok || leftObject.Kind != rightObject.Kind || leftObject.LateBound != rightObject.LateBound || leftObject.Empty != rightObject.Empty || leftObject.CompareMode != rightObject.CompareMode || !dcIntMapsEqual(leftObject.Keys, rightObject.Keys) || !dcStringMapsEqual(leftObject.Normalizations, rightObject.Normalizations) {
			return false
		}
	}
	return true
}

func dcCloneIntMap(in map[string]int) map[string]int {
	out := map[string]int{}
	for key, value := range in {
		out[key] = value
	}
	return out
}

func dcCloneStringMap(in map[string]string) map[string]string {
	out := map[string]string{}
	for key, value := range in {
		out[key] = value
	}
	return out
}

func dcIntMapsEqual(left, right map[string]int) bool {
	if len(left) != len(right) {
		return false
	}
	for key, value := range left {
		if right[key] != value {
			return false
		}
	}
	return true
}

func dcStringMapsEqual(left, right map[string]string) bool {
	if len(left) != len(right) {
		return false
	}
	for key, value := range left {
		if right[key] != value {
			return false
		}
	}
	return true
}

func dcMemberCall(text string) (string, string, []string, bool) {
	masked := strings.TrimSpace(text)
	dot := strings.Index(masked, ".")
	if dot <= 0 {
		return "", "", nil, false
	}
	receiver := cleanIdentifier(strings.TrimSpace(masked[:dot]))
	if strings.Contains(receiver, " ") || receiver == "" {
		// Member calls commonly appear after an assignment. Use the final bare
		// identifier before the dot in that case.
		parts := strings.FieldsFunc(masked[:dot], func(r rune) bool { return r == ' ' || r == '=' || r == '(' || r == ')' })
		if len(parts) == 0 {
			return "", "", nil, false
		}
		receiver = cleanIdentifier(parts[len(parts)-1])
	}
	rest := strings.TrimSpace(masked[dot+1:])
	end := 0
	for end < len(rest) && (rest[end] == '_' || rest[end] >= '0' && rest[end] <= '9' || rest[end] >= 'A' && rest[end] <= 'Z' || rest[end] >= 'a' && rest[end] <= 'z') {
		end++
	}
	if end == 0 {
		return "", "", nil, false
	}
	member := rest[:end]
	rest = strings.TrimSpace(rest[end:])
	if strings.HasPrefix(rest, "=") {
		return receiver, member, []string{strings.TrimSpace(rest[1:])}, true
	}
	if strings.HasPrefix(rest, "(") {
		inside, ok := dcBalancedContent(rest)
		if !ok {
			return receiver, member, nil, false
		}
		return receiver, member, dcSplitArgs(inside), true
	}
	if rest == "" || strings.HasPrefix(rest, ":") {
		return receiver, member, nil, true
	}
	return receiver, member, dcSplitArgs(rest), true
}

func dcKnownMemberCall(text string, state *dcFlowState) (string, string, []string, bool) {
	lower := strings.ToLower(text)
	keys := make([]string, 0, len(state.Bindings))
	for key := range state.Bindings {
		keys = append(keys, key)
	}
	sort.SliceStable(keys, func(i, j int) bool { return len(keys[i]) > len(keys[j]) })
	for _, receiver := range keys {
		needle := receiver + "."
		at := strings.Index(lower, needle)
		if at < 0 {
			continue
		}
		candidate := text[at:]
		gotReceiver, member, args, ok := dcMemberCall(candidate)
		if ok && strings.EqualFold(gotReceiver, receiver) {
			return gotReceiver, member, args, true
		}
	}
	return "", "", nil, false
}

func dcDefaultAccess(text string) (string, string, bool) {
	lower := strings.ToLower(text)
	for i := 0; i < len(text); i++ {
		if text[i] != '(' {
			continue
		}
		start := i - 1
		for start >= 0 && (text[start] == '_' || text[start] >= '0' && text[start] <= '9' || text[start] >= 'A' && text[start] <= 'Z' || text[start] >= 'a' && text[start] <= 'z') {
			start--
		}
		receiver := cleanIdentifier(text[start+1 : i])
		if receiver == "" || start >= 0 && text[start] == '.' {
			continue
		}
		if strings.HasSuffix(lower[:i], "if ") || strings.EqualFold(receiver, "lcase") || strings.EqualFold(receiver, "ucase") || strings.EqualFold(receiver, "trim") || strings.EqualFold(receiver, "createobject") {
			continue
		}
		inside, ok := dcBalancedContent(text[i:])
		if ok {
			return receiver, dcFirstArgument(inside), true
		}
	}
	return "", "", false
}

func dcDefaultAssignment(text string) (string, string, bool) {
	left := strings.SplitN(text, "=", 2)
	if len(left) != 2 {
		return "", "", false
	}
	return dcDefaultAccess(left[0])
}

func dcBalancedContent(text string) (string, bool) {
	inside, _, ok := dcBalancedContentSpan(text)
	return inside, ok
}

func dcBalancedContentExact(text string) (string, bool) {
	inside, end, ok := dcBalancedContentSpan(text)
	return inside, ok && strings.TrimSpace(text[end:]) == ""
}

func dcBalancedContentSpan(text string) (string, int, bool) {
	if text == "" || text[0] != '(' {
		return "", 0, false
	}
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
			if !inString {
				depth--
				if depth == 0 {
					return text[1:i], i + 1, true
				}
			}
		}
	}
	return "", 0, false
}

func dcSplitArgs(text string) []string {
	var args []string
	start, depth := 0, 0
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
		case ',':
			if !inString && depth == 0 {
				args = append(args, strings.TrimSpace(text[start:i]))
				start = i + 1
			}
		}
	}
	if strings.TrimSpace(text[start:]) != "" || len(args) > 0 {
		args = append(args, strings.TrimSpace(text[start:]))
	}
	return args
}

func dcFirstArgument(text string) string {
	args := dcSplitArgs(strings.TrimSpace(text))
	if len(args) == 0 {
		return ""
	}
	return args[0]
}

func dcBareIdentifier(text string) bool {
	cleaned := cleanIdentifier(strings.TrimSpace(text))
	if cleaned == "" || len(cleaned) != len(strings.TrimSpace(text)) {
		return false
	}
	for i, r := range cleaned {
		if r != '_' && (r < 'a' || r > 'z') && (r < 'A' || r > 'Z') && (i == 0 || r < '0' || r > '9') {
			return false
		}
	}
	return true
}

func dcSimpleFunctionCall(text string) (string, []string, bool) {
	trimmed := strings.TrimSpace(text)
	at := strings.Index(trimmed, "(")
	if at <= 0 {
		return "", nil, false
	}
	name := strings.TrimSpace(trimmed[:at])
	if !dcBareIdentifier(lastName(name)) {
		return "", nil, false
	}
	inside, ok := dcBalancedContentExact(trimmed[at:])
	return name, dcSplitArgs(inside), ok
}

func dcExistsCondition(text string) (string, string, bool, bool) {
	condition, ok := dcBareCondition(text)
	if !ok {
		return "", "", false, false
	}
	negated := false
	if strings.HasPrefix(strings.ToLower(condition), "not ") {
		negated = true
		condition = strings.TrimSpace(condition[4:])
	}
	lower := strings.ToLower(condition)
	at := strings.Index(lower, ".exists")
	if at < 0 {
		return "", "", false, false
	}
	receiver := strings.TrimSpace(condition[:at])
	if !dcBareIdentifier(receiver) {
		return "", "", false, false
	}
	rest := strings.TrimSpace(condition[at+len(".exists"):])
	inside, ok := dcBalancedContentExact(rest)
	if !ok {
		return "", "", false, false
	}
	return receiver, dcFirstArgument(inside), negated, true
}

func dcBareCondition(text string) (string, bool) {
	trimmed := strings.TrimSpace(dcStatementSource(procedureir.Statement{Text: text}))
	lower := strings.ToLower(trimmed)
	if !strings.HasPrefix(lower, "if ") && !strings.HasPrefix(lower, "elseif ") {
		return "", false
	}
	condition := strings.TrimSpace(trimmed[strings.Index(lower, "if ")+len("if "):])
	if then := strings.Index(strings.ToLower(condition), " then"); then >= 0 {
		condition = strings.TrimSpace(condition[:then])
	}
	return condition, condition != ""
}

func dcCallTargetUnqualified(text string) bool {
	trimmed := strings.TrimSpace(text)
	if strings.HasPrefix(strings.ToLower(trimmed), "call ") {
		trimmed = strings.TrimSpace(trimmed[len("call "):])
	}
	end := len(trimmed)
	if stop := strings.IndexAny(trimmed, "( \t"); stop >= 0 {
		end = stop
	}
	target := strings.TrimSpace(trimmed[:end])
	return !strings.Contains(target, ".") && dcBareIdentifier(target)
}

func dcHelperCondition(text string) (string, []string, bool, bool) {
	trimmed := strings.TrimSpace(text)
	lower := strings.ToLower(trimmed)
	if !strings.HasPrefix(lower, "if ") && !strings.HasPrefix(lower, "elseif ") {
		return "", nil, false, false
	}
	condition := strings.TrimSpace(trimmed[strings.Index(lower, "if ")+2:])
	negated := false
	if strings.HasPrefix(strings.ToLower(condition), "not ") {
		negated = true
		condition = strings.TrimSpace(condition[4:])
	}
	if then := strings.Index(strings.ToLower(condition), " then"); then >= 0 {
		condition = strings.TrimSpace(condition[:then])
	}
	name, args, ok := dcSimpleFunctionCall(condition)
	return name, args, negated, ok
}

func dcParseKeyExpr(text string, aliases map[string]dcKeyExpr) (dcKeyExpr, bool) {
	trimmed := strings.TrimSpace(text)
	lower := strings.ToLower(trimmed)
	for _, fn := range []struct {
		names []string
		norm  string
	}{
		{[]string{"lcase", "lcase$"}, "lower"}, {[]string{"ucase", "ucase$"}, "upper"}, {[]string{"trim", "trim$"}, "trim"},
	} {
		for _, name := range fn.names {
			if !strings.HasPrefix(lower, name+"(") {
				continue
			}
			inside, ok := dcBalancedContent(trimmed[len(name):])
			if !ok {
				return dcKeyExpr{}, false
			}
			inner, ok := dcParseKeyExpr(inside, aliases)
			if !ok {
				return dcKeyExpr{}, false
			}
			if inner.Norm == "" {
				inner.Norm = fn.norm
			} else {
				inner.Norm += "+" + fn.norm
			}
			return inner, true
		}
	}
	if dcBareIdentifier(trimmed) {
		key := strings.ToLower(cleanIdentifier(trimmed))
		if alias, ok := aliases[key]; ok {
			return alias, true
		}
		return dcKeyExpr{Base: key}, true
	}
	if strings.HasPrefix(trimmed, "\"") && strings.HasSuffix(trimmed, "\"") {
		return dcKeyExpr{Base: strings.ToLower(trimmed)}, true
	}
	if _, err := strconv.Atoi(trimmed); err == nil {
		return dcKeyExpr{Base: trimmed}, true
	}
	return dcKeyExpr{}, false
}

func dcKeyIdentity(expr dcKeyExpr) string { return expr.Base + "|" + expr.Norm }

func dcObserveNormalization(object *dcObjectState, expr dcKeyExpr) {
	if expr.Base == "" {
		return
	}
	if _, exists := object.Normalizations[expr.Base]; !exists {
		object.Normalizations[expr.Base] = expr.Norm
	}
}

func dcCompareMode(text string) string {
	switch strings.ToLower(cleanIdentifier(strings.TrimSpace(text))) {
	case "vbtextcompare", "textcompare":
		return "text"
	case "vbbinarycompare", "binarycompare":
		return "binary"
	default:
		return ""
	}
}

var dcLateCompareConstants = map[string]bool{"binarycompare": true, "textcompare": true, "databasecompare": true, "comparemethod.binarycompare": true, "comparemethod.textcompare": true, "comparemethod.databasecompare": true}

// dcConstantNamesForProcedure is the source-only compatibility path used when
// a standalone parsedFile has no attached module facts. Module constants and
// constants in the current procedure are visible; locals owned by another
// procedure are excluded using the IR procedure ranges.
func dcConstantNamesForProcedure(file parsedFile, proc sourceProcedure) map[string]bool {
	out := map[string]bool{}
	forEachSourceConstantForProcedure(file, proc, func(constant moduleConstantFact) {
		out[strings.ToLower(strings.TrimSpace(constant.Name))] = true
	})
	return out
}

func dcNarrowProbeStatements(proc sourceProcedure) map[int]bool {
	out := map[int]bool{}
	for i, statement := range proc.Statements.AllIndexed() {
		if !isOnErrorResumeNext(statement) {
			continue
		}
		operations := 0
		var candidate int
		for j := i + 1; j < proc.Statements.Len(); j++ {
			next := proc.Statements.valueAt(j)
			if next.Kind == procedureir.StatementDeclaration || next.Kind == procedureir.StatementLabel || resumeNextScopeErrProbeStatement(next) {
				continue
			}
			if restoresErrorHandling(next) {
				if operations == 1 {
					out[candidate] = true
				}
				break
			}
			if resumeNextScopeControlStatement(next) {
				break
			}
			operations++
			candidate = next.ID
			if operations > 1 {
				break
			}
		}
	}
	return out
}

func dcStatementInRepeatedLoop(statement procedureir.Statement, loops []excelLoopRegion) bool {
	for _, loop := range containingExcelLoops(loops, statement.ID, statement.Range.StartLine) {
		if loop.Body[statement.ID] {
			return true
		}
	}
	return false
}

func dcLoopHeaderMaterialization(statement procedureir.Statement, loops []excelLoopRegion) bool {
	for _, loop := range loops {
		if loop.StatementID == statement.ID {
			return true
		}
	}
	return false
}

func dcForEach(text string) (string, string, bool) {
	if newline := strings.IndexAny(text, "\r\n"); newline >= 0 {
		text = text[:newline]
	}
	match := forEachDirectRe.FindStringSubmatch(strings.Join(strings.Fields(gui.StripComment(text)), " "))
	if len(match) == 0 {
		return "", "", false
	}
	return match[1], match[2], true
}

func dcMutatesEnumeratedCollection(statement procedureir.Statement, objectID string, state *dcFlowState, proc sourceProcedure, loops []excelLoopRegion) bool {
	statements := map[int]procedureir.Statement{}
	for candidate := range proc.Statements.All() {
		statements[candidate.ID] = candidate
	}
	for current := statement; current.ParentID != 0; {
		parent, ok := statements[current.ParentID]
		if !ok {
			break
		}
		_, receiver, forEach := dcForEach(parent.Text)
		if forEach {
			id, _ := dcObjectForReceiver(state, receiver)
			if id == objectID && dcMutationCanContinueIteration(proc, statement, parent.ID) {
				return true
			}
		}
		current = parent
	}
	for _, loop := range containingExcelLoops(loops, statement.ID, statement.Range.StartLine) {
		if !loop.Body[statement.ID] {
			continue
		}
		_, receiver, ok := dcForEach(statements[loop.StatementID].Text)
		if ok && state.Bindings[strings.ToLower(receiver)] == objectID && dcMutationCanContinueIteration(proc, statement, loop.StatementID) {
			return true
		}
	}
	return false
}

func dcMutationCanContinueIteration(proc sourceProcedure, mutation procedureir.Statement, loopStatementID int) bool {
	if proc.Graph == nil {
		return false
	}
	if dcMutationHasUnconditionalExit(proc, mutation, loopStatementID) {
		return false
	}
	start, ok := proc.Graph.BlockForStatement(mutation.ID)
	if !ok {
		return false
	}
	header, ok := proc.Graph.BlockForStatement(loopStatementID)
	if !ok {
		return false
	}
	edges := map[vbacfg.BlockID][]vbacfg.Edge{}
	for _, edge := range proc.Graph.Edges {
		if edge.Class == vbacfg.EdgeNormal && !edge.Uncertain {
			edges[edge.From] = append(edges[edge.From], edge)
		}
	}
	seen := map[vbacfg.BlockID]bool{start.ID: true}
	queue := []vbacfg.BlockID{start.ID}
	for len(queue) > 0 {
		from := queue[0]
		queue = queue[1:]
		for _, edge := range edges[from] {
			if edge.Kind == vbacfg.EdgeLoopBack && edge.To == header.ID {
				return true
			}
			if edge.To == proc.Graph.NormalExit || edge.To == proc.Graph.ExceptionalExit || edge.To == proc.Graph.TerminationExit || edge.To == proc.Graph.UnknownExit || seen[edge.To] {
				continue
			}
			seen[edge.To] = true
			queue = append(queue, edge.To)
		}
	}
	return false
}

func dcMutationHasUnconditionalExit(proc sourceProcedure, mutation procedureir.Statement, loopStatementID int) bool {
	byID := map[int]procedureir.Statement{}
	for statement := range proc.Statements.All() {
		byID[statement.ID] = statement
	}
	ancestors := map[int]bool{}
	for current := mutation; current.ID != 0 && current.ID != loopStatementID; current = byID[current.ParentID] {
		ancestors[current.ParentID] = true
		if current.ParentID == 0 {
			break
		}
	}
	for statement := range proc.Statements.All() {
		if statement.Range.StartByte <= mutation.Range.StartByte {
			continue
		}
		if !dcStatementDescendsFrom(statement, loopStatementID, byID) {
			continue
		}
		lower := strings.ToLower(strings.TrimSpace(dcStatementSource(statement)))
		if (strings.HasPrefix(lower, "exit sub") || strings.HasPrefix(lower, "exit function") || strings.HasPrefix(lower, "exit property")) && ancestors[statement.ParentID] {
			return true
		}
	}
	return false
}

func dcStatementDescendsFrom(statement procedureir.Statement, ancestor int, byID map[int]procedureir.Statement) bool {
	for current := statement; current.ParentID != 0; {
		if current.ParentID == ancestor {
			return true
		}
		next, ok := byID[current.ParentID]
		if !ok {
			return false
		}
		current = next
	}
	return false
}

func dcZeroIndex(arg string, statement procedureir.Statement, proc sourceProcedure) bool {
	trimmed := strings.ReplaceAll(strings.TrimSpace(arg), " ", "")
	if trimmed == "0" {
		return true
	}
	if !dcBareIdentifier(trimmed) {
		return false
	}
	statements := map[int]procedureir.Statement{}
	for candidate := range proc.Statements.All() {
		statements[candidate.ID] = candidate
	}
	for current := statement; current.ParentID != 0; {
		parent, ok := statements[current.ParentID]
		if !ok {
			break
		}
		header := strings.ToLower(strings.Join(strings.Fields(dcStatementSource(parent)), " "))
		compactHeader := strings.ReplaceAll(header, " ", "")
		if strings.HasPrefix(compactHeader, "for"+strings.ToLower(trimmed)+"=0to") {
			return true
		}
		lboundPrefix := "for" + strings.ToLower(trimmed) + "=lbound("
		if strings.HasPrefix(compactHeader, lboundPrefix) {
			arrayStart := len(lboundPrefix)
			if closeAt := strings.Index(compactHeader[arrayStart:], ")"); closeAt >= 0 {
				arrayName := cleanIdentifier(compactHeader[arrayStart : arrayStart+closeAt])
				if dcArrayHasExplicitZeroLowerBound(proc, arrayName, parent.Range.StartByte) {
					return true
				}
			}
		}
		current = parent
	}
	return false
}

var dcDefaultAccessPattern = regexp.MustCompile(`(?i)\b([A-Z_][A-Z0-9_]*)\s*\(([^()\r\n]*)\)`)

func dcKnownDefaultAccesses(text string, state *dcFlowState) [][2]string {
	var accesses [][2]string
	for _, match := range dcDefaultAccessPattern.FindAllStringSubmatch(text, -1) {
		if len(match) != 3 {
			continue
		}
		receiver := cleanIdentifier(match[1])
		if state.Bindings[strings.ToLower(receiver)] == "" {
			continue
		}
		accesses = append(accesses, [2]string{receiver, dcFirstArgument(match[2])})
	}
	return accesses
}

func dcArrayHasExplicitZeroLowerBound(proc sourceProcedure, arrayName string, beforeByte int) bool {
	if arrayName == "" {
		return false
	}
	needle := strings.ToLower(arrayName) + "(0 to "
	compactNeedle := strings.ReplaceAll(needle, " ", "")
	for statement := range proc.Statements.All() {
		if statement.Range.StartByte >= beforeByte {
			continue
		}
		text := strings.ToLower(strings.Join(strings.Fields(dcStatementSource(statement)), " "))
		compact := strings.ReplaceAll(text, " ", "")
		if (statement.Kind == procedureir.StatementDeclaration || statement.Kind == procedureir.StatementReDim) && strings.Contains(compact, compactNeedle) {
			return true
		}
	}
	return false
}

func dcCaseOnlyNormalization(norm string) bool {
	if norm == "" {
		return true
	}
	for _, part := range strings.Split(norm, "+") {
		if part != "lower" && part != "upper" {
			return false
		}
	}
	return true
}

func dcNormLabel(norm string) string {
	if norm == "" {
		return "the raw key"
	}
	return norm + " normalization"
}

func dcAppendFinding(findings []Finding, seen map[string]bool, finding Finding) []Finding {
	key := finding.Code + ":" + strconv.Itoa(finding.Line) + ":" + finding.Message
	if seen[key] {
		return findings
	}
	seen[key] = true
	return append(findings, finding)
}
