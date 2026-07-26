package lspserver

import (
	"path/filepath"
	"sort"
	"strings"

	"github.com/tliron/glsp"
	protocol "github.com/tliron/glsp/protocol_3_16"

	"github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/calls"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

// callHierarchySymbolID is the project-local identity transported in LSP item
// data. It deliberately excludes source positions so a prepared item remains
// usable after unrelated edits in the same procedure.
type callHierarchySymbolID struct {
	File   string `json:"file"`
	Module string `json:"module"`
	Name   string `json:"name"`
	Kind   string `json:"kind"`
}

type callHierarchyProcedure struct {
	ID     callHierarchySymbolID
	Symbol intel.Symbol
}

func (s *Server) prepareCallHierarchy(_ *glsp.Context, params *protocol.CallHierarchyPrepareParams) ([]protocol.CallHierarchyItem, error) {
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		return nil, err
	}
	if s.documentKind(doc) != DocumentKindVBA {
		return []protocol.CallHierarchyItem{}, nil
	}
	s.updateWorkspaceSymbolOverlay(doc)
	procedure, ok, err := s.analysis.callHierarchyProcedureAt(doc.Path, fromProtocolPosition(params.Position))
	if err != nil {
		return nil, err
	}
	if !ok {
		return []protocol.CallHierarchyItem{}, nil
	}
	return []protocol.CallHierarchyItem{s.callHierarchyItem(procedure)}, nil
}

func (s *Server) callHierarchyIncomingCalls(_ *glsp.Context, params *protocol.CallHierarchyIncomingCallsParams) ([]protocol.CallHierarchyIncomingCall, error) {
	target, ok, err := s.callHierarchyProcedureForItem(params.Item)
	if err != nil {
		return nil, err
	}
	if !ok {
		return []protocol.CallHierarchyIncomingCall{}, nil
	}

	resolved, err := s.analysis.queryResolvedCalls(workspaceCallQuery{CalleeBase: target.Symbol.Name})
	if err != nil {
		return nil, err
	}
	grouped := make(map[callHierarchySymbolID][]protocol.Range)
	procedures := make(map[callHierarchySymbolID]callHierarchyProcedure)
	for _, call := range resolved {
		if !callHierarchyMatchesTarget(call, target.ID) {
			continue
		}
		caller, found, err := s.analysis.callHierarchyProcedureForCaller(call)
		if err != nil {
			return nil, err
		}
		if !found {
			continue
		}
		grouped[caller.ID] = append(grouped[caller.ID], callHierarchyRange(call.Range))
		procedures[caller.ID] = caller
	}

	ids := sortedCallHierarchyIDs(grouped)
	out := make([]protocol.CallHierarchyIncomingCall, 0, len(ids))
	for _, id := range ids {
		out = append(out, protocol.CallHierarchyIncomingCall{
			From:       s.callHierarchyItem(procedures[id]),
			FromRanges: sortCallHierarchyRanges(grouped[id]),
		})
	}
	return out, nil
}

func (s *Server) callHierarchyOutgoingCalls(_ *glsp.Context, params *protocol.CallHierarchyOutgoingCallsParams) ([]protocol.CallHierarchyOutgoingCall, error) {
	caller, ok, err := s.callHierarchyProcedureForItem(params.Item)
	if err != nil {
		return nil, err
	}
	if !ok {
		return []protocol.CallHierarchyOutgoingCall{}, nil
	}

	resolved, err := s.analysis.queryResolvedCalls(workspaceCallQuery{
		Caller:     caller.Symbol.Module + "." + caller.Symbol.Name,
		CallerKind: caller.Symbol.Kind,
	})
	if err != nil {
		return nil, err
	}
	grouped := make(map[callHierarchySymbolID][]protocol.Range)
	procedures := make(map[callHierarchySymbolID]callHierarchyProcedure)
	for _, call := range resolved {
		actualCaller, found, err := s.analysis.callHierarchyProcedureForCaller(call)
		if err != nil {
			return nil, err
		}
		if !found || actualCaller.ID != caller.ID || call.Resolution.Status != "matched" || len(call.Resolution.Candidates) != 1 {
			continue
		}
		callee, found, err := s.analysis.callHierarchyProcedureForCandidate(call.Resolution.Candidates[0])
		if err != nil {
			return nil, err
		}
		if !found {
			continue
		}
		grouped[callee.ID] = append(grouped[callee.ID], callHierarchyRange(call.Range))
		procedures[callee.ID] = callee
	}

	ids := sortedCallHierarchyIDs(grouped)
	out := make([]protocol.CallHierarchyOutgoingCall, 0, len(ids))
	for _, id := range ids {
		out = append(out, protocol.CallHierarchyOutgoingCall{
			To:         s.callHierarchyItem(procedures[id]),
			FromRanges: sortCallHierarchyRanges(grouped[id]),
		})
	}
	return out, nil
}

func (s *Server) callHierarchyProcedureForItem(item protocol.CallHierarchyItem) (callHierarchyProcedure, bool, error) {
	id, ok := callHierarchyIDFromData(item.Data)
	if !ok {
		return callHierarchyProcedure{}, false, nil
	}
	return s.analysis.callHierarchyProcedureForID(id)
}

func (s *Server) callHierarchyItem(procedure callHierarchyProcedure) protocol.CallHierarchyItem {
	detail := procedure.Symbol.Detail
	if detail == "" {
		detail = procedure.Symbol.Module + "." + procedure.Symbol.Name
	}
	return protocol.CallHierarchyItem{
		Name:           procedure.Symbol.Name,
		Kind:           symbolKind(procedure.Symbol.Kind),
		Detail:         &detail,
		URI:            protocol.DocumentUri(s.docs.uriForDisplayPath(procedure.ID.File)),
		Range:          toProtocolRange(procedure.Symbol.Range),
		SelectionRange: toProtocolRange(procedure.Symbol.Selection),
		Data:           procedure.ID,
	}
}

func (x *workspaceAnalysisIndex) callHierarchyProcedureAt(path string, position intel.Position) (callHierarchyProcedure, bool, error) {
	if err := x.waitReady(); err != nil {
		return callHierarchyProcedure{}, false, err
	}
	key := symbolFileKey(path)
	x.mu.RLock()
	defer x.mu.RUnlock()
	entry, ok := x.effective[key]
	if !ok {
		return callHierarchyProcedure{}, false, nil
	}
	for _, symbol := range entry.symbols {
		if callHierarchyProcedureSymbol(symbol) && callHierarchyRangeContains(symbol.Range, position) {
			return callHierarchyProcedure{ID: callHierarchyIDForSymbol(x.root, entry.path, symbol), Symbol: symbol}, true, nil
		}
	}
	return callHierarchyProcedure{}, false, nil
}

func (x *workspaceAnalysisIndex) callHierarchyProcedureForID(id callHierarchySymbolID) (callHierarchyProcedure, bool, error) {
	if err := x.waitReady(); err != nil {
		return callHierarchyProcedure{}, false, err
	}
	id = normalizeCallHierarchyID(id)
	if !id.valid() {
		return callHierarchyProcedure{}, false, nil
	}
	x.mu.RLock()
	defer x.mu.RUnlock()
	for _, ref := range x.qualified[normalizeSymbolQuery(id.Module+"."+id.Name)] {
		symbol, ok := x.symbolForRefLocked(ref)
		if !ok || !callHierarchyProcedureSymbol(symbol) {
			continue
		}
		entry, ok := x.effective[ref.path]
		if !ok {
			continue
		}
		candidate := callHierarchyIDForSymbol(x.root, entry.path, symbol)
		if candidate == id {
			return callHierarchyProcedure{ID: candidate, Symbol: symbol}, true, nil
		}
	}
	return callHierarchyProcedure{}, false, nil
}

func (x *workspaceAnalysisIndex) callHierarchyProcedureForCaller(call calls.Call) (callHierarchyProcedure, bool, error) {
	if call.Caller == nil {
		return callHierarchyProcedure{}, false, nil
	}
	return x.callHierarchyProcedureForID(callHierarchySymbolID{
		File:   workspaceDisplayPath(x.root, call.File),
		Module: call.Module,
		Name:   call.Caller.Name,
		Kind:   call.Caller.Kind,
	})
}

func (x *workspaceAnalysisIndex) callHierarchyProcedureForCandidate(candidate calls.Candidate) (callHierarchyProcedure, bool, error) {
	module, name, ok := callHierarchyQualifiedName(candidate.QualifiedName)
	if !ok {
		return callHierarchyProcedure{}, false, nil
	}
	return x.callHierarchyProcedureForID(callHierarchySymbolID{
		File:   candidate.File,
		Module: module,
		Name:   name,
		Kind:   candidate.Kind,
	})
}

func callHierarchyMatchesTarget(call calls.Call, target callHierarchySymbolID) bool {
	if call.Resolution.Status != "matched" || len(call.Resolution.Candidates) != 1 {
		return false
	}
	module, name, ok := callHierarchyQualifiedName(call.Resolution.Candidates[0].QualifiedName)
	if !ok {
		return false
	}
	return normalizeCallHierarchyID(callHierarchySymbolID{
		File:   call.Resolution.Candidates[0].File,
		Module: module,
		Name:   name,
		Kind:   call.Resolution.Candidates[0].Kind,
	}) == target
}

func callHierarchyIDForSymbol(root, path string, symbol intel.Symbol) callHierarchySymbolID {
	return normalizeCallHierarchyID(callHierarchySymbolID{
		File:   workspaceDisplayPath(root, path),
		Module: symbol.Module,
		Name:   symbol.Name,
		Kind:   symbol.Kind,
	})
}

func callHierarchyIDFromData(data any) (callHierarchySymbolID, bool) {
	var id callHierarchySymbolID
	switch value := data.(type) {
	case callHierarchySymbolID:
		id = value
	case map[string]any:
		file, fileOK := value["file"].(string)
		module, moduleOK := value["module"].(string)
		name, nameOK := value["name"].(string)
		kind, kindOK := value["kind"].(string)
		if !fileOK || !moduleOK || !nameOK || !kindOK {
			return callHierarchySymbolID{}, false
		}
		id = callHierarchySymbolID{File: file, Module: module, Name: name, Kind: kind}
	default:
		return callHierarchySymbolID{}, false
	}
	id = normalizeCallHierarchyID(id)
	return id, id.valid()
}

func normalizeCallHierarchyID(id callHierarchySymbolID) callHierarchySymbolID {
	id.File = filepath.ToSlash(filepath.Clean(strings.TrimSpace(id.File)))
	id.Module = strings.ToLower(strings.TrimSpace(id.Module))
	id.Name = strings.ToLower(strings.TrimSpace(id.Name))
	id.Kind = strings.ToLower(strings.TrimSpace(id.Kind))
	return id
}

func (id callHierarchySymbolID) valid() bool {
	return id.File != "" && id.File != "." && id.Module != "" && id.Name != "" && callHierarchyKind(id.Kind)
}

func callHierarchyQualifiedName(value string) (string, string, bool) {
	value = strings.TrimSpace(value)
	index := strings.LastIndex(value, ".")
	if index <= 0 || index == len(value)-1 {
		return "", "", false
	}
	return value[:index], value[index+1:], true
}

func callHierarchyProcedureSymbol(symbol intel.Symbol) bool {
	return callHierarchyKind(symbol.Kind)
}

func callHierarchyKind(kind string) bool {
	switch strings.ToLower(strings.TrimSpace(kind)) {
	case "sub", "function", "property", "property_get", "property_let", "property_set":
		return true
	default:
		return false
	}
}

func callHierarchyRangeContains(r intel.Range, position intel.Position) bool {
	return compareIntelPosition(r.Start, position) <= 0 && compareIntelPosition(position, r.End) <= 0
}

func compareIntelPosition(left, right intel.Position) int {
	if left.Line != right.Line {
		if left.Line < right.Line {
			return -1
		}
		return 1
	}
	if left.Character < right.Character {
		return -1
	}
	if left.Character > right.Character {
		return 1
	}
	return 0
}

func callHierarchyRange(r ast.Range) protocol.Range {
	return protocol.Range{
		Start: protocol.Position{Line: protocol.UInteger(max(0, r.StartLine-1)), Character: protocol.UInteger(max(0, r.StartColumn-1))},
		End:   protocol.Position{Line: protocol.UInteger(max(0, r.EndLine-1)), Character: protocol.UInteger(max(0, r.EndColumn-1))},
	}
}

func sortedCallHierarchyIDs(grouped map[callHierarchySymbolID][]protocol.Range) []callHierarchySymbolID {
	ids := make([]callHierarchySymbolID, 0, len(grouped))
	for id := range grouped {
		ids = append(ids, id)
	}
	sort.Slice(ids, func(i, j int) bool {
		if ids[i].File != ids[j].File {
			return ids[i].File < ids[j].File
		}
		if ids[i].Module != ids[j].Module {
			return ids[i].Module < ids[j].Module
		}
		if ids[i].Name != ids[j].Name {
			return ids[i].Name < ids[j].Name
		}
		return ids[i].Kind < ids[j].Kind
	})
	return ids
}

func sortCallHierarchyRanges(ranges []protocol.Range) []protocol.Range {
	out := append([]protocol.Range(nil), ranges...)
	sort.Slice(out, func(i, j int) bool {
		if out[i].Start.Line != out[j].Start.Line {
			return out[i].Start.Line < out[j].Start.Line
		}
		if out[i].Start.Character != out[j].Start.Character {
			return out[i].Start.Character < out[j].Start.Character
		}
		if out[i].End.Line != out[j].End.Line {
			return out[i].End.Line < out[j].End.Line
		}
		return out[i].End.Character < out[j].End.Character
	})
	return out
}
