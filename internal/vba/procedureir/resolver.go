package procedureir

import (
	"path/filepath"
	"sort"
	"strings"
)

// Resolve returns a deep copy with project-dependent call and symbol
// resolutions applied. Syntax-local parameter/local/module scopes are kept.
func Resolve(in DocumentIR, resolver Resolver) DocumentIR {
	out := Clone(in)
	if resolver == nil {
		return out
	}
	for procedureIndex := range out.Procedures {
		procedure := &out.Procedures[procedureIndex]
		for callIndex := range procedure.Calls {
			call := &procedure.Calls[callIndex]
			call.Resolution = cloneCallResolution(resolver.ResolveCall(cloneCall(*call)))
		}
		for accessIndex := range procedure.Accesses {
			access := &procedure.Accesses[accessIndex]
			if access.Scope != ScopeUnresolved && access.Scope != ScopeProject {
				access.Resolution = SymbolResolution{Scope: access.Scope}
				continue
			}
			resolution := resolver.ResolveSymbol(SymbolReference{
				Name: access.Name, Module: out.ModuleName,
				Caller: ProcedureRef{
					Name: procedure.Symbol.Name, Kind: procedure.Symbol.Kind,
					QualifiedName: procedure.Symbol.QualifiedName,
				},
				Range: access.Range,
			})
			access.Resolution = cloneSymbolResolution(resolution)
			access.Scope = resolution.Scope
			if access.Scope == "" {
				access.Scope = ScopeUnresolved
				access.Resolution.Scope = ScopeUnresolved
			}
		}
	}
	return out
}

func cloneCallResolution(in CallResolution) CallResolution {
	in.Candidates = append([]Candidate(nil), in.Candidates...)
	return in
}

func cloneSymbolResolution(in SymbolResolution) SymbolResolution {
	in.Candidates = append([]Candidate(nil), in.Candidates...)
	return in
}

type resolverEntry struct {
	Candidate
	module     string
	moduleKind string
	visibility string
}

// SymbolResolver is a deterministic protocol-neutral project resolver.
type SymbolResolver struct {
	byName map[string][]resolverEntry
}

func NewSymbolResolver(symbols []ResolverSymbol) SymbolResolver {
	out := SymbolResolver{byName: map[string][]resolverEntry{}}
	for _, symbol := range symbols {
		if strings.TrimSpace(symbol.Name) == "" {
			continue
		}
		qualified := symbol.Name
		if symbol.Module != "" {
			qualified = symbol.Module + "." + symbol.Name
		}
		entry := resolverEntry{
			Candidate: Candidate{
				QualifiedName: qualified, Kind: symbol.Kind,
				File: normalizeCandidateFile(symbol.File), Line: symbol.Line,
			},
			module: symbol.Module, moduleKind: symbol.ModuleKind, visibility: symbol.Visibility,
		}
		key := strings.ToLower(cleanIdentifier(symbol.Name))
		out.byName[key] = append(out.byName[key], entry)
	}
	for key := range out.byName {
		sort.Slice(out.byName[key], func(i, j int) bool {
			a, b := out.byName[key][i], out.byName[key][j]
			if a.QualifiedName != b.QualifiedName {
				return a.QualifiedName < b.QualifiedName
			}
			if a.Kind != b.Kind {
				return a.Kind < b.Kind
			}
			if a.File != b.File {
				return a.File < b.File
			}
			return a.Line < b.Line
		})
	}
	return out
}

// NewResolver creates the default resolver backed by project symbols.
func NewResolver(symbols []ResolverSymbol) SymbolResolver {
	return NewSymbolResolver(symbols)
}

func (r SymbolResolver) ResolveCall(site CallSite) CallResolution {
	base := cleanIdentifier(strings.TrimPrefix(site.Callee.BaseName, "New "))
	caller := callerModule(site.Caller)
	entries := r.visible(r.byName[strings.ToLower(base)], caller)
	procedures := make([]resolverEntry, 0, len(entries))
	for _, entry := range entries {
		if isProcedureSymbolKind(entry.Kind) && (site.Callee.Receiver != nil || isReceiverlessProcedureCandidate(entry, caller)) {
			procedures = append(procedures, entry)
		}
	}
	if site.Callee.Receiver != nil {
		receiver := cleanQualifiedName(*site.Callee.Receiver)
		if isExternalLikeReceiver(receiver) {
			return CallResolution{Status: ResolutionExternal}
		}
		matches := receiverMatches(procedures, receiver, base)
		switch len(matches) {
		case 1:
			return CallResolution{Status: ResolutionMatched, Candidates: entriesToCandidates(matches)}
		case 0:
			return CallResolution{Status: ResolutionMemberCall}
		default:
			return CallResolution{Status: ResolutionAmbiguous, Candidates: entriesToCandidates(matches)}
		}
	}
	switch len(procedures) {
	case 1:
		return CallResolution{Status: ResolutionMatched, Candidates: entriesToCandidates(procedures)}
	default:
		if len(procedures) > 1 {
			return CallResolution{Status: ResolutionAmbiguous, Candidates: entriesToCandidates(procedures)}
		}
	}
	textKey := strings.ToLower(strings.TrimPrefix(site.Callee.Text, "New "))
	if builtinLikeNames[textKey] || builtinLikeNames[strings.ToLower(base)] {
		return CallResolution{Status: ResolutionBuiltinLike}
	}
	return CallResolution{Status: ResolutionUnresolved}
}

func isReceiverlessProcedureCandidate(entry resolverEntry, callerModule string) bool {
	if strings.EqualFold(entry.module, callerModule) {
		return true
	}
	return entry.moduleKind == "" || strings.EqualFold(entry.moduleKind, "standard")
}

func (r SymbolResolver) ResolveSymbol(ref SymbolReference) SymbolResolution {
	entries := r.visible(r.byName[strings.ToLower(cleanIdentifier(ref.Name))], callerModule(ref.Caller))
	var candidates []Candidate
	for _, entry := range entries {
		if isProcedureSymbolKind(entry.Kind) || strings.EqualFold(entry.Kind, "module") {
			continue
		}
		candidates = append(candidates, entry.Candidate)
	}
	if len(candidates) == 0 {
		return SymbolResolution{Scope: ScopeUnresolved}
	}
	return SymbolResolution{Scope: ScopeProject, Candidates: candidates}
}

func (r SymbolResolver) visible(entries []resolverEntry, callerModule string) []resolverEntry {
	out := make([]resolverEntry, 0, len(entries))
	for _, entry := range entries {
		if symbolIsPrivate(entry) && !strings.EqualFold(entry.module, callerModule) {
			continue
		}
		out = append(out, entry)
	}
	return out
}

func symbolIsPrivate(entry resolverEntry) bool {
	if strings.EqualFold(entry.visibility, "private") {
		return true
	}
	if strings.TrimSpace(entry.visibility) != "" {
		return false
	}
	switch strings.ToLower(entry.Kind) {
	case "variable", "module_variable", "field", "withevents_field", "const", "constant":
		return true
	default:
		return false
	}
}

func callerModule(caller ProcedureRef) string {
	if index := strings.LastIndex(strings.TrimSpace(caller.QualifiedName), "."); index > 0 {
		return caller.QualifiedName[:index]
	}
	return ""
}

func receiverMatches(entries []resolverEntry, receiver, base string) []resolverEntry {
	qualified := strings.ToLower(receiver + "." + base)
	shortQualified := strings.ToLower(cleanIdentifier(lastNamePart(receiver)) + "." + base)
	out := make([]resolverEntry, 0, len(entries))
	for _, entry := range entries {
		name := strings.ToLower(entry.QualifiedName)
		if name == qualified || name == shortQualified {
			out = append(out, entry)
		}
	}
	return out
}

func entriesToCandidates(entries []resolverEntry) []Candidate {
	out := make([]Candidate, len(entries))
	for i := range entries {
		out[i] = entries[i].Candidate
	}
	return out
}

func isProcedureSymbolKind(kind string) bool {
	switch strings.ToLower(kind) {
	case "sub", "function", "property", "property_get", "property_let", "property_set",
		"declare", "declare_sub", "declare_function", "event":
		return true
	default:
		return false
	}
}

func cleanQualifiedName(text string) string {
	parts := strings.FieldsFunc(strings.TrimSpace(text), func(r rune) bool { return r == '.' || r == '!' })
	cleaned := make([]string, 0, len(parts))
	for _, part := range parts {
		if part = cleanIdentifier(part); part != "" {
			cleaned = append(cleaned, part)
		}
	}
	return strings.Join(cleaned, ".")
}

func isExternalLikeReceiver(receiver string) bool {
	for _, part := range strings.Split(receiver, ".") {
		switch strings.ToLower(cleanIdentifier(part)) {
		case "application", "debug", "excel", "worksheetfunction":
			return true
		}
	}
	return false
}

func normalizeCandidateFile(file string) string {
	if strings.TrimSpace(file) == "" {
		return ""
	}
	return filepath.ToSlash(filepath.Clean(file))
}

var builtinLikeNames = map[string]bool{
	"array": true, "asc": true, "cbool": true, "cbyte": true, "ccur": true,
	"cdate": true, "cdbl": true, "cdec": true, "choose": true, "chr": true,
	"cint": true, "clng": true, "clnglng": true, "clngptr": true, "cos": true,
	"createobject": true, "cstr": true, "date": true, "dateadd": true,
	"debug.print": true, "dir": true, "doevents": true, "environ": true, "error": true,
	"format": true, "getobject": true, "inputbox": true, "instr": true,
	"isarray": true, "isdate": true, "isempty": true, "iserror": true,
	"isnull": true, "isnumeric": true, "join": true, "lbound": true,
	"lcase": true, "left": true, "len": true, "mid": true, "msgbox": true,
	"replace": true, "right": true, "rnd": true, "shell": true, "split": true, "str": true,
	"trim": true, "typename": true, "ubound": true, "ucase": true, "val": true,
}
