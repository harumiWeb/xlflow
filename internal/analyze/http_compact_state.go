package analyze

// This file contains the HTTP domain adapter for semanticstate.  The legacy
// httpAnalysisState remains available to the module-constant planner and to
// compatibility tests, but it never crosses a semanticstate solver edge.

import (
	"context"
	"fmt"
	"sort"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/analyze/semanticstate"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

type httpCompactSlotKind uint8

const (
	httpCompactObjectSlot httpCompactSlotKind = iota
	httpCompactLauncherSlot
	httpCompactValueSlot
	httpCompactSensitiveSlot
	httpCompactSavedSlot
)

// httpScalar is deliberately map-free.  Its strings are immutable values
// shared by the revision; Clone is therefore a plain value copy.
type httpScalar struct {
	class   httpCompactSlotKind
	present bool
	flag    bool
	text    string
	object  httpCompactObject
}

type httpCompactObject struct {
	present           bool
	kind              httpClientKind
	identity          string
	url               string
	urlKnown          bool
	hasCredentials    bool
	timeoutConfigured bool
	timeoutInfinite   bool
	downloaded        bool
}

type httpCompactSlot struct {
	id       semanticstate.SymbolID
	name     string
	variable string
	class    httpCompactSlotKind
	siteID   int
	path     string
}

type httpCompactEnvironment struct {
	variables []string
	saveSites []int
	names     []string

	objects   map[string]httpCompactSlot
	launchers map[string]httpCompactSlot
	values    map[string]httpCompactSlot
	sensitive map[string]httpCompactSlot
	saved     map[string][]httpCompactSlot
	savedPath map[string][]httpCompactSlot
	saveVars  map[string]bool
	savePaths []string

	initial       httpAnalysisState
	identityByKey map[string]string
	analyzer      Analyzer
}

func httpScalarZero(class httpCompactSlotKind) httpScalar {
	return httpScalar{class: class}
}

type httpScalarLattice struct{}

func (httpScalarLattice) Clone(value httpScalar) httpScalar { return value }

func (httpScalarLattice) Join(dst *httpScalar, src httpScalar) bool {
	if dst == nil {
		return false
	}
	if dst.class != src.class {
		return false
	}
	merged := *dst
	switch dst.class {
	case httpCompactObjectSlot:
		if !dst.object.present || !src.object.present {
			merged.object = httpCompactObject{}
			break
		}
		object := dst.object
		if object.kind != src.object.kind {
			object.kind = httpUnknown
		}
		if object.identity != src.object.identity {
			object.identity = ""
		}
		object.urlKnown = object.urlKnown && src.object.urlKnown && object.url == src.object.url
		if !object.urlKnown {
			object.url = ""
		}
		object.hasCredentials = object.hasCredentials || src.object.hasCredentials
		object.timeoutConfigured = object.timeoutConfigured && src.object.timeoutConfigured
		object.timeoutInfinite = object.timeoutInfinite || src.object.timeoutInfinite
		object.downloaded = object.downloaded && src.object.downloaded
		merged.object = object
	case httpCompactLauncherSlot, httpCompactValueSlot:
		if !dst.present || !src.present || dst.text != src.text {
			merged.present = false
			merged.text = ""
		}
	case httpCompactSensitiveSlot:
		merged.present = dst.present || src.present
		merged.flag = (dst.present && dst.flag) || (src.present && src.flag)
	case httpCompactSavedSlot:
		merged.present = dst.present && src.present && dst.text == src.text
		if !merged.present {
			merged.text = ""
		}
	}
	if *dst == merged {
		return false
	}
	*dst = merged
	return true
}

func buildHTTPCompactEnvironment(file parsedFile, proc sourceProcedure, initial httpAnalysisState) (*httpCompactEnvironment, semanticstate.Environment) {
	variables := map[string]bool{}
	addVariable := func(name string) {
		name = strings.ToLower(strings.TrimSpace(name))
		if name != "" {
			variables[name] = true
		}
	}
	addIdentifiers := func(text string) {
		for _, match := range httpIdentifierRe.FindAllString(text, -1) {
			addVariable(match)
		}
	}
	relevantStatement := func(text string) bool {
		text = strings.TrimSpace(stripVBAFileComment(text))
		if text == "" {
			return false
		}
		return httpObjectAssignmentRe.MatchString(text) ||
			httpMemberCallRe.MatchString(text) ||
			httpOptionAssignmentRe.MatchString(text) ||
			httpSetOptionCallRe.MatchString(text) ||
			httpLogRe.MatchString(text) ||
			httpObjectLauncherRe.MatchString(text) ||
			httpWin32LauncherRe.MatchString(text) ||
			strings.HasPrefix(strings.ToLower(text), "shell ")
	}

	// Start with declarations that are already proven HTTP objects and names
	// used by HTTP statements. Do not index every identifier in the module:
	// that turns a compact per-procedure environment into a dense copy of the
	// whole source file on large modules.
	for name := range initial.objects {
		addVariable(name)
	}
	for name := range initial.launchers {
		addVariable(name)
	}
	for _, statement := range proc.IR.Statements {
		if relevantStatement(statement.Text) {
			addIdentifiers(statement.Text)
		}
	}
	// Save-path slots are needed only for variables that can be an ADO stream
	// or that are an explicit SaveToFile receiver. Avoid allocating the full
	// variable-by-site cross product for unrelated identifiers in a large
	// procedure while remaining conservative for late-bound receivers.
	saveVars := map[string]bool{}
	for name, object := range initial.objects {
		if object.kind == httpADOStream {
			saveVars[name] = true
		}
	}
	for _, statement := range proc.IR.Statements {
		text := strings.TrimSpace(stripVBAFileComment(statement.Text))
		if match := httpObjectAssignmentRe.FindStringSubmatch(text); len(match) == 3 {
			target, rhs := strings.ToLower(match[1]), strings.TrimSpace(match[2])
			if httpKindFromConstruction(rhs) == httpADOStream {
				saveVars[target] = true
			}
		}
		if match := httpMemberCallRe.FindStringSubmatchIndex(text); match != nil && strings.EqualFold(text[match[4]:match[5]], "SaveToFile") {
			saveVars[strings.ToLower(text[match[2]:match[3]])] = true
		}
	}
	saveChanged := true
	for saveChanged {
		saveChanged = false
		for _, statement := range proc.IR.Statements {
			text := strings.TrimSpace(stripVBAFileComment(statement.Text))
			match := httpObjectAssignmentRe.FindStringSubmatch(text)
			if len(match) != 3 {
				continue
			}
			target, source := strings.ToLower(match[1]), strings.ToLower(strings.TrimSpace(match[2]))
			if saveVars[source] && !saveVars[target] {
				saveVars[target] = true
				saveChanged = true
			}
		}
	}

	// Local constants and aliases can be assigned outside the HTTP statement
	// itself (for example, url = BaseURL & "/download"). Pull in only the
	// transitive assignment dependencies of names already referenced above.
	changed := true
	for changed {
		changed = false
		for _, statement := range proc.IR.Statements {
			text := strings.TrimSpace(stripVBAFileComment(statement.Text))
			name, expression, ok := fileAssignment(text)
			if !ok || !variables[strings.ToLower(name)] {
				continue
			}
			before := len(variables)
			addIdentifiers(expression)
			if len(variables) != before {
				changed = true
			}
		}
	}

	// Module constants are immutable facts, but only constants reachable from
	// the HTTP expression set need indexed slots. Their expression dependencies
	// are added by the same fixed-point scan below.
	moduleConstants := map[string]string{}
	if facts := file.ModuleFacts; facts != nil {
		facts.forEachConstant(func(constant moduleConstantFact) {
			moduleConstants[strings.ToLower(strings.TrimSpace(constant.Name))] = constant.Expression
		})
	}
	changed = true
	for changed {
		changed = false
		for name, expression := range moduleConstants {
			if !variables[name] {
				continue
			}
			before := len(variables)
			addIdentifiers(expression)
			if len(variables) != before {
				changed = true
			}
		}
	}
	for name := range initial.strings {
		if variables[name] {
			addVariable(name)
		}
	}
	for name := range initial.known {
		if variables[name] {
			addVariable(name)
		}
	}
	for name := range initial.sensitive {
		if variables[name] {
			addVariable(name)
		}
	}

	vars := make([]string, 0, len(variables))
	for name := range variables {
		vars = append(vars, name)
	}
	sort.Strings(vars)
	sites := map[int]bool{}
	paths := map[string]bool{}
	identityByKey := map[string]string{}
	staticStrings := httpCompactProcedureStringCandidates(proc, initial)
	for _, statement := range proc.IR.Statements {
		text := strings.TrimSpace(stripVBAFileComment(statement.Text))
		if match := httpMemberCallRe.FindStringSubmatchIndex(text); match != nil && strings.EqualFold(text[match[4]:match[5]], "SaveToFile") {
			sites[statement.ID] = true
			args := httpMethodArgs(text, match[1])
			if len(args) > 0 {
				for _, path := range httpCompactStringCandidates(args[0], staticStrings) {
					if httpExecutableExtRe.MatchString(strings.ToLower(path)) {
						paths[httpPathKey(path)] = true
					}
				}
			}
		}
		if match := httpObjectAssignmentRe.FindStringSubmatch(text); len(match) == 3 {
			if httpKindFromConstruction(strings.TrimSpace(match[2])) != httpUnknown {
				identityByKey[strings.ToLower(match[1])+fmt.Sprintf("|%d", statement.ID)] = fmt.Sprintf("%s@%d", strings.ToLower(match[1]), statement.ID)
			}
		}
	}
	saveSites := make([]int, 0, len(sites))
	for site := range sites {
		saveSites = append(saveSites, site)
	}
	sort.Ints(saveSites)
	savePaths := make([]string, 0, len(paths))
	for path := range paths {
		savePaths = append(savePaths, path)
	}
	sort.Strings(savePaths)

	names := make([]string, 0, len(vars)*4+len(saveVars)*(len(saveSites)+len(savePaths)))
	for _, variable := range vars {
		names = append(names, "object:"+variable, "launcher:"+variable, "value:"+variable, "sensitive:"+variable)
		if saveVars[variable] {
			for _, site := range saveSites {
				names = append(names, fmt.Sprintf("saved:%s:%d", variable, site))
			}
			for _, path := range savePaths {
				names = append(names, "saved-path:"+variable+":"+path)
			}
		}
	}
	sort.Strings(names)
	return &httpCompactEnvironment{
		variables: vars, saveSites: saveSites, names: names, initial: initial,
		objects: map[string]httpCompactSlot{}, launchers: map[string]httpCompactSlot{},
		values: map[string]httpCompactSlot{}, sensitive: map[string]httpCompactSlot{},
		saved: map[string][]httpCompactSlot{}, savedPath: map[string][]httpCompactSlot{}, saveVars: saveVars, savePaths: savePaths, identityByKey: identityByKey,
	}, semanticstate.NewEnvironment(names, names)
}

// httpCompactProcedureStringCandidates is an allocation-bounded, conservative
// pre-scan used only to allocate saved-path symbols.  It gathers literal and
// module-constant values assigned to procedure-local names, retaining all
// values seen across branches so that a later SaveToFile site has a stable
// path slot even when its path is not available in the entry state.
func httpCompactProcedureStringCandidates(proc sourceProcedure, initial httpAnalysisState) map[string]map[string]bool {
	values := map[string]map[string]bool{}
	for name, value := range initial.strings {
		if !initial.known[name] {
			continue
		}
		key := strings.ToLower(strings.TrimSpace(name))
		if key == "" {
			continue
		}
		values[key] = map[string]bool{value: true}
	}
	for pass := 0; pass <= len(proc.IR.Statements); pass++ {
		changed := false
		for _, statement := range proc.IR.Statements {
			text := strings.TrimSpace(stripVBAFileComment(statement.Text))
			name, expression, ok := fileAssignment(text)
			if !ok || strings.HasPrefix(strings.ToLower(text), "set ") {
				continue
			}
			candidates := httpCompactStringCandidates(expression, values)
			if len(candidates) == 0 {
				continue
			}
			key := strings.ToLower(strings.TrimSpace(name))
			if key == "" {
				continue
			}
			destination := values[key]
			if destination == nil {
				destination = map[string]bool{}
				values[key] = destination
			}
			for _, candidate := range candidates {
				if !destination[candidate] {
					destination[candidate] = true
					changed = true
				}
			}
		}
		if !changed {
			break
		}
	}
	return values
}

func httpCompactStringCandidates(expr string, values map[string]map[string]bool) []string {
	parts := splitHTTPConcat(strings.TrimSpace(expr))
	if len(parts) == 0 || (len(parts) == 1 && strings.TrimSpace(parts[0]) == "") {
		return nil
	}
	candidates := []string{""}
	const maxCandidates = 64
	for _, rawPart := range parts {
		part := strings.TrimSpace(rawPart)
		var pieces []string
		if len(part) >= 2 && part[0] == '"' && part[len(part)-1] == '"' {
			pieces = []string{strings.ReplaceAll(part[1:len(part)-1], `""`, `"`)}
		} else {
			key := strings.ToLower(strings.TrimRight(part, "$%&@!#^"))
			for value := range values[key] {
				pieces = append(pieces, value)
			}
			sort.Strings(pieces)
		}
		if len(pieces) == 0 {
			return nil
		}
		next := make([]string, 0, len(candidates)*len(pieces))
		for _, prefix := range candidates {
			for _, piece := range pieces {
				if len(next) == maxCandidates {
					break
				}
				next = append(next, prefix+piece)
			}
			if len(next) == maxCandidates {
				break
			}
		}
		candidates = next
	}
	return candidates
}

func (e *httpCompactEnvironment) bind(environment semanticstate.Environment) {
	for _, variable := range e.variables {
		objectName := "object:" + variable
		launcherName := "launcher:" + variable
		valueName := "value:" + variable
		sensitiveName := "sensitive:" + variable
		objectID, _ := environment.Symbol(objectName)
		launcherID, _ := environment.Symbol(launcherName)
		valueID, _ := environment.Symbol(valueName)
		sensitiveID, _ := environment.Symbol(sensitiveName)
		e.objects[variable] = httpCompactSlot{id: objectID, name: objectName, variable: variable, class: httpCompactObjectSlot}
		e.launchers[variable] = httpCompactSlot{id: launcherID, name: launcherName, variable: variable, class: httpCompactLauncherSlot}
		e.values[variable] = httpCompactSlot{id: valueID, name: valueName, variable: variable, class: httpCompactValueSlot}
		e.sensitive[variable] = httpCompactSlot{id: sensitiveID, name: sensitiveName, variable: variable, class: httpCompactSensitiveSlot}
		if !e.saveVars[variable] {
			continue
		}
		for _, site := range e.saveSites {
			name := fmt.Sprintf("saved:%s:%d", variable, site)
			id, _ := environment.Symbol(name)
			e.saved[variable] = append(e.saved[variable], httpCompactSlot{id: id, name: name, variable: variable, class: httpCompactSavedSlot, siteID: site})
		}
		for _, path := range e.savePaths {
			name := "saved-path:" + variable + ":" + path
			id, _ := environment.Symbol(name)
			e.savedPath[variable] = append(e.savedPath[variable], httpCompactSlot{id: id, name: name, variable: variable, class: httpCompactSavedSlot, path: path})
		}
	}
}

func (e *httpCompactEnvironment) seed(state *semanticstate.State[httpScalar]) {
	for _, variable := range e.variables {
		state.Set(e.objects[variable].id, httpScalarZero(httpCompactObjectSlot))
		state.Set(e.launchers[variable].id, httpScalarZero(httpCompactLauncherSlot))
		state.Set(e.values[variable].id, httpScalarZero(httpCompactValueSlot))
		state.Set(e.sensitive[variable].id, httpScalarZero(httpCompactSensitiveSlot))
		for _, slot := range e.saved[variable] {
			state.Set(slot.id, httpScalarZero(httpCompactSavedSlot))
		}
		for _, slot := range e.savedPath[variable] {
			state.Set(slot.id, httpScalarZero(httpCompactSavedSlot))
		}
	}
	for name, object := range e.initial.objects {
		if slot, ok := e.objects[name]; ok {
			state.Set(slot.id, httpScalar{class: httpCompactObjectSlot, object: httpCompactObject{present: object.kind != httpUnknown, kind: object.kind, identity: object.identity, url: object.url, urlKnown: object.urlKnown, hasCredentials: object.hasCredentials, timeoutConfigured: object.timeoutConfigured, timeoutInfinite: object.timeoutInfinite, downloaded: object.downloaded}})
		}
	}
	for name, value := range e.initial.strings {
		if e.initial.known[name] {
			if slot, ok := e.values[name]; ok {
				state.Set(slot.id, httpScalar{class: httpCompactValueSlot, present: true, text: value})
			}
		}
	}
	for name, value := range e.initial.launchers {
		if slot, ok := e.launchers[name]; ok {
			state.Set(slot.id, httpScalar{class: httpCompactLauncherSlot, present: value != "", text: value})
		}
	}
	for name, value := range e.initial.sensitive {
		if slot, ok := e.sensitive[name]; ok {
			state.Set(slot.id, httpScalar{class: httpCompactSensitiveSlot, present: true, flag: value})
		}
	}
}

func httpCompactValue(view semanticstate.StateView[httpScalar], slot httpCompactSlot) httpScalar {
	value, ok := view.Value(slot.id)
	if !ok {
		return httpScalarZero(slot.class)
	}
	return value
}

func httpCompactSet(state *semanticstate.State[httpScalar], slot httpCompactSlot, value httpScalar) {
	state.Set(slot.id, value)
}

func (e *httpCompactEnvironment) object(view semanticstate.StateView[httpScalar], name string) httpCompactObject {
	if slot, ok := e.objects[name]; ok {
		return httpCompactValue(view, slot).object
	}
	return httpCompactObject{}
}

func (e *httpCompactEnvironment) copySaved(view semanticstate.StateView[httpScalar], out *semanticstate.State[httpScalar], from, to string) {
	for index, slot := range e.saved[from] {
		if index < len(e.saved[to]) {
			httpCompactSet(out, e.saved[to][index], httpCompactValue(view, slot))
		}
	}
	for index, slot := range e.savedPath[from] {
		if index < len(e.savedPath[to]) {
			httpCompactSet(out, e.savedPath[to][index], httpCompactValue(view, slot))
		}
	}
}

func (e *httpCompactEnvironment) clearSaved(out *semanticstate.State[httpScalar], variable string) {
	for _, slot := range e.saved[variable] {
		httpCompactSet(out, slot, httpScalarZero(httpCompactSavedSlot))
	}
	for _, slot := range e.savedPath[variable] {
		httpCompactSet(out, slot, httpScalarZero(httpCompactSavedSlot))
	}
}

func (e *httpCompactEnvironment) setObjectAliases(view semanticstate.StateView[httpScalar], out *semanticstate.State[httpScalar], receiver string, object httpCompactObject) {
	if slot, ok := e.objects[receiver]; ok {
		httpCompactSet(out, slot, httpScalar{class: httpCompactObjectSlot, object: object})
	}
	if object.identity == "" {
		return
	}
	for _, name := range e.variables {
		if name == receiver {
			continue
		}
		candidate := e.object(view, name)
		if candidate.present && candidate.identity == object.identity {
			if slot, ok := e.objects[name]; ok {
				httpCompactSet(out, slot, httpScalar{class: httpCompactObjectSlot, object: object})
			}
			for index, saved := range e.saved[receiver] {
				if index < len(e.saved[name]) {
					httpCompactSet(out, e.saved[name][index], httpCompactValue(view, saved))
				}
			}
			for index, saved := range e.savedPath[receiver] {
				if index < len(e.savedPath[name]) {
					httpCompactSet(out, e.savedPath[name][index], httpCompactValue(view, saved))
				}
			}
		}
	}
}

func (e *httpCompactEnvironment) constructorIdentity(variable string, statementID int) string {
	return e.identityByKey[variable+fmt.Sprintf("|%d", statementID)]
}

func httpCompactConstantString(expr string, view semanticstate.StateView[httpScalar], e *httpCompactEnvironment) (string, bool) {
	parts := splitHTTPConcat(strings.TrimSpace(expr))
	if len(parts) == 0 || (len(parts) == 1 && strings.TrimSpace(parts[0]) == "") {
		return "", false
	}
	var builder strings.Builder
	for _, part := range parts {
		part = strings.TrimSpace(part)
		if len(part) >= 2 && part[0] == '"' && part[len(part)-1] == '"' {
			builder.WriteString(strings.ReplaceAll(part[1:len(part)-1], `""`, `"`))
			continue
		}
		key := strings.ToLower(strings.TrimRight(part, "$%&@!#^"))
		slot, ok := e.values[key]
		if !ok {
			return "", false
		}
		value := httpCompactValue(view, slot)
		if !value.present {
			return "", false
		}
		builder.WriteString(value.text)
	}
	return builder.String(), true
}

func httpCompactInteger(expr string, view semanticstate.StateView[httpScalar], e *httpCompactEnvironment) (int, bool) {
	text := strings.TrimSpace(expr)
	if slot, ok := e.values[strings.ToLower(text)]; ok {
		value := httpCompactValue(view, slot)
		if value.present {
			text = value.text
		}
	}
	text = strings.ReplaceAll(strings.ReplaceAll(text, "&H", "0x"), "&h", "0x")
	text = strings.TrimRight(text, "%&@!#$")
	n, err := strconv.ParseInt(text, 0, strconv.IntSize)
	return int(n), err == nil
}

func httpCompactSensitive(expr string, view semanticstate.StateView[httpScalar], e *httpCompactEnvironment) bool {
	if httpAuthLiteralRe.MatchString(expr) {
		return true
	}
	lower := strings.ToLower(expr)
	for name, slot := range e.sensitive {
		value := httpCompactValue(view, slot)
		if value.present && value.flag && containsHTTPIdentifierToken(lower, name) {
			return true
		}
	}
	return false
}

func httpCompactMarkSensitive(expr string, out *semanticstate.State[httpScalar], view semanticstate.StateView[httpScalar], e *httpCompactEnvironment) {
	for _, match := range httpIdentifierRe.FindAllStringIndex(expr, -1) {
		name := strings.ToLower(expr[match[0]:match[1]])
		rest := strings.TrimLeft(expr[match[1]:], " \t")
		if strings.HasPrefix(rest, "(") {
			continue
		}
		if slot, ok := e.sensitive[name]; ok {
			httpCompactSet(out, slot, httpScalar{class: httpCompactSensitiveSlot, present: true, flag: true})
		}
	}
}

func httpCompactTLSRisk(kind httpClientKind, option, value string, view semanticstate.StateView[httpScalar], e *httpCompactEnvironment) string {
	lower := strings.ToLower(option)
	n, known := httpCompactInteger(value, view, e)
	if kind == httpServerXML && (lower == "2" || strings.Contains(lower, "ignore_server_ssl_cert")) && known && n != 0 {
		return "certificate_validation_bypass"
	}
	if kind != httpWinHTTP {
		return ""
	}
	if (lower == "4" || strings.Contains(lower, "sslerrorignoreflags")) && known && n != 0 {
		return "certificate_validation_bypass"
	}
	if lower == "18" || strings.Contains(lower, "enablecertificaterevocationcheck") {
		if (known && n == 0) || strings.EqualFold(strings.TrimSpace(value), "false") {
			return "certificate_validation_bypass"
		}
	}
	if lower == "9" || strings.Contains(lower, "secureprotocols") {
		if known && n&(0x8|0x20|0x80|0x200) != 0 {
			return "obsolete_tls_protocol"
		}
	}
	return ""
}

func httpCompactResponseExpression(expr string, view semanticstate.StateView[httpScalar], e *httpCompactEnvironment) bool {
	lower := strings.ToLower(expr)
	if !strings.Contains(lower, "responsebody") && !strings.Contains(lower, "responsestream") {
		return false
	}
	for name, slot := range e.objects {
		object := httpCompactValue(view, slot).object
		if object.present && (object.kind == httpXML || object.kind == httpServerXML || object.kind == httpWinHTTP) && containsHTTPIdentifierToken(lower, name) {
			return true
		}
	}
	return false
}

type httpCompactSolve struct {
	index       *semanticstate.Index
	env         *httpCompactEnvironment
	environment semanticstate.Environment
	result      semanticstate.Result[httpScalar]
	blocks      map[vbacfg.BlockID]semanticstate.BlockOrdinal
	reachable   map[vbacfg.BlockID]bool
	statements  map[int]vbacfg.BlockID
	evidence    map[vbacfg.BlockID]httpEvidenceState
	projection  semanticstate.State[httpScalar]
}

type httpCompactInput struct {
	view  semanticstate.StateView[httpScalar]
	valid bool
}

func (s *httpCompactSolve) entryState(id vbacfg.BlockID) httpCompactInput {
	if !s.reachable[id] {
		return httpCompactInput{}
	}
	ordinal, ok := s.blocks[id]
	if !ok {
		return httpCompactInput{}
	}
	return httpCompactInput{view: s.result.State(ordinal, 0), valid: true}
}

func (s *httpCompactSolve) collect(file parsedFile, proc sourceProcedure, statement procedureir.Statement, view semanticstate.StateView[httpScalar]) []httpFindingSpec {
	// Projection runs are ordered by CFG block and never overlap. Reuse one
	// scalar scratch state instead of allocating a complete indexed state for
	// every diagnostic statement after fixed-point convergence.
	s.projection.CloneFrom(view, httpScalarLattice{}.Clone)
	var evidence httpEvidenceState
	if block, ok := s.statements[statement.ID]; ok {
		evidence = s.evidence[block]
	}
	return httpCompactTransfer(file, proc, statement, s.env, view, &s.projection, true, &evidence)
}

func solveHTTPCompactStates(ctx context.Context, a Analyzer, file parsedFile, proc sourceProcedure, graph vbacfg.Graph, initial httpAnalysisState) (*httpCompactSolve, error) {
	view := graph.View(vbacfg.EdgeFilter{})
	index, err := semanticstate.NewIndexView(view)
	if err != nil {
		return nil, err
	}
	env, environment := buildHTTPCompactEnvironment(file, proc, initial)
	env.analyzer = a
	env.bind(environment)
	solver, err := semanticstate.NewSolver(index, environment, httpScalarLattice{}, []semanticstate.Lane[httpScalar]{{
		Initialize: func(_ context.Context, _ semanticstate.LaneOrdinal, state *semanticstate.State[httpScalar]) error {
			env.seed(state)
			return nil
		},
		Transfer: func(_ context.Context, _ semanticstate.LaneOrdinal, block semanticstate.BlockOrdinal, input semanticstate.StateView[httpScalar], output *semanticstate.State[httpScalar]) error {
			output.CloneFrom(input, httpScalarLattice{}.Clone)
			candidate, ok := view.BlockAtOrdinal(vbacfg.BlockOrdinal(block))
			if !ok || candidate.Statement == nil || candidate.Statement.Recovered {
				return nil
			}
			httpCompactTransfer(file, proc, *candidate.Statement, env, input, output, false, nil)
			return nil
		},
	}})
	if err != nil {
		return nil, err
	}
	result, err := solver.SolveContext(ctx)
	if err != nil {
		return nil, err
	}
	blocks := make(map[vbacfg.BlockID]semanticstate.BlockOrdinal, index.BlockCount())
	for _, block := range index.Blocks() {
		blocks[block.ID] = block.Ordinal
	}
	reachable := make(map[vbacfg.BlockID]bool)
	for _, id := range graph.Reachable(vbacfg.EdgeFilter{}) {
		reachable[id] = true
	}
	statements := make(map[int]vbacfg.BlockID)
	for _, block := range graph.Blocks {
		if block.Statement != nil {
			statements[block.Statement.ID] = block.ID
		}
	}
	base := &httpCompactSolve{index: index, env: env, environment: environment, result: result, blocks: blocks, reachable: reachable, statements: statements}
	evidence, err := solveHTTPEvidence(ctx, file, proc, graph, initial, base)
	if err != nil {
		return nil, err
	}
	base.evidence = evidence
	base.projection = semanticstate.NewState[httpScalar](environment.Layout())
	return base, nil
}

func httpCompactTransfer(file parsedFile, proc sourceProcedure, statement procedureir.Statement, e *httpCompactEnvironment, input semanticstate.StateView[httpScalar], output *semanticstate.State[httpScalar], collect bool, evidence *httpEvidenceState) []httpFindingSpec {
	text := strings.TrimSpace(stripVBAFileComment(statement.Text))
	if text == "" {
		return nil
	}
	line := statement.Range.StartLine
	if line <= 0 {
		line = proc.StartLine
	}
	var findings []httpFindingSpec

	if match := httpObjectAssignmentRe.FindStringSubmatch(text); len(match) == 3 {
		target, rhs := strings.ToLower(match[1]), strings.TrimSpace(match[2])
		previousObject := e.object(input, target)
		if objectSlot, ok := e.objects[target]; ok {
			httpCompactSet(output, objectSlot, httpScalarZero(httpCompactObjectSlot))
		}
		if launcherSlot, ok := e.launchers[target]; ok {
			httpCompactSet(output, launcherSlot, httpScalarZero(httpCompactLauncherSlot))
		}
		sourceName := strings.ToLower(strings.TrimSpace(rhs))
		sourceObject := e.object(input, sourceName)
		sourceLauncher := httpScalarZero(httpCompactLauncherSlot)
		if launcherSlot, ok := e.launchers[sourceName]; ok {
			sourceLauncher = httpCompactValue(input, launcherSlot)
		}
		switch {
		case httpKindFromConstruction(rhs) != httpUnknown:
			kind := httpKindFromConstruction(rhs)
			identity := e.constructorIdentity(target, statement.ID)
			if identity == "" {
				identity = fmt.Sprintf("%s@%d", target, statement.ID)
			}
			if objectSlot, ok := e.objects[target]; ok {
				httpCompactSet(output, objectSlot, httpScalar{class: httpCompactObjectSlot, object: httpCompactObject{present: true, kind: kind, identity: identity}})
			}
			e.clearSaved(output, target)
		case sourceObject.present && sourceObject.kind != httpUnknown:
			source := sourceName
			object := sourceObject
			if object.identity == "" {
				object.identity = source
			}
			if slot, ok := e.objects[target]; ok {
				httpCompactSet(output, slot, httpScalar{class: httpCompactObjectSlot, object: object})
			}
			// Read the RHS from the incoming block state.  In particular, a
			// self-assignment must not observe the target after its destination
			// slots were cleared above.
			e.copySaved(input, output, source, target)
		case httpLauncherFromConstruction(rhs) != "":
			if slot, ok := e.launchers[target]; ok {
				httpCompactSet(output, slot, httpScalar{class: httpCompactLauncherSlot, present: true, text: httpLauncherFromConstruction(rhs)})
			}
			e.clearSaved(output, target)
		case sourceLauncher.present:
			launcher := sourceLauncher
			if slot, ok := e.launchers[target]; ok {
				httpCompactSet(output, slot, launcher)
			}
			e.clearSaved(output, target)
		case httpPartialNewMatches(previousObject.kind, rhs):
			if slot, ok := e.objects[target]; ok {
				httpCompactSet(output, slot, httpScalar{class: httpCompactObjectSlot, object: previousObject})
			}
		case !httpPartialNewMatches(previousObject.kind, rhs):
			// The target was already cleared above, matching the legacy unknown
			// reassignment behavior.
			e.clearSaved(output, target)
		}
	}

	if name, expr, ok := fileAssignment(text); ok && !strings.HasPrefix(strings.ToLower(text), "set ") {
		key := strings.ToLower(name)
		if slot, exists := e.values[key]; exists {
			if value, known := httpCompactConstantString(expr, output.View(), e); known {
				httpCompactSet(output, slot, httpScalar{class: httpCompactValueSlot, present: true, text: value})
			} else {
				httpCompactSet(output, slot, httpScalarZero(httpCompactValueSlot))
			}
		}
		if slot, exists := e.sensitive[key]; exists {
			httpCompactSet(output, slot, httpScalar{class: httpCompactSensitiveSlot, present: true, flag: httpCompactSensitive(expr, output.View(), e)})
		}
	}

	if match := httpOptionAssignmentRe.FindStringSubmatch(text); len(match) == 5 {
		receiver := strings.ToLower(match[1])
		object := e.object(output.View(), receiver)
		option := strings.TrimSpace(match[2])
		if option == "" {
			option = strings.TrimSpace(match[3])
		}
		value := strings.TrimSpace(match[4])
		if object.present && (object.kind == httpWinHTTP || object.kind == httpServerXML) {
			if risk := httpCompactTLSRisk(object.kind, option, value, output.View(), e); risk != "" && collect {
				findings = append(findings, httpFindingSpec{line: line, column: strings.Index(strings.ToLower(text), "option"), code: "VBA246", api: string(object.kind), risk: risk, redact: true})
			}
		}
	}
	if match := httpSetOptionCallRe.FindStringSubmatchIndex(text); match != nil {
		receiver := strings.ToLower(text[match[2]:match[3]])
		object := e.object(output.View(), receiver)
		args := httpMethodArgs(text, match[1])
		if object.present && (object.kind == httpServerXML || object.kind == httpWinHTTP) && len(args) >= 2 {
			if risk := httpCompactTLSRisk(object.kind, args[0], args[1], output.View(), e); risk != "" && collect {
				findings = append(findings, httpFindingSpec{line: line, column: strings.Index(strings.ToLower(text), "setoption"), code: "VBA246", api: string(object.kind), risk: risk})
			}
		}
	}
	if httpLogRe.MatchString(text) && httpCompactSensitive(text, output.View(), e) && collect {
		findings = append(findings, httpFindingSpec{line: line, code: "VBA246", risk: "authorization_logging", redact: true})
	}
	if collect && httpCompactIsProcessLaunch(text, output.View(), e) {
		for _, arg := range httpLaunchArgs(text) {
			path, known := httpCompactConstantString(arg, output.View(), e)
			if !known {
				continue
			}
			launchKey := httpPathKey(path)
			for name, objectSlot := range e.objects {
				object := httpCompactValue(output.View(), objectSlot).object
				if !object.present || object.kind != httpADOStream {
					continue
				}
				matched := false
				for _, savedSlot := range e.saved[name] {
					saved := httpCompactValue(output.View(), savedSlot)
					if saved.present && (launchKey == saved.text || httpCommandContainsPath(launchKey, saved.text)) {
						matched = true
						break
					}
				}
				if !matched {
					for _, savedSlot := range e.savedPath[name] {
						saved := httpCompactValue(output.View(), savedSlot)
						if saved.present && (launchKey == savedSlot.path || httpCommandContainsPath(launchKey, savedSlot.path)) {
							matched = true
							break
						}
					}
				}
				if matched {
					findings = append(findings, httpFindingSpec{line: line, code: "VBA246", api: "ADODB.Stream", risk: "download_and_execute", redact: true})
				}
			}
		}
	}

	match := httpMemberCallRe.FindStringSubmatchIndex(text)
	if match == nil {
		return findings
	}
	receiver := strings.ToLower(text[match[2]:match[3]])
	method := strings.ToLower(text[match[4]:match[5]])
	object := e.object(output.View(), receiver)
	if !object.present {
		return findings
	}
	args := httpMethodArgs(text, match[1])
	switch method {
	case "open":
		if object.kind != httpXML && object.kind != httpServerXML && object.kind != httpWinHTTP {
			break
		}
		if len(args) > 1 {
			object.url, object.urlKnown = httpCompactConstantString(args[1], output.View(), e)
			if object.urlKnown && httpURLHasCredentials(object.url) && httpIsWebURL(object.url) && collect {
				findings = append(findings, httpFindingSpec{line: line, code: "VBA246", api: string(object.kind), risk: "credentials_in_url", origin: httpRedactedOrigin(object.url), redact: true})
			}
		}
		if len(args) > 3 && !httpCompactKnownEmpty(args[3], output.View(), e) || len(args) > 4 && !httpCompactKnownEmpty(args[4], output.View(), e) {
			object.hasCredentials = true
		}
		e.setObjectAliases(output.View(), output, receiver, object)
	case "setrequestheader":
		if object.kind == httpXML || object.kind == httpServerXML || object.kind == httpWinHTTP {
			header, headerKnown := "", false
			if len(args) > 0 {
				header, headerKnown = httpCompactConstantString(args[0], output.View(), e)
			}
			if headerKnown && httpSensitiveHeaderRe.MatchString(strings.TrimSpace(header)) {
				object.hasCredentials = true
				sensitiveModuleConstant := len(args) > 1 && httpCompactUsesSensitiveModuleConstant(args[1], file, output.View(), e)
				if len(args) > 1 {
					httpCompactMarkSensitive(args[1], output, output.View(), e)
				}
				if collect && sensitiveModuleConstant {
					findings = append(findings, httpFindingSpec{line: line, code: "VBA246", api: string(object.kind), risk: "sensitive_module_constant", header: canonicalHTTPHeader(header), redact: true})
				}
			}
			e.setObjectAliases(output.View(), output, receiver, object)
		}
	case "setcredentials":
		if object.kind == httpWinHTTP {
			object.hasCredentials = true
			e.setObjectAliases(output.View(), output, receiver, object)
		}
	case "settimeouts":
		if object.kind == httpWinHTTP || object.kind == httpServerXML {
			object.timeoutConfigured = len(args) == 4
			object.timeoutInfinite = false
			for _, arg := range args {
				if n, ok := httpCompactInteger(arg, output.View(), e); ok && (n == 0 || n == -1) {
					object.timeoutInfinite = true
					object.timeoutConfigured = false
				}
			}
			e.setObjectAliases(output.View(), output, receiver, object)
		}
	case "send":
		if object.kind != httpXML && object.kind != httpServerXML && object.kind != httpWinHTTP {
			break
		}
		if collect && object.urlKnown && strings.EqualFold(httpURLScheme(object.url), "http") && object.hasCredentials && !e.analyzer.isDevelopmentHTTPURL(object.url) {
			owned := map[int]bool{}
			if evidence != nil {
				owned = cloneHTTPIntSet(evidence.sinks[receiver])
			}
			findings = append(findings, httpFindingSpec{line: line, code: "VBA246", api: string(object.kind), risk: "plain_http_credentials", origin: httpRedactedOrigin(object.url), redact: true, ownedSinks: owned})
		}
		if collect && (object.kind == httpServerXML || object.kind == httpWinHTTP) && !object.timeoutConfigured {
			timeout := "missing"
			if object.timeoutInfinite {
				timeout = "unbounded"
			}
			findings = append(findings, httpFindingSpec{line: line, code: "VBA247", api: string(object.kind), risk: "missing_http_timeout", timeout: timeout})
		}
	case "write":
		if object.kind == httpADOStream && len(args) > 0 {
			object.downloaded = httpCompactResponseExpression(args[0], output.View(), e)
			e.setObjectAliases(output.View(), output, receiver, object)
			// Write replaces the stream contents, so any executable path
			// witness attached to the receiver (and its known aliases) is
			// invalidated.  An uncertain identity must not accidentally clear
			// unrelated variables whose scalar identity is also empty.
			e.clearSaved(output, receiver)
			if object.identity != "" {
				for _, name := range e.variables {
					if name == receiver {
						continue
					}
					candidate := e.object(output.View(), name)
					if candidate.present && candidate.identity == object.identity {
						e.clearSaved(output, name)
					}
				}
			}
		}
	case "savetofile":
		if object.kind == httpADOStream && object.downloaded && len(args) > 0 {
			if path, known := httpCompactConstantString(args[0], output.View(), e); known && httpExecutableExtRe.MatchString(strings.ToLower(path)) {
				site := statement.ID
				for _, slot := range e.saved[receiver] {
					if slot.siteID == site {
						httpCompactSet(output, slot, httpScalar{class: httpCompactSavedSlot, present: true, text: httpPathKey(path)})
					}
				}
				canonicalPath := httpPathKey(path)
				for _, slot := range e.savedPath[receiver] {
					if slot.path == canonicalPath {
						httpCompactSet(output, slot, httpScalar{class: httpCompactSavedSlot, present: true, text: canonicalPath})
					}
				}
				// SaveToFile mutates the stream identity. Propagate the indexed
				// witness to every live alias just as object member mutation
				// propagates the semantic object facts.
				e.setObjectAliases(output.View(), output, receiver, object)
			}
		}
	}
	return findings
}

func httpCompactKnownEmpty(expr string, view semanticstate.StateView[httpScalar], e *httpCompactEnvironment) bool {
	value, known := httpCompactConstantString(expr, view, e)
	return known && value == ""
}

func httpCompactUsesSensitiveModuleConstant(expr string, file parsedFile, view semanticstate.StateView[httpScalar], e *httpCompactEnvironment) bool {
	if facts := file.ModuleFacts; facts != nil {
		found := false
		facts.forEachConstant(func(constant moduleConstantFact) {
			if found || !constant.Module {
				return
			}
			name := strings.ToLower(constant.Name)
			if slot, ok := e.sensitive[name]; ok {
				value := httpCompactValue(view, slot)
				if value.present && value.flag && containsHTTPIdentifierToken(strings.ToLower(expr), name) {
					found = true
				}
			}
		})
		return found
	}
	for _, declaration := range file.IR.Declarations {
		name := strings.ToLower(declaration.Name)
		if declaration.Scope == procedureir.ScopeModule && declaration.Kind == "const" && containsHTTPIdentifierToken(strings.ToLower(expr), name) {
			if slot, ok := e.sensitive[name]; ok {
				value := httpCompactValue(view, slot)
				if value.present && value.flag {
					return true
				}
			}
		}
	}
	return false
}

func httpCompactIsProcessLaunch(text string, view semanticstate.StateView[httpScalar], e *httpCompactEnvironment) bool {
	lower := strings.ToLower(strings.TrimSpace(text))
	if strings.HasPrefix(lower, "shell ") || httpWin32LauncherRe.MatchString(lower) {
		return true
	}
	match := httpObjectLauncherRe.FindStringSubmatchIndex(text)
	if match == nil {
		return false
	}
	// The match starts at the member-call dot. Scan backwards from that
	// position so assignment-form calls such as `Set result = shell.Run ...`
	// resolve the launcher receiver rather than the statement prefix.
	end := match[0]
	for end > 0 && (text[end-1] == ' ' || text[end-1] == '\t') {
		end--
	}
	start := end
	for start > 0 && isVBAIdentifierPart(text[start-1]) {
		start--
	}
	if start == end {
		return false
	}
	receiver := strings.ToLower(text[start:end])
	matchedCall := strings.ToLower(text[match[0]:match[1]])
	method := ""
	for _, candidate := range []string{"shellexecute", "exec", "run"} {
		if strings.Contains(matchedCall, candidate) {
			method = candidate
			break
		}
	}
	launcherSlot, ok := e.launchers[receiver]
	if !ok {
		return false
	}
	launcher := httpCompactValue(view, launcherSlot)
	return launcher.present && ((launcher.text == "wscript.shell" && (method == "run" || method == "exec")) || (launcher.text == "shell.application" && method == "shellexecute"))
}

// Evidence is intentionally separate from semanticstate.State. It only
// carries source offsets needed to suppress duplicate VBA224 findings.
type httpEvidenceState struct {
	objects map[string]string
	unknown map[string]bool
	sinks   map[string]map[int]bool
}

func cloneHTTPEvidence(in httpEvidenceState) httpEvidenceState {
	out := httpEvidenceState{objects: map[string]string{}, unknown: map[string]bool{}, sinks: map[string]map[int]bool{}}
	for name, identity := range in.objects {
		out.objects[name] = identity
	}
	for name, unknown := range in.unknown {
		out.unknown[name] = unknown
	}
	for name, values := range in.sinks {
		out.sinks[name] = cloneHTTPIntSet(values)
	}
	return out
}

func joinHTTPEvidence(left, right httpEvidenceState, initialized bool) (httpEvidenceState, bool) {
	if !initialized {
		return cloneHTTPEvidence(right), true
	}
	out := cloneHTTPEvidence(left)
	changed := false
	for name, identity := range left.objects {
		rightIdentity, present := right.objects[name]
		if !present {
			delete(out.objects, name)
			delete(out.unknown, name)
			delete(out.sinks, name)
			changed = true
			continue
		}
		if left.unknown[name] || right.unknown[name] || identity != rightIdentity {
			if out.objects[name] != "" || !out.unknown[name] {
				out.objects[name] = ""
				out.unknown[name] = true
				changed = true
			}
		}
		merged := cloneHTTPIntSet(left.sinks[name])
		for sink := range right.sinks[name] {
			merged[sink] = true
		}
		if !sameHTTPIntSet(merged, left.sinks[name]) {
			out.sinks[name] = merged
			changed = true
		}
	}
	for name := range out.objects {
		if _, ok := right.objects[name]; !ok {
			delete(out.objects, name)
			delete(out.unknown, name)
			delete(out.sinks, name)
			changed = true
		}
	}
	return out, changed
}

func solveHTTPEvidence(ctx context.Context, file parsedFile, proc sourceProcedure, graph vbacfg.Graph, initial httpAnalysisState, solve *httpCompactSolve) (map[vbacfg.BlockID]httpEvidenceState, error) {
	states := map[vbacfg.BlockID]httpEvidenceState{graph.Entry: {objects: map[string]string{}, unknown: map[string]bool{}, sinks: map[string]map[int]bool{}}}
	for name, object := range initial.objects {
		if object.identity != "" {
			states[graph.Entry].objects[name] = object.identity
		}
	}
	queue := []vbacfg.BlockID{graph.Entry}
	queued := map[vbacfg.BlockID]bool{graph.Entry: true}
	for len(queue) > 0 {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		id := queue[0]
		queue = queue[1:]
		queued[id] = false
		state := cloneHTTPEvidence(states[id])
		block, ok := graph.BlockByID(id)
		if ok && block.Statement != nil && !block.Statement.Recovered {
			ordinal, found := solve.blocks[id]
			if found {
				input := solve.result.State(ordinal, 0)
				transferHTTPEvidence(*block.Statement, &state, input, solve.env)
			}
		}
		for _, edge := range graph.OutgoingEdges(id) {
			candidate := state
			if edge.Class == vbacfg.EdgeExceptional || edge.Uncertain {
				candidate = states[id]
			}
			previous, initialized := states[edge.To]
			merged, changed := joinHTTPEvidence(previous, candidate, initialized)
			if !changed {
				continue
			}
			states[edge.To] = merged
			if !queued[edge.To] {
				queue = append(queue, edge.To)
				queued[edge.To] = true
			}
		}
	}
	return states, nil
}

func transferHTTPEvidence(statement procedureir.Statement, state *httpEvidenceState, view semanticstate.StateView[httpScalar], e *httpCompactEnvironment) {
	text := strings.TrimSpace(stripVBAFileComment(statement.Text))
	if text == "" {
		return
	}
	if match := httpObjectAssignmentRe.FindStringSubmatch(text); len(match) == 3 {
		target, rhs := strings.ToLower(match[1]), strings.TrimSpace(match[2])
		switch {
		case httpKindFromConstruction(rhs) != httpUnknown:
			identity := e.constructorIdentity(target, statement.ID)
			if identity == "" {
				identity = fmt.Sprintf("%s@%d", target, statement.ID)
			}
			state.objects[target] = identity
			delete(state.unknown, target)
			delete(state.sinks, target)
		case state.objects[strings.ToLower(rhs)] != "" || state.unknown[strings.ToLower(rhs)]:
			source := strings.ToLower(rhs)
			state.objects[target] = state.objects[source]
			state.unknown[target] = state.unknown[source]
			state.sinks[target] = cloneHTTPIntSet(state.sinks[source])
		case httpPartialNewMatches(e.object(view, target).kind, rhs):
			// The IR may expose a typed `New ADODB`/`New MSXML2` prefix
			// while the declaration carries the complete dotted type. Preserve
			// the declaration identity and its sink witness in that case.
		default:
			delete(state.objects, target)
			delete(state.unknown, target)
			delete(state.sinks, target)
		}
	}
	match := httpMemberCallRe.FindStringSubmatchIndex(text)
	if match == nil {
		return
	}
	receiver := strings.ToLower(text[match[2]:match[3]])
	identity, present := state.objects[receiver]
	if !present {
		return
	}
	args := httpMethodArgs(text, match[1])
	method := strings.ToLower(text[match[4]:match[5]])
	credential := false
	switch method {
	case "open":
		credential = len(args) > 3 && !httpCompactKnownEmpty(args[3], view, e) || len(args) > 4 && !httpCompactKnownEmpty(args[4], view, e)
	case "setcredentials":
		credential = true
	case "setrequestheader":
		if len(args) > 0 {
			header, known := httpCompactConstantString(args[0], view, e)
			credential = known && httpSensitiveHeaderRe.MatchString(strings.TrimSpace(header))
		}
	}
	if !credential {
		return
	}
	if state.unknown[receiver] || identity == "" {
		if state.sinks[receiver] == nil {
			state.sinks[receiver] = map[int]bool{}
		}
		state.sinks[receiver][statement.Range.StartByte] = true
		return
	}
	for name, candidate := range state.objects {
		if !state.unknown[name] && candidate == identity {
			if state.sinks[name] == nil {
				state.sinks[name] = map[int]bool{}
			}
			state.sinks[name][statement.Range.StartByte] = true
		}
	}
}
