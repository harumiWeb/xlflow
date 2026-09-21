package analyze

import (
	"errors"
	"os"
	pathpkg "path"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
	"sync"

	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/effects"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/userforms"
)

type eventFindingCandidate struct {
	line           int
	statementID    int
	boundary       string
	category       string
	classification string
	uncertainty    string
	effect         effects.Evidence
	same           bool
}

var userFormWithEventsFieldRE = regexp.MustCompile(`(?i)^\s*(?:(?:public|private|friend|dim|static)\s+)*with\s*events\s+([A-Za-z_][A-Za-z0-9_]*)\b`)
var userFormWithEventsTokenRE = regexp.MustCompile(`(?i)\bwith\s*events\b`)

func vbaCodeWithoutStringsAndComments(line string) string {
	var code strings.Builder
	for index := 0; index < len(line); index++ {
		switch line[index] {
		case '\'':
			return code.String()
		case '"':
			index++
			for index < len(line) {
				if line[index] != '"' {
					index++
					continue
				}
				if index+1 < len(line) && line[index+1] == '"' {
					index += 2
					continue
				}
				break
			}
			code.WriteByte(' ')
		default:
			code.WriteByte(line[index])
		}
	}
	withoutStrings := code.String()
	for index := 0; index+3 <= len(withoutStrings); index++ {
		if !strings.EqualFold(withoutStrings[index:index+3], "rem") {
			continue
		}
		if index+3 < len(withoutStrings) && isVBAIdentifierByte(withoutStrings[index+3]) {
			continue
		}
		prefix := strings.TrimSpace(withoutStrings[:index])
		if prefix == "" {
			return ""
		}
		if strings.HasSuffix(prefix, ":") {
			return strings.TrimSpace(strings.TrimSuffix(prefix, ":"))
		}
	}
	return withoutStrings
}

func isVBAIdentifierByte(value byte) bool {
	return value >= 'A' && value <= 'Z' || value >= 'a' && value <= 'z' || value >= '0' && value <= '9' || value == '_'
}

func userFormWithEventsFieldNames(source string) (map[string]struct{}, bool) {
	names := make(map[string]struct{})
	complete := true
	for rawLine := range strings.SplitSeq(strings.ReplaceAll(source, "\r\n", "\n"), "\n") {
		for rawStatement := range strings.SplitSeq(vbaCodeWithoutStringsAndComments(rawLine), ":") {
			line := strings.TrimSpace(rawStatement)
			if line == "" || !userFormWithEventsTokenRE.MatchString(line) {
				continue
			}
			match := userFormWithEventsFieldRE.FindStringSubmatch(line)
			if len(match) != 2 {
				complete = false
				continue
			}
			names[strings.ToLower(strings.TrimSpace(match[1]))] = struct{}{}
		}
	}
	return names, complete
}

func hasUserFormEventCandidate(file parsedFile) bool {
	if !strings.EqualFold(file.ModuleKind, "form") {
		return false
	}
	procedures := file.procedureView()
	for procedure := range procedures.All() {
		name := strings.ToLower(procedure.Name)
		if strings.HasSuffix(name, "_change") || strings.HasSuffix(name, "_click") {
			return true
		}
	}
	return false
}

// VBA220 deliberately reports only the initial event surface. The IR records
// broader event metadata, but treating every document procedure as a handler
// would make this safety rule noisier than its published contract. For forms,
// the control prefix must also exist in the designer artifact when that
// artifact is available; helper names such as Apply_BitmapVector_Change are
// not UserForm events merely because they share an event suffix.
func eventHandlerKind(file parsedFile, proc sourceProcedure) string {
	name := strings.ToLower(proc.Name)
	if proc.ModuleKind == "document" {
		switch name {
		case "worksheet_change", "workbook_sheetchange":
			return "cell"
		case "worksheet_calculate":
			return "calculation"
		case "worksheet_selectionchange":
			return "selection"
		case "workbook_open":
			return "open"
		case "workbook_beforeclose":
			return "close"
		}
	}
	if proc.ModuleKind == "form" {
		suffix, kind := "", ""
		switch {
		case strings.HasSuffix(name, "_change"):
			suffix, kind = "_change", "control-change"
		case strings.HasSuffix(name, "_click"):
			suffix, kind = "_click", "control-click"
		default:
			return ""
		}
		if strings.HasPrefix(name, "test") && !file.UserFormControlsKnown {
			return ""
		}
		if file.UserFormControlsKnown {
			control := strings.TrimSuffix(name, suffix)
			// UserForm_Click is the intrinsic form event procedure. It is
			// not represented as a control in the designer artifact.
			if control != "userform" {
				if _, ok := file.UserFormControlNames[control]; !ok {
					return ""
				}
			} else if suffix != "_click" {
				return ""
			}
		}
		return kind
	}
	return ""
}

func (a Analyzer) eventHandlerReentryFindings(file parsedFile, proc sourceProcedure, project effects.ProjectSummary) []Finding {
	handler := eventHandlerKind(file, proc)
	if handler == "" || proc.Effects == nil {
		return nil
	}
	candidates := map[string]eventFindingCandidate{}
	record := func(line int, statementID int, boundary string, effect effects.Evidence, uncertainty string) {
		if uncertainty == "" && a.eventSafeProcedures[effect.Origin.Key()] {
			return
		}
		category, same, ok := eventEffectRisk(proc, handler, effect)
		if uncertainty != "" {
			category, same, ok = "unknown", false, true
		}
		if !ok || (category != "control" && eventGuardedAt(proc, statementID)) {
			return
		}
		classification := "broader event-chain risk"
		if same {
			classification = "same-event recursion hazard"
		}
		candidate := eventFindingCandidate{
			line: line, statementID: statementID, boundary: boundary, category: category,
			classification: classification, uncertainty: uncertainty,
			effect: effect, same: same,
		}
		key := strings.Join([]string{strconvItoa(line), boundary}, ":")
		if previous, exists := candidates[key]; exists && !eventFindingCandidatePreferred(candidate, previous) {
			return
		}
		candidates[key] = candidate
	}

	for _, evidence := range proc.Effects.Direct {
		record(evidence.Range.StartLine, evidence.StatementID, eventStatementBoundary(evidence.StatementID), evidence, "")
	}
	for _, uncertainty := range proc.Effects.DirectUncertainty {
		record(uncertainty.Range.StartLine, uncertainty.StatementID, eventStatementBoundary(uncertainty.StatementID), effects.Evidence{}, string(uncertainty.Kind))
	}
	for call := range proc.Calls.All() {
		if call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
			continue
		}
		summary, ok := summaryForCandidate(project, call.Resolution.Candidates[0])
		if !ok {
			continue
		}
		safeCallee := a.eventSafeProcedures[summary.Identity.Key()]
		boundary := eventCallBoundary(call)
		for _, evidence := range append(append([]effects.Evidence{}, summary.Direct...), summary.Propagated...) {
			if safeCallee {
				continue
			}
			record(call.Range.StartLine, call.StatementID, boundary, evidence, "")
		}
		for _, uncertainty := range append(append([]effects.CallUncertainty{}, summary.DirectUncertainty...), summary.PropagatedUncertainty...) {
			record(call.Range.StartLine, call.StatementID, boundary, effects.Evidence{}, string(uncertainty.Kind))
		}
	}

	ordered := make([]eventFindingCandidate, 0, len(candidates))
	for _, candidate := range candidates {
		ordered = append(ordered, candidate)
	}
	sort.SliceStable(ordered, func(i, j int) bool {
		if ordered[i].line != ordered[j].line {
			return ordered[i].line < ordered[j].line
		}
		if ordered[i].statementID != ordered[j].statementID {
			return ordered[i].statementID < ordered[j].statementID
		}
		if ordered[i].boundary != ordered[j].boundary {
			return ordered[i].boundary < ordered[j].boundary
		}
		return eventFindingCandidateKey(ordered[i]) < eventFindingCandidateKey(ordered[j])
	})

	out := make([]Finding, 0, len(ordered))
	for _, candidate := range ordered {
		message := "Event handler " + proc.Name + " has a " + candidate.classification + "."
		reason := "This handler can trigger " + candidate.category + " event processing"
		if candidate.uncertainty != "" {
			message = "Event handler " + proc.Name + " reaches an " + candidate.uncertainty + " call that may trigger an event."
			reason = "The call cannot be resolved safely, so its event-triggering effects are uncertain."
		} else if candidate.effect.Origin.QualifiedName != "" && !strings.EqualFold(candidate.effect.Origin.QualifiedName, proc.Effects.Identity.QualifiedName) {
			reason += " through " + candidate.effect.Origin.QualifiedName + "."
		} else {
			reason += "."
		}
		out = append(out, a.simpleFinding(file, proc, candidate.line, "VBA220", "warning", message, reason, "Disable Application.EnableEvents around Excel event-triggering work and restore it on every exit; use a re-entry guard for UserForm controls."))
	}
	return out
}

func (a Analyzer) userFormControlNames(file parsedFile) (map[string]struct{}, bool) {
	if !strings.EqualFold(file.ModuleKind, "form") {
		return nil, false
	}

	var source string
	var designerPath string
	designerFound := false
	if file.sourceProject {
		logicalPath := strings.ReplaceAll(file.Path, "\\", "/")
		source = file.UserFormDesignerSource
		designerPath = file.UserFormDesignerSourcePath
		designerFound = file.UserFormDesignerSourceFound
		if strings.EqualFold(pathpkg.Ext(logicalPath), ".frm") {
			source = string(file.Source)
			designerPath = file.Path
			designerFound = true
		}
		if !designerFound {
			return nil, false
		}
	} else if strings.EqualFold(filepath.Ext(file.Path), ".frm") {
		source = string(file.Source)
		designerPath = file.Path
		designerFound = true
	} else {
		name := strings.TrimSuffix(filepath.Base(file.Path), filepath.Ext(file.Path))
		candidates := make([]string, 0, 2)
		formsRoot := ""
		if a.Config.Src.Forms != "" {
			formsRoot = filepath.Join(a.RootDir, filepath.FromSlash(a.Config.Src.Forms))
			candidates = append(candidates, filepath.Join(formsRoot, name+".frm"))
		}
		source, designerPath, designerFound = readUserFormDesigner(candidates)
		if !designerFound {
			return nil, false
		}
	}

	form := userforms.Parse(source)
	expectedName := strings.TrimSuffix(filepath.Base(file.Path), filepath.Ext(file.Path))
	if file.sourceProject {
		logicalPath := strings.ReplaceAll(file.Path, "\\", "/")
		expectedName = strings.TrimSuffix(pathpkg.Base(logicalPath), pathpkg.Ext(logicalPath))
	}
	if !form.Complete || strings.TrimSpace(form.Name) == "" || !strings.EqualFold(strings.TrimSpace(form.Name), expectedName) {
		return nil, false
	}
	controls := make(map[string]struct{}, len(form.Controls))
	for _, control := range form.Controls {
		name := strings.ToLower(strings.TrimSpace(control.Name))
		if name != "" {
			controls[name] = struct{}{}
		}
	}
	for _, withEventsSource := range []string{string(file.Source), source} {
		withEvents, complete := userFormWithEventsFieldNames(withEventsSource)
		if !complete {
			return controls, false
		}
		for name := range withEvents {
			controls[name] = struct{}{}
		}
	}
	if !designerFound {
		return controls, false
	}
	if file.sourceProject {
		// The .frm text is caller-supplied, but a referenced .frx is a separate
		// external artifact. Keep the event classification conservative without
		// resolving that reference against the host filesystem.
		if _, referenced := userFormDesignerFRXReference(source); referenced {
			return controls, false
		}
		return controls, true
	}
	frxPath, frxReferenced, err := userFormFRXPath(designerPath, source)
	if err != nil {
		return controls, false
	}
	if !frxReferenced {
		return controls, true
	}
	frxControls, parsed, complete, err := cachedUserFormFRXControlNames(frxPath)
	if err != nil || !parsed || !complete || len(frxControls) == 0 {
		return controls, false
	}
	for name := range frxControls {
		controls[name] = struct{}{}
	}
	return controls, true
}

func userFormFRXPath(designerPath, designerSource string) (string, bool, error) {
	reference, referenced := userFormDesignerFRXReference(designerSource)
	if !referenced {
		return "", false, nil
	}
	if reference == "" {
		return "", false, errors.New("invalid UserForm FRX reference")
	}
	entries, err := os.ReadDir(filepath.Dir(designerPath))
	if err != nil {
		return "", false, err
	}
	matches := make([]string, 0, 1)
	for _, entry := range entries {
		if !strings.EqualFold(entry.Name(), reference) {
			continue
		}
		if entry.Type()&os.ModeSymlink != 0 {
			return "", false, errors.New("UserForm FRX reference resolves through a symlink")
		}
		if entry.IsDir() {
			continue
		}
		matches = append(matches, entry.Name())
	}
	if len(matches) == 0 {
		return filepath.Join(filepath.Dir(designerPath), reference), true, nil
	}
	if len(matches) != 1 {
		return "", false, errors.New("ambiguous UserForm FRX artifacts")
	}
	return filepath.Join(filepath.Dir(designerPath), matches[0]), true, nil
}

func userFormDesignerFRXReference(source string) (string, bool) {
	sawReference := false
	source = strings.ReplaceAll(source, "\r\n", "\n")
	source = strings.ReplaceAll(source, "\r", "\n")
	for _, line := range strings.Split(source, "\n") {
		trimmed := strings.TrimSpace(line)
		if len(trimmed) < len("OleObjectBlob") || !strings.EqualFold(trimmed[:len("OleObjectBlob")], "OleObjectBlob") {
			continue
		}
		if !strings.Contains(strings.ToLower(trimmed), ".frx") {
			continue
		}
		sawReference = true
		start := strings.IndexByte(trimmed, '"')
		if start < 0 {
			continue
		}
		end := strings.IndexByte(trimmed[start+1:], '"')
		if end < 0 {
			continue
		}
		reference := strings.TrimSpace(trimmed[start+1 : start+1+end])
		reference = filepath.Base(filepath.FromSlash(strings.ReplaceAll(reference, "\\", "/")))
		if strings.EqualFold(filepath.Ext(reference), ".frx") && reference != "" {
			return reference, true
		}
	}
	return "", sawReference
}

func readUserFormDesigner(paths []string) (string, string, bool) {
	for _, path := range paths {
		if body, ok := readUserFormDesignerPath(path); ok {
			return body, path, true
		}
		entries, err := os.ReadDir(filepath.Dir(path))
		if err != nil {
			continue
		}
		matches := make([]string, 0, 1)
		for _, entry := range entries {
			if entry.IsDir() || entry.Type()&os.ModeSymlink != 0 || !strings.EqualFold(entry.Name(), filepath.Base(path)) {
				continue
			}
			matches = append(matches, filepath.Join(filepath.Dir(path), entry.Name()))
		}
		if len(matches) != 1 {
			continue
		}
		if body, ok := readUserFormDesignerPath(matches[0]); ok {
			return body, matches[0], true
		}
	}
	return "", "", false
}

func readUserFormDesignerPath(path string) (string, bool) {
	info, err := os.Lstat(path)
	if err != nil || info.Mode()&os.ModeSymlink != 0 || info.IsDir() {
		return "", false
	}
	body, err := os.ReadFile(path)
	if err != nil {
		return "", false
	}
	return string(body), true
}

type userFormFRXCacheEntry struct {
	size     int64
	modTime  int64
	names    map[string]struct{}
	parsed   bool
	complete bool
}

var userFormFRXCache = struct {
	sync.RWMutex
	entries map[string]userFormFRXCacheEntry
}{entries: make(map[string]userFormFRXCacheEntry)}

func cachedUserFormFRXControlNames(path string) (map[string]struct{}, bool, bool, error) {
	cachePath, err := filepath.Abs(path)
	if err != nil {
		cachePath = filepath.Clean(path)
	}
	info, err := os.Stat(cachePath)
	if err != nil {
		return nil, false, false, err
	}
	modTime := info.ModTime().UnixNano()
	userFormFRXCache.RLock()
	entry, ok := userFormFRXCache.entries[cachePath]
	userFormFRXCache.RUnlock()
	if ok && entry.size == info.Size() && entry.modTime == modTime {
		return cloneUserFormFRXNames(entry.names), entry.parsed, entry.complete, nil
	}

	data, err := os.ReadFile(cachePath)
	if err != nil {
		return nil, false, false, err
	}
	names, parsed, complete := userforms.ExtractFRXControlNames(data)
	stored := userFormFRXCacheEntry{
		size:     info.Size(),
		modTime:  modTime,
		names:    cloneUserFormFRXNames(names),
		parsed:   parsed,
		complete: complete,
	}
	userFormFRXCache.Lock()
	userFormFRXCache.entries[cachePath] = stored
	userFormFRXCache.Unlock()
	return cloneUserFormFRXNames(stored.names), stored.parsed, stored.complete, nil
}

func cloneUserFormFRXNames(names map[string]struct{}) map[string]struct{} {
	if len(names) == 0 {
		return nil
	}
	clone := make(map[string]struct{}, len(names))
	for name := range names {
		clone[name] = struct{}{}
	}
	return clone
}

func eventFindingCandidatePreferred(candidate, previous eventFindingCandidate) bool {
	candidateRank, previousRank := eventFindingCandidateRank(candidate), eventFindingCandidateRank(previous)
	if candidateRank != previousRank {
		return candidateRank > previousRank
	}
	return eventFindingCandidateKey(candidate) < eventFindingCandidateKey(previous)
}

func eventFindingCandidateRank(candidate eventFindingCandidate) int {
	if candidate.uncertainty != "" {
		return 1
	}
	if candidate.same {
		return 3
	}
	return 2
}

func eventFindingCandidateKey(candidate eventFindingCandidate) string {
	return strings.Join([]string{candidate.boundary, candidate.classification, candidate.category, candidate.uncertainty, candidate.effect.Origin.QualifiedName}, ":")
}

func eventStatementBoundary(statementID int) string {
	return "statement:" + strconvItoa(statementID)
}

func eventCallBoundary(call procedureir.CallSite) string {
	if call.ID != 0 {
		return "call:" + strconvItoa(call.ID)
	}
	return strings.Join([]string{
		"call-range",
		strconvItoa(call.Range.StartByte),
		strconvItoa(call.Range.EndByte),
	}, ":")
}

func eventEffectRisk(proc sourceProcedure, handler string, evidence effects.Evidence) (string, bool, bool) {
	category := ""
	switch evidence.Effect {
	case effects.WritesCells:
		category = "cell"
	case effects.Recalculates:
		category = "calculation"
	case effects.ChangesSelection:
		category = "selection"
	case effects.OpensWorkbook:
		category = "open"
	case effects.ClosesWorkbook:
		category = "close"
	case effects.ChangesControls:
		category = "control"
	case effects.ChangesWorkbook:
		lower := strings.ToLower(evidence.Target)
		if strings.Contains(lower, "range(") || strings.Contains(lower, "cells(") || strings.Contains(lower, "rows(") || strings.Contains(lower, "columns(") || strings.HasSuffix(lower, ".value") || strings.HasSuffix(lower, ".value2") || strings.HasSuffix(lower, ".formula") || strings.HasSuffix(lower, ".formular1c1") {
			return "", false, false
		}
		category = "workbook structure"
	default:
		return "", false, false
	}
	if handler == "control-change" && category == "control" {
		control := strings.TrimSuffix(strings.ToLower(proc.Name), "_change")
		return category, strings.EqualFold(eventControlTargetName(evidence.Target), control), true
	}
	return category, handler == category, true
}

func eventControlTargetName(target string) string {
	lower := strings.ToLower(strings.TrimSpace(target))
	lower = strings.TrimPrefix(lower, "me.")
	if strings.HasPrefix(lower, "controls(\"") {
		if end := strings.Index(lower[len("controls(\""):], "\""); end >= 0 {
			return lower[len("controls(\"") : len("controls(\"")+end]
		}
	}
	if index := strings.IndexByte(lower, '.'); index >= 0 {
		return lower[:index]
	}
	return ""
}

func eventSafeProcedures(files []parsedFile, project effects.ProjectSummary) map[string]bool {
	safe := map[string]bool{}
	for _, file := range files {
		procedures := file.procedureView()
		for index := 0; index < procedures.Len(); index++ {
			proc := procedures.valueAt(index)
			if index >= len(file.IR.Procedures) {
				continue
			}
			id := procedureEffectIdentity(file.IR, file.IR.Procedures[index].Symbol)
			summary, ok := project.LookupDirect(id)
			if !ok || !summary.Has(effects.DisablesEvents) || !summary.Has(effects.RestoresEvents) {
				continue
			}
			hasTrigger, allGuarded := false, true
			for _, evidence := range summary.Direct {
				if _, _, relevant := eventEffectRisk(proc, "", evidence); relevant {
					hasTrigger = true
					if !eventGuardedAt(proc, evidence.StatementID) {
						allGuarded = false
					}
				}
			}
			for call := range proc.Calls.All() {
				if call.Resolution.Status != procedureir.ResolutionMatched || len(call.Resolution.Candidates) != 1 {
					continue
				}
				callee, ok := summaryForCandidate(project, call.Resolution.Candidates[0])
				if !ok {
					continue
				}
				for _, evidence := range append(append([]effects.Evidence{}, callee.Direct...), callee.Propagated...) {
					if _, _, relevant := eventEffectRisk(proc, "", evidence); relevant {
						hasTrigger = true
						if !eventGuardedAt(proc, call.StatementID) {
							allGuarded = false
						}
					}
				}
			}
			if hasTrigger && allGuarded {
				safe[id.Key()] = true
			}
		}
	}
	return safe
}

func summaryForCandidate(project effects.ProjectSummary, candidate procedureir.Candidate) (effects.ProcedureSummary, bool) {
	return project.LookupCandidate(candidate)
}

// eventGuardedAt accepts a guard only when its False assignment dominates the
// effect on normal paths and VBA203's existing all-exit analysis proves that
// assignment is restored. UserForm controls are intentionally excluded.
func eventGuardedAt(proc sourceProcedure, statementID int) bool {
	if proc.ModuleKind == "form" || proc.Graph == nil || statementID == 0 {
		return false
	}
	facts := proc.analysisFacts()
	unsafe := applicationStateExitWitnesses(proc, "enableevents", facts)
	target, ok := proc.Graph.BlockForStatement(statementID)
	if !ok {
		return false
	}
	dominators := proc.Graph.Dominators(vbacfg.EdgeFilter{NormalOnly: true})[target.ID]
	for statement := range proc.Statements.All() {
		property, value, ok := applicationPropertyAssignment(statement, facts)
		if !ok || property != "enableevents" || unsafe[statement.ID].Kind != "" || !eventDisableValue(value) {
			continue
		}
		block, ok := proc.Graph.BlockForStatement(statement.ID)
		if ok && eventContainsBlock(dominators, block.ID) {
			return true
		}
	}
	return false
}

func eventContainsBlock(blocks []vbacfg.BlockID, target vbacfg.BlockID) bool {
	for _, block := range blocks {
		if block == target {
			return true
		}
	}
	return false
}

func eventDisableValue(value string) bool {
	value = strings.ToLower(strings.TrimSpace(value))
	return value == "false" || value == "0"
}
