package gui

import (
	"bufio"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/callgraph"
	"github.com/harumiWeb/xlflow/internal/vba/calls"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

type Boundary struct {
	File       string `json:"file"`
	Line       int    `json:"line"`
	Kind       string `json:"kind"`
	Symbol     string `json:"symbol"`
	Severity   string `json:"severity"`
	Message    string `json:"message"`
	Suggestion string `json:"suggestion"`
}

type Analyzer struct {
	RootDir string
	Config  config.Config
}

type detector struct {
	re         *regexp.Regexp
	kind       string
	symbol     string
	message    string
	suggestion string
	keepString bool
}

var (
	msgBoxFunctionRe   = regexp.MustCompile(`(?i)\b(?:(?:public|private|friend)\s+)?function\s+msgbox\b`)
	inputBoxFunctionRe = regexp.MustCompile(`(?i)\b(?:(?:public|private|friend)\s+)?function\s+inputbox\b`)
)

var detectors = []detector{
	detect(`(?i)\bapplication\s*\.\s*getopenfilename\b`, "file_picker", "Application.GetOpenFilename", "File picker requires human interaction and bypasses XlflowUI.", "Replace it with XlflowUI.GetOpenFilename(\"<dialog-id>\", ...) or XlflowUI.FileDialogOpen(\"<dialog-id>\", ...) so headless runs can pass --filedialog responses."),
	detect(`(?i)\bapplication\s*\.\s*getsaveasfilename\b`, "file_picker", "Application.GetSaveAsFilename", "File picker requires human interaction and bypasses XlflowUI.", "Replace it with XlflowUI.GetSaveAsFilename(\"<dialog-id>\", ...) so headless runs can pass --filedialog responses."),
	detect(`(?i)\bapplication\s*\.\s*filedialog\b`, "file_picker", "Application.FileDialog", "File dialog requires human interaction and bypasses XlflowUI.", "Replace file-open and folder-picker flows with XlflowUI.FileDialogOpen(\"<dialog-id>\", ...) or XlflowUI.FolderPicker(\"<dialog-id>\", ...) so headless runs can pass --filedialog responses."),
	detect(`(?i)\binputbox\s*(?:\(|")?`, "modal_dialog", "InputBox", "Raw InputBox requires human input and bypasses XlflowUI.", "Replace it with XlflowUI.InputBox(\"<dialog-id>\", ...) so headless, test, and agent runs can pass --inputbox responses."),
	detect(`(?i)\bmsgbox\s*(?:\(|")?`, "modal_dialog", "MsgBox", "Raw MsgBox blocks unattended execution and bypasses XlflowUI.", "Replace it with XlflowUI.MsgBox(\"<dialog-id>\", ...) so headless, test, and agent runs can pass --msgbox responses."),
	detect(`(?i)\b[A-Za-z_][A-Za-z0-9_]*\s*\.\s*show\b`, "user_form", "UserForm.Show", "UserForm display requires human interaction.", "Keep UserForm entrypoints interactive-only and extract core logic into parameterized procedures."),
	detect(`(?i)\.\s*show\s+vbmodal\b`, "user_form", ".Show vbModal", "Modal form display requires human interaction.", "Keep modal UI entrypoints interactive-only and extract core logic into parameterized procedures."),
	detect(`(?i)\bdoevents\b`, "message_pump", "DoEvents", "DoEvents can hide GUI waits or message-pump dependent behavior.", "Avoid message-pump dependent control flow in headless macros."),
	detect(`(?i)^\s*shell\s*(?:\(|")?`, "external_process", "Shell", "Shell starts an external process from VBA.", "Prefer explicit CLI orchestration or document this macro as interactive/external-process dependent."),
	detectWithStrings(`(?i)\bcreateobject\s*\(\s*"wscript\.shell"\s*\)\s*\.\s*popup\b`, "modal_dialog", `CreateObject("WScript.Shell").Popup`, "WScript popup blocks unattended execution.", "If this is just a confirmation dialog, prefer XlflowUI.MsgBox with a stable dialog id; otherwise keep it behind an interactive-only adapter."),
}

func detect(pattern, kind, symbol, message, suggestion string) detector {
	return detector{
		re:         regexp.MustCompile(pattern),
		kind:       kind,
		symbol:     symbol,
		message:    message,
		suggestion: suggestion,
	}
}

func detectWithStrings(pattern, kind, symbol, message, suggestion string) detector {
	d := detect(pattern, kind, symbol, message, suggestion)
	d.keepString = true
	return d
}

func shouldIgnoreDetectorLine(detector detector, code string) bool {
	lower := strings.ToLower(code)
	switch detector.symbol {
	case "MsgBox":
		if strings.Contains(lower, "xlflowui.msgbox") {
			return true
		}
		return msgBoxFunctionRe.MatchString(code)
	case "InputBox":
		if strings.Contains(lower, "xlflowui.inputbox") {
			return true
		}
		return inputBoxFunctionRe.MatchString(code)
	default:
		return false
	}
}

func (a Analyzer) Run() ([]Boundary, error) {
	files, err := a.files()
	if err != nil {
		return nil, err
	}
	boundaries := make([]Boundary, 0)
	for _, file := range files {
		fileBoundaries, err := a.AnalyzeFile(file)
		if err != nil {
			return nil, err
		}
		boundaries = append(boundaries, fileBoundaries...)
	}
	return boundaries, nil
}

// RunHeadless analyzes GUI boundaries in the context of one requested macro.
// The ordinary Run method intentionally remains a project-wide report for
// inspect-gui, lint, and doctor. Unknown project dispatch is treated as a
// reason to retain every boundary, since it is not sound to use an incomplete
// call graph as proof that a boundary is unreachable.
func (a Analyzer) RunHeadless(macro string) ([]Boundary, error) {
	files, err := a.headlessFiles()
	if err != nil {
		return nil, err
	}
	all, err := a.runFiles(files)
	if err != nil {
		return nil, err
	}
	if strings.TrimSpace(macro) == "" {
		return all, nil
	}

	callResult, err := calls.Inspect(calls.Options{RootDir: a.RootDir, Config: a.Config})
	if err != nil {
		return nil, err
	}
	if !validateHeadlessTarget(callResult.Symbols, macro) {
		// Keep the existing bridge-side macro_not_found behavior when no GUI
		// boundary is present, but never interpret an unknown or ambiguous root
		// as proof that boundaries are unreachable.
		return all, nil
	}
	roots := headlessRoots(callResult, macro)
	reachability := callgraph.AnalyzeReachability(callgraph.SnapshotFromResult(callResult), callgraph.ReachabilityRequest{
		Roots: roots,
	})
	reachable := make(map[string]bool, len(reachability.Confirmed)+len(reachability.Possible))
	for _, node := range append(reachability.Confirmed, reachability.Possible...) {
		reachable[headlessProcedureKey(node.ID.QualifiedName, node.ID.Kind, node.ID.File, node.ID.Line)] = true
	}

	// A parse recovery node or an uncertain project dispatch means the graph is
	// not a proof of non-reachability. Keep the project-wide behavior in that
	// case instead of silently allowing a GUI boundary.
	if headlessNeedsGlobalFallback(callResult, reachable) {
		return all, nil
	}

	filtered := make([]Boundary, 0, len(all))
	for _, boundary := range all {
		owner, ok := boundaryOwner(callResult.Symbols, boundary)
		if !ok || reachable[headlessProcedureKey(owner.Module+"."+owner.Name, owner.Kind, owner.File, owner.StartLine)] {
			filtered = append(filtered, boundary)
		}
	}
	return filtered, nil
}

func headlessRoots(result *calls.Result, macro string) []callgraph.Root {
	roots := []callgraph.Root{{Target: macro, Confidence: callgraph.RootConfirmed, Reason: "run target"}}
	for _, item := range result.Symbols {
		if !isProcedureSymbol(item.Kind) {
			continue
		}
		moduleKind := result.ModuleKinds[item.File]
		if !strings.EqualFold(moduleKind, "standard") {
			// Document modules, classes, and forms have host lifecycle and event
			// entrypoints that are not necessarily called by the requested macro.
			// Treating them as possible roots prevents pruning a boundary that
			// Excel can enter outside the ordinary project call graph.
			roots = append(roots, callgraph.Root{Target: item.Module + "." + item.Name, Confidence: callgraph.RootPossible, Reason: "host or object lifecycle"})
			continue
		}
		if event, _ := procedureir.ClassifyEvent(moduleKind, item.Name); event {
			roots = append(roots, callgraph.Root{Target: item.Module + "." + item.Name, Confidence: callgraph.RootPossible, Reason: "automatic macro"})
			continue
		}
		if strings.HasPrefix(item.Kind, "property") {
			// Property Get/Let/Set can be invoked implicitly by an assignment or
			// expression, so the syntax-local call extractor cannot prove that a
			// private accessor is unused. Keep every standard-module accessor as a
			// possible root for headless safety.
			roots = append(roots, callgraph.Root{Target: item.Module + "." + item.Name, Confidence: callgraph.RootPossible, Reason: "implicit property access"})
			continue
		}
		if !strings.EqualFold(item.Visibility, "Private") && item.Kind == "function" {
			// Public standard-module functions can be called by worksheet formulas
			// or external VBA without a project-local edge.
			roots = append(roots, callgraph.Root{Target: item.Module + "." + item.Name, Confidence: callgraph.RootPossible, Reason: "public standard-module API"})
		}
	}
	return roots
}

func (a Analyzer) headlessFiles() ([]string, error) {
	files, err := a.files()
	if err != nil {
		return nil, err
	}
	// Analyzer.files deliberately preserves the historical inspect-gui input
	// set. Headless reachability also includes document modules, because they
	// can be called by a target macro and must not become an analysis blind spot.
	root := filepath.Join(a.RootDir, a.Config.Src.Workbook)
	if strings.TrimSpace(a.Config.Src.Workbook) != "" {
		extra, walkErr := sourceFilesUnder(root)
		if walkErr != nil {
			return nil, walkErr
		}
		files = appendUniquePaths(files, extra...)
	}
	return files, nil
}

func (a Analyzer) runFiles(files []string) ([]Boundary, error) {
	boundaries := make([]Boundary, 0)
	for _, file := range files {
		fileBoundaries, err := a.AnalyzeFile(file)
		if err != nil {
			return nil, err
		}
		boundaries = append(boundaries, fileBoundaries...)
	}
	return boundaries, nil
}

func sourceFilesUnder(root string) ([]string, error) {
	var files []string
	err := filepath.WalkDir(root, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if d.IsDir() {
			return nil
		}
		switch strings.ToLower(filepath.Ext(path)) {
		case ".bas", ".cls", ".frm":
			files = append(files, path)
		}
		return nil
	})
	if err != nil {
		if os.IsNotExist(err) {
			return nil, nil
		}
		return nil, err
	}
	return files, nil
}

func appendUniquePaths(paths []string, extras ...string) []string {
	seen := make(map[string]bool, len(paths)+len(extras))
	for _, path := range paths {
		seen[strings.ToLower(filepath.Clean(path))] = true
	}
	for _, path := range extras {
		key := strings.ToLower(filepath.Clean(path))
		if !seen[key] {
			paths = append(paths, path)
			seen[key] = true
		}
	}
	return paths
}

func validateHeadlessTarget(items []symbols.Symbol, macro string) bool {
	target := normalizeHeadlessTarget(macro)
	if target == "" {
		return false
	}
	matches := 0
	for _, item := range items {
		if !isProcedureSymbol(item.Kind) {
			continue
		}
		qualified := strings.ToLower(item.Module + "." + item.Name)
		name := strings.ToLower(item.Name)
		if strings.Contains(target, ".") {
			if qualified == target {
				matches++
			}
		} else if name == target {
			matches++
		}
	}
	switch matches {
	case 0:
		return false
	case 1:
		return true
	default:
		return false
	}
}

func normalizeHeadlessTarget(target string) string {
	target = strings.TrimSpace(target)
	if i := strings.LastIndex(target, "!"); i >= 0 {
		target = target[i+1:]
	}
	target = strings.Trim(strings.TrimSpace(target), "'")
	return strings.ToLower(strings.TrimSpace(target))
}

func isProcedureSymbol(kind string) bool {
	switch kind {
	case "sub", "function", "property", "property_get", "property_let", "property_set":
		return true
	default:
		return false
	}
}

func headlessProcedureKey(qualified, kind, file string, line int) string {
	return strings.ToLower(strings.Join([]string{qualified, kind, filepath.ToSlash(filepath.Clean(file)), strconv.Itoa(line)}, "\x00"))
}

func boundaryOwner(items []symbols.Symbol, boundary Boundary) (symbols.Symbol, bool) {
	line := boundary.Line
	best := symbols.Symbol{}
	found := false
	for _, item := range items {
		if !isProcedureSymbol(item.Kind) || !strings.EqualFold(filepath.ToSlash(item.File), boundary.File) {
			continue
		}
		if line < item.StartLine || (item.EndLine > 0 && line > item.EndLine) {
			continue
		}
		if !found || item.StartLine > best.StartLine {
			best = item
			found = true
		}
	}
	return best, found
}

func headlessNeedsGlobalFallback(result *calls.Result, reachable map[string]bool) bool {
	if result.Summary.ParseErrors > 0 || result.Summary.MissingNodes > 0 {
		// A malformed source cannot safely support pruning. The call extractor
		// reports this state in the aggregate summary even when no call sites were
		// recovered from the affected file.
		return true
	}
	for _, call := range result.Calls {
		if call.Caller == nil || !headlessCallerReachable(result.Symbols, call, reachable) {
			continue
		}
		switch call.Resolution.Status {
		case "matched", "external", "builtin_like", "non_callable":
			continue
		case "unresolved":
			// Known dynamic VBA APIs are represented by DynamicReferences. Their
			// static targets are propagated by the call graph; an unknown target
			// is handled below, scoped to reachable callers. Do not let the
			// legacy unresolved projection turn every static Application.Run
			// into a false global fallback.
			if headlessKnownDynamicCall(call) {
				continue
			}
			return true
		default:
			return true
		}
	}
	for _, ref := range result.DynamicReferences {
		if ref.Caller != nil && !headlessCallerReachable(result.Symbols, calls.Call{CallSite: calls.CallSite{File: ref.File, Caller: ref.Caller}}, reachable) {
			continue
		}
		if ref.Target == "" {
			return true
		}
	}
	return false
}

func headlessKnownDynamicCall(call calls.Call) bool {
	text := strings.ToLower(strings.ReplaceAll(strings.TrimSpace(call.Callee.Text), " ", ""))
	text = strings.ReplaceAll(text, "_", "")
	switch text {
	case "application.run", "application.ontime", "application.onkey", "callbyname":
		return true
	default:
		return false
	}
}

func headlessCallerReachable(items []symbols.Symbol, call calls.Call, reachable map[string]bool) bool {
	if call.Caller == nil {
		return false
	}
	for _, item := range items {
		if isProcedureSymbol(item.Kind) && strings.EqualFold(item.File, call.File) && strings.EqualFold(item.Module+"."+item.Name, call.Caller.QualifiedName) {
			return reachable[headlessProcedureKey(item.Module+"."+item.Name, item.Kind, item.File, item.StartLine)]
		}
	}
	return false
}

func (a Analyzer) files() ([]string, error) {
	dirs := []string{
		a.Config.Src.Modules,
		a.Config.Src.Classes,
		a.Config.Src.Forms,
		"tests",
	}
	var files []string
	for _, dir := range dirs {
		root := filepath.Join(a.RootDir, dir)
		if _, err := os.Stat(root); err != nil {
			if os.IsNotExist(err) {
				continue
			}
			return nil, err
		}
		err := filepath.WalkDir(root, func(path string, d os.DirEntry, err error) error {
			if err != nil {
				return err
			}
			if d.IsDir() {
				return nil
			}
			switch strings.ToLower(filepath.Ext(path)) {
			case ".bas", ".cls", ".frm":
				files = append(files, path)
			}
			return nil
		})
		if err != nil {
			return nil, err
		}
	}
	return files, nil
}

func (a Analyzer) AnalyzeFile(path string) (boundaries []Boundary, err error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer func() {
		if closeErr := f.Close(); err == nil && closeErr != nil {
			err = closeErr
		}
	}()

	scanner := bufio.NewScanner(f)
	lineNo := 0
	for scanner.Scan() {
		lineNo++
		code := StripComment(scanner.Text())
		codeWithoutStrings := detectionText(code)
		for _, detector := range detectors {
			if strings.EqualFold(filepath.Base(path), "XlflowUI.bas") && (detector.symbol == "MsgBox" || detector.symbol == "InputBox" || detector.symbol == "UserForm.Show" || detector.kind == "file_picker") {
				continue
			}
			if shouldIgnoreDetectorLine(detector, code) {
				continue
			}
			input := codeWithoutStrings
			if detector.keepString {
				input = code
			}
			if detector.re.MatchString(input) {
				boundaries = append(boundaries, a.boundary(path, lineNo, detector))
			}
		}
	}
	if err := scanner.Err(); err != nil {
		return nil, err
	}
	return boundaries, nil
}

func detectionText(line string) string {
	var b strings.Builder
	inString := false
	for i := 0; i < len(line); i++ {
		if line[i] != '"' {
			if inString {
				b.WriteByte(' ')
			} else {
				b.WriteByte(line[i])
			}
			continue
		}
		b.WriteByte('"')
		if inString && i+1 < len(line) && line[i+1] == '"' {
			b.WriteByte('"')
			i++
			continue
		}
		inString = !inString
	}
	return b.String()
}

func (a Analyzer) boundary(path string, line int, d detector) Boundary {
	file, err := filepath.Rel(a.RootDir, path)
	if err != nil {
		file = path
	}
	return Boundary{
		File:       filepath.ToSlash(file),
		Line:       line,
		Kind:       d.kind,
		Symbol:     d.symbol,
		Severity:   "interactive-only",
		Message:    d.message,
		Suggestion: d.suggestion,
	}
}

func StripComment(line string) string {
	inString := false
	for i := 0; i < len(line); i++ {
		switch line[i] {
		case '"':
			if inString && i+1 < len(line) && line[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '\'':
			if !inString {
				return line[:i]
			}
		}
	}
	return line
}
