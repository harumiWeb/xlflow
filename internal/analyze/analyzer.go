package analyze

import (
	"bufio"
	"os"
	"path/filepath"
	"regexp"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/gui"
)

type Finding struct {
	Code       string   `json:"code"`
	Severity   string   `json:"severity"`
	File       string   `json:"file"`
	Module     string   `json:"module,omitempty"`
	Procedure  string   `json:"procedure,omitempty"`
	Line       int      `json:"line"`
	Message    string   `json:"message"`
	Reason     string   `json:"reason"`
	Suggestion string   `json:"suggestion"`
	NearbyCode []string `json:"nearby_code,omitempty"`
}

type Analyzer struct {
	RootDir string
	Config  config.Config
}

var (
	declRe            = regexp.MustCompile(`(?i)^\s*(?:dim|private|public|static)\s+([A-Za-z_][A-Za-z0-9_]*)\s+as\s+(?:new\s+)?([A-Za-z_][A-Za-z0-9_.]*)\b`)
	procRe            = regexp.MustCompile(`(?i)^\s*(?:public\s+|private\s+|friend\s+)?(sub|function|property\s+get)\s+([A-Za-z_][A-Za-z0-9_]*)\b(?:[^']*?\bas\s+([A-Za-z_][A-Za-z0-9_.]*))?`)
	publicProcDeclRe  = regexp.MustCompile(`(?i)^\s*public\s+(?:sub|function|property\s+get)\s+([A-Za-z_][A-Za-z0-9_]*)\b`)
	assignRe          = regexp.MustCompile(`(?i)^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=`)
	setAssign         = regexp.MustCompile(`(?i)^\s*set\s+([A-Za-z_][A-Za-z0-9_]*)\s*=`)
	callAssign        = regexp.MustCompile(`(?i)^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([A-Za-z_][A-Za-z0-9_.]*)\s*(?:\(|$)`)
	withRe            = regexp.MustCompile(`(?i)^\s*with\s+(.+)$`)
	endWithRe         = regexp.MustCompile(`(?i)^\s*end\s+with\b`)
	withMember        = regexp.MustCompile(`(?i)^\s*\.([A-Za-z_][A-Za-z0-9_]*)\b`)
	memberRe          = regexp.MustCompile(`(?i)\b([A-Za-z_][A-Za-z0-9_]*)\s*\.\s*([A-Za-z_][A-Za-z0-9_]*)\b`)
	traceHelperCallRe = regexp.MustCompile(`(?i)^\s*(?:call\s+)?(XlflowLog|XlflowSetTraceFile)\b`)
	traceHelperQualRe = regexp.MustCompile(`(?i)\bXlflowTrace\s*\.\s*(XlflowLog|XlflowSetTraceFile)\b`)
)

var objectTypes = map[string]bool{
	"application": true, "workbook": true, "worksheet": true, "range": true,
	"chart": true, "pivot table": true, "pivottable": true, "listobject": true,
	"dictionary": true, "collection": true, "object": true,
}

type invalidMemberRule struct {
	Reason     string
	Suggestion string
}

type helperDependencyRule struct {
	Code       string
	Reason     string
	Suggestion string
}

var invalidObjectMembers = map[string]map[string]invalidMemberRule{
	"worksheet": {
		"displaygridlines": {
			Reason:     "DisplayGridlines is a Window property, not a Worksheet member.",
			Suggestion: "Set DisplayGridlines on ActiveWindow or another Window object instead of the Worksheet.",
		},
	},
}

var traceHelperDependencies = map[string]helperDependencyRule{
	"xlflowlog": {
		Code:       "VBA105",
		Reason:     "XlflowLog is provided by the xlflow trace helper module. Without a Public standard-module definition, Excel compiles the project with 'Sub or Function not defined'.",
		Suggestion: "Run `xlflow trace enable` to persist XlflowTrace.bas in source, or use `xlflow run --trace` for temporary helper injection during execution.",
	},
	"xlflowsettracefile": {
		Code:       "VBA106",
		Reason:     "XlflowSetTraceFile is provided by the xlflow trace helper module. Without a Public standard-module definition, Excel compiles the project with 'Sub or Function not defined'.",
		Suggestion: "Prefer `xlflow run --trace`, or add the bundled XlflowTrace.bas helper with `xlflow trace enable` before relying on XlflowSetTraceFile in source-controlled VBA.",
	},
}

type analysisContext struct {
	standardModulePublicProcs map[string]bool
}

func (a Analyzer) Run() ([]Finding, error) {
	files, err := a.files()
	if err != nil {
		return nil, err
	}
	ctx, err := a.buildContext(files)
	if err != nil {
		return nil, err
	}
	var findings []Finding
	for _, file := range files {
		items, err := a.analyzeFile(file, ctx)
		if err != nil {
			return nil, err
		}
		findings = append(findings, items...)
	}
	return findings, nil
}

func (a Analyzer) files() ([]string, error) {
	dirs := []string{a.Config.Src.Modules, a.Config.Src.Classes, a.Config.Src.Forms, a.Config.Src.Workbook, "tests"}
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

type procInfo struct {
	Name       string
	ReturnType string
}

type withInfo struct {
	Target string
	Type   string
}

func (a Analyzer) buildContext(files []string) (analysisContext, error) {
	ctx := analysisContext{standardModulePublicProcs: map[string]bool{}}
	modulesRoot := filepath.Clean(filepath.Join(a.RootDir, a.Config.Src.Modules))
	for _, path := range files {
		if !isStandardModuleFile(path, modulesRoot) {
			continue
		}
		lines, err := readLines(path)
		if err != nil {
			return analysisContext{}, err
		}
		for _, raw := range lines {
			code := strings.TrimSpace(gui.StripComment(raw))
			if m := publicProcDeclRe.FindStringSubmatch(code); len(m) > 0 {
				ctx.standardModulePublicProcs[strings.ToLower(m[1])] = true
			}
		}
	}
	return ctx, nil
}

func (a Analyzer) analyzeFile(path string, ctx analysisContext) ([]Finding, error) {
	lines, err := readLines(path)
	if err != nil {
		return nil, err
	}
	module := strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
	decls := map[string]string{}
	funcReturns := map[string]string{}
	current := procInfo{}
	withStack := make([]withInfo, 0)
	reportedMissingHelpers := map[string]bool{}
	var findings []Finding

	for i, raw := range lines {
		lineNo := i + 1
		code := strings.TrimSpace(gui.StripComment(raw))
		if code == "" {
			continue
		}
		isProcDecl := false
		if m := procRe.FindStringSubmatch(code); len(m) > 0 {
			current = procInfo{Name: m[2], ReturnType: m[3]}
			isProcDecl = true
			if strings.EqualFold(m[1], "function") || strings.Contains(strings.ToLower(m[1]), "property") {
				if isObjectType(m[3]) {
					funcReturns[strings.ToLower(m[2])] = m[3]
				}
			}
		}
		if strings.EqualFold(code, "end sub") || strings.EqualFold(code, "end function") || strings.EqualFold(code, "end property") {
			current = procInfo{}
			withStack = withStack[:0]
			continue
		}
		if endWithRe.MatchString(code) {
			if len(withStack) > 0 {
				withStack = withStack[:len(withStack)-1]
			}
			continue
		}
		if m := withRe.FindStringSubmatch(code); len(m) > 0 {
			withStack = append(withStack, resolveWithInfo(m[1], decls))
			continue
		}
		if m := declRe.FindStringSubmatch(code); len(m) > 0 {
			decls[strings.ToLower(m[1])] = m[2]
			continue
		}
		if !isProcDecl {
			for _, helper := range referencedTraceHelpers(code) {
				lower := strings.ToLower(helper)
				if reportedMissingHelpers[lower] || ctx.standardModulePublicProcs[lower] {
					continue
				}
				if rule, ok := traceHelperDependencies[lower]; ok {
					findings = append(findings, a.helperFinding(path, module, current.Name, lineNo, lines, helper, rule))
					reportedMissingHelpers[lower] = true
				}
			}
		}
		if currentWith, ok := currentWithInfo(withStack); ok {
			if m := withMember.FindStringSubmatch(code); len(m) > 0 {
				if rule, ok := invalidMemberRuleFor(currentWith.Type, m[1]); ok {
					findings = append(findings, a.memberFinding(path, module, current.Name, lineNo, lines, currentWith.Target, currentWith.Type, m[1], rule))
				}
			}
		}
		for _, m := range memberRe.FindAllStringSubmatch(code, -1) {
			if typ, ok := decls[strings.ToLower(m[1])]; ok {
				if rule, ok := invalidMemberRuleFor(typ, m[2]); ok {
					findings = append(findings, a.memberFinding(path, module, current.Name, lineNo, lines, m[1], typ, m[2], rule))
				}
			}
		}
		if setAssign.MatchString(code) {
			continue
		}
		if m := assignRe.FindStringSubmatch(code); len(m) > 0 {
			target := strings.ToLower(m[1])
			if current.Name != "" && strings.EqualFold(target, current.Name) && isObjectType(current.ReturnType) {
				findings = append(findings, a.finding(path, module, current.Name, lineNo, lines, "VBA103", m[1], current.ReturnType))
				continue
			}
			if cm := callAssign.FindStringSubmatch(code); len(cm) > 0 {
				callee := strings.ToLower(lastName(cm[2]))
				if typ, ok := decls[target]; ok && isObjectType(typ) && isObjectType(funcReturns[callee]) {
					findings = append(findings, a.finding(path, module, current.Name, lineNo, lines, "VBA102", m[1], funcReturns[callee]))
					continue
				}
			}
			if typ, ok := decls[target]; ok && isObjectType(typ) {
				findings = append(findings, a.finding(path, module, current.Name, lineNo, lines, "VBA101", m[1], typ))
			}
		}
	}
	return findings, nil
}

func BlockingFindings(findings []Finding) []Finding {
	blocking := make([]Finding, 0)
	for _, finding := range findings {
		if strings.EqualFold(finding.Severity, "error") {
			blocking = append(blocking, finding)
		}
	}
	return blocking
}

func (a Analyzer) finding(path, module, procedure string, line int, lines []string, code, target, typ string) Finding {
	file, err := filepath.Rel(a.RootDir, path)
	if err != nil {
		file = path
	}
	msg := target + " is declared As " + typ + " and is assigned without Set."
	reason := "VBA object references require Set when assigning an object value."
	suggestion := "Use Set " + target + " = ... when the right-hand side returns an object."
	if code == "VBA103" {
		msg = target + " returns As " + typ + " and is assigned without Set."
		suggestion = "Use Set " + target + " = ... when returning an object from this function."
	}
	return Finding{
		Code:       code,
		Severity:   "warning",
		File:       filepath.ToSlash(file),
		Module:     module,
		Procedure:  procedure,
		Line:       line,
		Message:    msg,
		Reason:     reason,
		Suggestion: suggestion,
		NearbyCode: nearby(lines, line, 2),
	}
}

func (a Analyzer) memberFinding(path, module, procedure string, line int, lines []string, target, typ, member string, rule invalidMemberRule) Finding {
	file, err := filepath.Rel(a.RootDir, path)
	if err != nil {
		file = path
	}
	targetLabel := target
	if targetLabel == "" {
		targetLabel = "This object"
	}
	return Finding{
		Code:       "VBA104",
		Severity:   "error",
		File:       filepath.ToSlash(file),
		Module:     module,
		Procedure:  procedure,
		Line:       line,
		Message:    targetLabel + " is declared As " + typ + " but ." + member + " is not a member of " + typ + ".",
		Reason:     rule.Reason,
		Suggestion: rule.Suggestion,
		NearbyCode: nearby(lines, line, 2),
	}
}

func (a Analyzer) helperFinding(path, module, procedure string, line int, lines []string, helper string, rule helperDependencyRule) Finding {
	file, err := filepath.Rel(a.RootDir, path)
	if err != nil {
		file = path
	}
	return Finding{
		Code:       rule.Code,
		Severity:   "error",
		File:       filepath.ToSlash(file),
		Module:     module,
		Procedure:  procedure,
		Line:       line,
		Message:    helper + " is called but no Public standard-module definition was found in source.",
		Reason:     rule.Reason,
		Suggestion: rule.Suggestion,
		NearbyCode: nearby(lines, line, 2),
	}
}

func readLines(path string) ([]string, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	var lines []string
	scanner := bufio.NewScanner(f)
	for scanner.Scan() {
		lines = append(lines, scanner.Text())
	}
	if err := scanner.Err(); err != nil {
		if closeErr := f.Close(); closeErr != nil {
			return nil, closeErr
		}
		return nil, err
	}
	if err := f.Close(); err != nil {
		return nil, err
	}
	return lines, nil
}

func nearby(lines []string, line, radius int) []string {
	start := line - radius
	if start < 1 {
		start = 1
	}
	end := line + radius
	if end > len(lines) {
		end = len(lines)
	}
	out := make([]string, 0, end-start+1)
	for i := start; i <= end; i++ {
		prefix := "  "
		if i == line {
			prefix = "> "
		}
		out = append(out, prefix+strconvItoa(i)+" | "+lines[i-1])
	}
	return out
}

func isObjectType(typ string) bool {
	typ = strings.ToLower(strings.TrimSpace(typ))
	if typ == "" {
		return false
	}
	return objectTypes[typ] || strings.HasSuffix(typ, ".application") || strings.HasSuffix(typ, ".workbook") || strings.HasSuffix(typ, ".worksheet") || strings.HasSuffix(typ, ".range")
}

func lastName(name string) string {
	parts := strings.Split(name, ".")
	return parts[len(parts)-1]
}

func referencedTraceHelpers(code string) []string {
	seen := map[string]bool{}
	helpers := make([]string, 0, 2)
	for _, re := range []*regexp.Regexp{traceHelperQualRe, traceHelperCallRe} {
		for _, m := range re.FindAllStringSubmatch(code, -1) {
			if len(m) < 2 {
				continue
			}
			name := m[1]
			key := strings.ToLower(name)
			if seen[key] {
				continue
			}
			seen[key] = true
			helpers = append(helpers, name)
		}
	}
	return helpers
}

func resolveWithInfo(expr string, decls map[string]string) withInfo {
	expr = strings.TrimSpace(expr)
	if expr == "" {
		return withInfo{}
	}
	base := expr
	if idx := strings.Index(base, "("); idx >= 0 {
		base = base[:idx]
	}
	base = lastName(strings.TrimSpace(base))
	if typ, ok := decls[strings.ToLower(base)]; ok {
		return withInfo{Target: base, Type: typ}
	}
	return withInfo{}
}

func currentWithInfo(stack []withInfo) (withInfo, bool) {
	for i := len(stack) - 1; i >= 0; i-- {
		if stack[i].Type != "" {
			return stack[i], true
		}
	}
	return withInfo{}, false
}

func invalidMemberRuleFor(typ, member string) (invalidMemberRule, bool) {
	rules, ok := invalidObjectMembers[strings.ToLower(strings.TrimSpace(typ))]
	if !ok {
		return invalidMemberRule{}, false
	}
	rule, ok := rules[strings.ToLower(strings.TrimSpace(member))]
	if !ok {
		return invalidMemberRule{}, false
	}
	return rule, true
}

func isStandardModuleFile(path, modulesRoot string) bool {
	if !strings.EqualFold(filepath.Ext(path), ".bas") {
		return false
	}
	path = filepath.Clean(path)
	modulesRoot = filepath.Clean(modulesRoot)
	rel, err := filepath.Rel(modulesRoot, path)
	if err != nil {
		return false
	}
	return rel != "." && rel != ".." && !strings.HasPrefix(rel, ".."+string(filepath.Separator))
}

func strconvItoa(v int) string {
	const digits = "0123456789"
	if v == 0 {
		return "0"
	}
	var buf [20]byte
	i := len(buf)
	for v > 0 {
		i--
		buf[i] = digits[v%10]
		v /= 10
	}
	return string(buf[i:])
}
