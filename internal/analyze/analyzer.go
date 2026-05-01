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
	declRe     = regexp.MustCompile(`(?i)^\s*(?:dim|private|public|static)\s+([A-Za-z_][A-Za-z0-9_]*)\s+as\s+(?:new\s+)?([A-Za-z_][A-Za-z0-9_.]*)\b`)
	procRe     = regexp.MustCompile(`(?i)^\s*(?:public\s+|private\s+|friend\s+)?(sub|function|property\s+get)\s+([A-Za-z_][A-Za-z0-9_]*)\b(?:[^']*?\bas\s+([A-Za-z_][A-Za-z0-9_.]*))?`)
	assignRe   = regexp.MustCompile(`(?i)^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=`)
	setAssign  = regexp.MustCompile(`(?i)^\s*set\s+([A-Za-z_][A-Za-z0-9_]*)\s*=`)
	callAssign = regexp.MustCompile(`(?i)^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([A-Za-z_][A-Za-z0-9_.]*)\s*(?:\(|$)`)
)

var objectTypes = map[string]bool{
	"application": true, "workbook": true, "worksheet": true, "range": true,
	"chart": true, "pivot table": true, "pivottable": true, "listobject": true,
	"dictionary": true, "collection": true, "object": true,
}

func (a Analyzer) Run() ([]Finding, error) {
	files, err := a.files()
	if err != nil {
		return nil, err
	}
	var findings []Finding
	for _, file := range files {
		items, err := a.analyzeFile(file)
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

func (a Analyzer) analyzeFile(path string) ([]Finding, error) {
	lines, err := readLines(path)
	if err != nil {
		return nil, err
	}
	module := strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
	decls := map[string]string{}
	funcReturns := map[string]string{}
	current := procInfo{}
	var findings []Finding

	for i, raw := range lines {
		lineNo := i + 1
		code := strings.TrimSpace(gui.StripComment(raw))
		if code == "" {
			continue
		}
		if m := procRe.FindStringSubmatch(code); len(m) > 0 {
			current = procInfo{Name: m[2], ReturnType: m[3]}
			if strings.EqualFold(m[1], "function") || strings.Contains(strings.ToLower(m[1]), "property") {
				if isObjectType(m[3]) {
					funcReturns[strings.ToLower(m[2])] = m[3]
				}
			}
		}
		if strings.EqualFold(code, "end sub") || strings.EqualFold(code, "end function") || strings.EqualFold(code, "end property") {
			current = procInfo{}
			continue
		}
		if m := declRe.FindStringSubmatch(code); len(m) > 0 {
			decls[strings.ToLower(m[1])] = m[2]
			continue
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
