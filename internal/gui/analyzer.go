package gui

import (
	"bufio"
	"os"
	"path/filepath"
	"regexp"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
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
	detect(`(?i)\bapplication\s*\.\s*getopenfilename\b`, "file_picker", "Application.GetOpenFilename", "File picker requires human interaction.", "Pass the path with xlflow run --arg or extract a headless entrypoint that accepts a path."),
	detect(`(?i)\bapplication\s*\.\s*getsaveasfilename\b`, "file_picker", "Application.GetSaveAsFilename", "File picker requires human interaction.", "Pass the path with xlflow run --arg or extract a headless entrypoint that accepts a path."),
	detect(`(?i)\bapplication\s*\.\s*filedialog\b`, "file_picker", "Application.FileDialog", "File dialog requires human interaction.", "Pass the path with xlflow run --arg or keep this code behind an interactive-only adapter."),
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
			if strings.EqualFold(filepath.Base(path), "XlflowUI.bas") && (detector.symbol == "MsgBox" || detector.symbol == "InputBox") {
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
