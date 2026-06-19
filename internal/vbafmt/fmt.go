package vbafmt

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
)

const (
	indentWidth = 4
)

type FileResult struct {
	Path       string
	Changed    bool
	Formatted  string
	Original   string
	Skipped    bool
	SkipReason string
	LineNums   lineNumberFileResult
}

type Result struct {
	Changed         int
	Unchanged       int
	Skipped         int
	Total           int
	ChangedPaths    []string
	SkippedPaths    []string
	SkippedReasons  []SkippedReason
	FormattedByPath map[string]string
	LineNumbers     LineNumberSummary
}

type SkippedReason struct {
	Path   string
	Reason string
}

// FmtOptions controls the format operation.
type FmtOptions struct {
	Write       bool
	Check       bool
	Diff        bool
	Paths       []string
	Root        string
	Cfg         config.Config
	LineNumbers LineNumberMode
}

func Run(opts FmtOptions) (*Result, error) {
	files, err := resolveFiles(opts)
	if err != nil {
		return nil, err
	}

	results := make([]FileResult, 0, len(files))
	for _, path := range files {
		fr, err := formatFile(path, opts.LineNumbers)
		if err != nil {
			return nil, fmt.Errorf("%s: %w", path, err)
		}
		results = append(results, fr)
	}

	return summarizeResults(results, opts)
}

func resolveFiles(opts FmtOptions) ([]string, error) {
	if len(opts.Paths) > 0 {
		return resolveExplicitPaths(opts)
	}
	return resolveProjectFiles(opts)
}

func resolveExplicitPaths(opts FmtOptions) ([]string, error) {
	seen := make(map[string]bool)
	var files []string
	for _, path := range opts.Paths {
		absPath := resolvePath(opts.Root, path)
		info, err := os.Stat(absPath)
		if err != nil {
			return nil, fmt.Errorf("cannot access %s: %w", path, err)
		}
		if info.IsDir() {
			dirFiles, err := collectFiles(absPath)
			if err != nil {
				return nil, err
			}
			for _, f := range dirFiles {
				if seen[f] {
					continue
				}
				seen[f] = true
				files = append(files, f)
			}
		} else {
			if seen[absPath] {
				continue
			}
			seen[absPath] = true
			files = append(files, absPath)
		}
	}
	sort.Strings(files)
	return files, nil
}

func resolvePath(root, path string) string {
	if filepath.IsAbs(path) {
		return path
	}
	return filepath.Join(root, path)
}

func resolveProjectFiles(opts FmtOptions) ([]string, error) {
	formsDir := opts.Cfg.Src.Forms
	if formsDir == "" {
		formsDir = filepath.ToSlash(filepath.Join("src", "forms"))
	}
	userFormCodeSource := opts.Cfg.UserForm.CodeSource
	if userFormCodeSource == "" {
		userFormCodeSource = config.Default().UserForm.CodeSource
	}

	dirs := []string{
		opts.Cfg.Src.Modules,
		opts.Cfg.Src.Classes,
		formsDir,
		opts.Cfg.Src.Workbook,
	}
	if opts.Cfg.Src.Modules == "" {
		dirs[0] = filepath.ToSlash(filepath.Join("src", "modules"))
	}
	if opts.Cfg.Src.Classes == "" {
		dirs[1] = filepath.ToSlash(filepath.Join("src", "classes"))
	}
	if userFormCodeSource == "sidecar" {
		dirs = append(dirs, filepath.ToSlash(filepath.Join(formsDir, "code")))
	}
	if opts.Cfg.Src.Workbook == "" {
		dirs[3] = filepath.ToSlash(filepath.Join("src", "workbook"))
	}
	dirs = append(dirs, "tests")

	seen := make(map[string]bool)
	var files []string
	for _, dir := range dirs {
		root := filepath.Join(opts.Root, dir)
		if _, err := os.Stat(root); os.IsNotExist(err) {
			continue
		}
		if err := filepath.WalkDir(root, func(path string, d os.DirEntry, err error) error {
			if err != nil {
				return err
			}
			if d.IsDir() {
				return nil
			}
			ext := strings.ToLower(filepath.Ext(path))
			if ext != ".bas" && ext != ".cls" {
				return nil
			}
			abs, err := filepath.Abs(path)
			if err != nil {
				return err
			}
			if seen[abs] {
				return nil
			}
			seen[abs] = true
			files = append(files, path)
			return nil
		}); err != nil {
			return nil, err
		}
	}
	sort.Strings(files)
	return files, nil
}

func collectFiles(root string) ([]string, error) {
	seen := make(map[string]bool)
	var files []string
	err := filepath.WalkDir(root, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if d.IsDir() {
			return nil
		}
		ext := strings.ToLower(filepath.Ext(path))
		if ext != ".bas" && ext != ".cls" {
			return nil
		}
		if seen[path] {
			return nil
		}
		seen[path] = true
		files = append(files, path)
		return nil
	})
	if err != nil {
		return nil, err
	}
	return files, nil
}

func formatFile(path string, lineNumbers LineNumberMode) (FileResult, error) {
	ext := strings.ToLower(filepath.Ext(path))
	if ext != ".bas" && ext != ".cls" {
		return FileResult{
			Path:       path,
			Skipped:    true,
			SkipReason: "unsupported extension: " + ext,
		}, nil
	}
	data, err := os.ReadFile(path)
	if err != nil {
		return FileResult{}, err
	}
	isClass := ext == ".cls"
	original := string(data)
	formatted, lineNumResult, err := formatTextDetailed(original, isClass, FormatConfig{LineNumbers: lineNumbers})
	if err != nil {
		if isFormatParseError(err) {
			return FileResult{
				Path:       path,
				Skipped:    true,
				SkipReason: "parse error: " + err.Error(),
			}, nil
		}
		return FileResult{}, err
	}
	return FileResult{
		Path:      path,
		Changed:   formatted != original,
		Formatted: formatted,
		Original:  original,
		LineNums:  lineNumResult,
	}, nil
}

func FormatText(text string, isClass bool) (string, error) {
	return FormatTextWithOptions(text, isClass, FormatConfig{})
}

func FormatTextWithOptions(text string, isClass bool, cfg FormatConfig) (string, error) {
	formatted, _, err := formatTextDetailed(text, isClass, cfg)
	return formatted, err
}

func splitLines(text string) []string {
	text = strings.ReplaceAll(text, "\r\n", "\n")
	text = strings.ReplaceAll(text, "\r", "\n")
	lines := strings.Split(text, "\n")
	clean := make([]string, 0, len(lines))
	clean = append(clean, lines...)
	if len(clean) > 0 && clean[len(clean)-1] == "" {
		clean = clean[:len(clean)-1]
	}
	return clean
}

func isBlankLine(line string) bool {
	return strings.TrimSpace(line) == ""
}

func isVBACommentLine(line string) bool {
	trimmed := strings.TrimLeft(line, " \t")
	if strings.HasPrefix(trimmed, "'") {
		return true
	}
	upper := strings.ToUpper(trimmed)
	return strings.HasPrefix(upper, "REM") && (len(upper) == 3 || upper[3] == ' ')
}

func isClassHeaderLine(line string) bool {
	trimmed := strings.TrimLeft(line, " \t")
	upper := strings.ToUpper(trimmed)
	if strings.HasPrefix(upper, "ATTRIBUTE VB_") {
		return true
	}
	if strings.HasPrefix(upper, "VERSION ") {
		return true
	}
	if upper == "BEGIN" || upper == "END" {
		return true
	}
	if strings.HasPrefix(upper, "MULTIUSE ") {
		return true
	}
	return false
}

func classifyLine(trimmedLine string) (keyword string, isStructural bool) {
	upper := strings.ToUpper(trimmedLine)

	// Single-line If (e.g. "If x Then y = 1") must not open a block.
	if strings.HasPrefix(strings.TrimLeft(upper, " \t"), "IF") {
		stripped := stripTrailingComment(upper)
		stripped = strings.TrimRight(stripped, " \t")
		if strings.HasSuffix(stripped, "_") {
			stripped = strings.TrimRight(strings.TrimSuffix(stripped, "_"), " \t")
		}
		if !strings.HasSuffix(stripped, "THEN") {
			return "", false
		}
	}

	for _, kw := range indentKeywords {
		if matchKeywordStartsWith(upper, kw) {
			if lineContainsDedentForIndent(upper, kw) {
				return "", false
			}
			return kw, true
		}
	}

	for _, kw := range dedentKeywords {
		if matchKeywordStartsWith(upper, kw) {
			return kw, true
		}
	}

	return "", false
}

// stripTrailingComment removes a trailing VBA comment (' or Rem) from a line,
// respecting doubled quote escapes in string literals.
func stripTrailingComment(line string) string {
	inString := false
	for i := 0; i < len(line); i++ {
		ch := line[i]
		if ch == '"' {
			inString = !inString
			continue
		}
		if inString {
			continue
		}
		if ch == '\'' {
			return line[:i]
		}
		if i+3 <= len(line) && strings.ToUpper(line[i:i+3]) == "REM" {
			if i+3 == len(line) || line[i+3] == ' ' {
				return line[:i]
			}
		}
	}
	return line
}

var vbaModifiers = []string{"PUBLIC", "PRIVATE", "FRIEND", "STATIC"}

func matchKeywordStartsWith(upper, keyword string) bool {
	upper = strings.TrimLeft(upper, " \t")
	upper = stripModifiers(upper)
	return strings.HasPrefix(upper, keyword+" ") || upper == keyword
}

// lineContainsDedentForIndent checks whether a colon-separated statement
// contains both an indent keyword and its matching dedent keyword on the same
// line (e.g. "Do While i < 1: i = i + 1: Loop"). When this returns true the
// line should not increase indentation because the block is fully closed.
func lineContainsDedentForIndent(upper, indentKW string) bool {
	parts := strings.Split(upper, ":")
	if len(parts) < 2 {
		return false
	}
	for _, dk := range dedentKeywords {
		if !isMatchingPair(indentKW, dk) {
			continue
		}
		for _, part := range parts[1:] {
			if matchKeywordStartsWith(part, dk) {
				return true
			}
		}
	}
	return false
}

func isMatchingPair(indentKW, dedentKW string) bool {
	switch indentKW {
	case "DO", "DO WHILE", "DO UNTIL":
		return dedentKW == "LOOP"
	case "FOR", "FOR EACH":
		return dedentKW == "NEXT"
	case "WHILE":
		return dedentKW == "WEND"
	case "WITH":
		return dedentKW == "END WITH"
	case "SELECT CASE":
		return dedentKW == "END SELECT"
	case "IF":
		return dedentKW == "END IF"
	case "ENUM":
		return dedentKW == "END ENUM"
	case "TYPE":
		return dedentKW == "END TYPE"
	case "SUB":
		return dedentKW == "END SUB"
	case "FUNCTION":
		return dedentKW == "END FUNCTION"
	case "PROPERTY GET", "PROPERTY LET", "PROPERTY SET":
		return dedentKW == "END PROPERTY"
	}
	return false
}

func stripModifiers(upper string) string {
	for {
		found := false
		for _, mod := range vbaModifiers {
			if strings.HasPrefix(upper, mod+" ") {
				upper = strings.TrimLeft(upper[len(mod)+1:], " ")
				found = true
				break
			}
		}
		if !found {
			break
		}
	}
	return upper
}

var indentKeywords = []string{
	"CASE ELSE",
	"DO UNTIL",
	"DO WHILE",
	"FOR EACH",
	"PROPERTY GET",
	"PROPERTY LET",
	"PROPERTY SET",
	"SELECT CASE",
	"CASE",
	"DO",
	"ELSE",
	"ELSEIF",
	"ENUM",
	"FOR",
	"FUNCTION",
	"IF",
	"SUB",
	"TYPE",
	"WHILE",
	"WITH",
}

var dedentKeywords = []string{
	"END SUB",
	"END FUNCTION",
	"END PROPERTY",
	"LOOP WHILE",
	"LOOP UNTIL",
	"END SELECT",
	"END WITH",
	"END TYPE",
	"END ENUM",
	"END IF",
	"CASE ELSE",
	"ELSEIF",
	"ELSE",
	"CASE",
	"LOOP",
	"NEXT",
	"WEND",
}

func isOptionExplicitGap(prev string) bool {
	prevTrim := strings.TrimSpace(prev)
	return strings.EqualFold(prevTrim, "Option Explicit")
}

func isProcedureGap(prev, current string) bool {
	prevUpper := strings.ToUpper(strings.TrimSpace(prev))
	currentUpper := strings.ToUpper(strings.TrimSpace(current))

	prevIsProcEnd := matchKeywordStartsWith(prevUpper, "END SUB") ||
		matchKeywordStartsWith(prevUpper, "END FUNCTION") ||
		matchKeywordStartsWith(prevUpper, "END PROPERTY")

	currentIsProcStart := matchKeywordStartsWith(currentUpper, "SUB") ||
		matchKeywordStartsWith(currentUpper, "FUNCTION") ||
		matchKeywordStartsWith(currentUpper, "PROPERTY GET") ||
		matchKeywordStartsWith(currentUpper, "PROPERTY LET") ||
		matchKeywordStartsWith(currentUpper, "PROPERTY SET")

	return prevIsProcEnd && currentIsProcStart
}

func summarizeResults(results []FileResult, opts FmtOptions) (*Result, error) {
	r := &Result{
		Total:           len(results),
		FormattedByPath: make(map[string]string),
		LineNumbers: LineNumberSummary{
			Mode: normalizeLineNumberMode(opts.LineNumbers),
		},
	}
	for _, fr := range results {
		if fr.Skipped {
			r.Skipped++
			r.SkippedPaths = append(r.SkippedPaths, fr.Path)
			r.SkippedReasons = append(r.SkippedReasons, SkippedReason{
				Path:   fr.Path,
				Reason: fr.SkipReason,
			})
			continue
		}
		if fr.Changed {
			r.Changed++
			r.ChangedPaths = append(r.ChangedPaths, fr.Path)
			r.FormattedByPath[fr.Path] = fr.Formatted
			if opts.Write {
				if err := os.WriteFile(fr.Path, []byte(fr.Formatted), 0644); err != nil {
					return nil, fmt.Errorf("write %s: %w", fr.Path, err)
				}
			}
		} else {
			r.Unchanged++
		}
		if fr.LineNums.Changed {
			r.LineNumbers.FilesChanged++
		}
		r.LineNumbers.LinesAdded += fr.LineNums.LinesAdded
		r.LineNumbers.LinesRemoved += fr.LineNums.LinesRemoved
		r.LineNumbers.LinesRenumbered += fr.LineNums.LinesRenumbered
		for _, warning := range fr.LineNums.Warnings {
			warning.Path = fr.Path
			r.LineNumbers.Warnings = append(r.LineNumbers.Warnings, warning)
		}
	}
	return r, nil
}

// Diff returns a unified diff for a single file.
func Diff(path, original, formatted string) string {
	if original == formatted {
		return ""
	}
	var buf bytes.Buffer
	fmt.Fprintf(&buf, "--- a/%s\n", path)
	fmt.Fprintf(&buf, "+++ b/%s\n", path)
	origLines := strings.Split(original, "\n")
	fmtLines := strings.Split(formatted, "\n")

	ctx := 3
	diffLines := computeLineDiff(origLines, fmtLines, ctx)
	for _, dl := range diffLines {
		buf.WriteString(dl)
		buf.WriteByte('\n')
	}
	return buf.String()
}

func computeLineDiff(a, b []string, context int) []string {
	type edit struct {
		kind string
		line string
		oldN int
		newN int
	}

	m := len(a)
	n := len(b)

	dp := make([][]int, m+1)
	for i := range dp {
		dp[i] = make([]int, n+1)
	}
	for i := 0; i <= m; i++ {
		dp[i][0] = i
	}
	for j := 0; j <= n; j++ {
		dp[0][j] = j
	}
	for i := 1; i <= m; i++ {
		for j := 1; j <= n; j++ {
			if a[i-1] == b[j-1] {
				dp[i][j] = dp[i-1][j-1]
			} else {
				del := dp[i-1][j] + 1
				ins := dp[i][j-1] + 1
				rep := dp[i-1][j-1] + 1
				minVal := del
				if ins < minVal {
					minVal = ins
				}
				if rep < minVal {
					minVal = rep
				}
				dp[i][j] = minVal
			}
		}
	}

	edits := make([]edit, 0)
	i, j := m, n
	for i > 0 || j > 0 {
		if i > 0 && j > 0 && a[i-1] == b[j-1] {
			edits = append(edits, edit{kind: " ", line: a[i-1], oldN: i, newN: j})
			i--
			j--
		} else if i > 0 && j > 0 && dp[i][j] == dp[i-1][j-1]+1 {
			edits = append(edits, edit{kind: "-", line: a[i-1], oldN: i})
			edits = append(edits, edit{kind: "+", line: b[j-1], newN: j})
			i--
			j--
		} else if i > 0 && dp[i][j] == dp[i-1][j]+1 {
			edits = append(edits, edit{kind: "-", line: a[i-1], oldN: i})
			i--
		} else {
			edits = append(edits, edit{kind: "+", line: b[j-1], newN: j})
			j--
		}
	}

	for left, right := 0, len(edits)-1; left < right; left, right = left+1, right-1 {
		edits[left], edits[right] = edits[right], edits[left]
	}

	var result []string
	for idx := 0; idx < len(edits); {
		chunkStart := idx
		for idx < len(edits) && edits[idx].kind == " " {
			idx++
		}
		if idx == len(edits) {
			break
		}
		for idx < len(edits) && edits[idx].kind != " " {
			idx++
		}

		hunkStart := chunkStart - context
		if hunkStart < 0 {
			hunkStart = 0
		}
		hunkEnd := idx + context
		if hunkEnd > len(edits) {
			hunkEnd = len(edits)
		}

		firstOld := 0
		firstNew := 0
		if hunkStart < len(edits) {
			if edits[hunkStart].oldN > 0 {
				firstOld = edits[hunkStart].oldN
			}
			if edits[hunkStart].newN > 0 {
				firstNew = edits[hunkStart].newN
			}
		}

		oldCount := 0
		newCount := 0
		for k := hunkStart; k < hunkEnd; k++ {
			if edits[k].kind == "-" || edits[k].kind == " " {
				oldCount++
			}
			if edits[k].kind == "+" || edits[k].kind == " " {
				newCount++
			}
		}

		result = append(result, fmt.Sprintf("@@ -%d,%d +%d,%d @@", firstOld, oldCount, firstNew, newCount))
		for k := hunkStart; k < hunkEnd; k++ {
			result = append(result, edits[k].kind+edits[k].line)
		}
	}

	return result
}
