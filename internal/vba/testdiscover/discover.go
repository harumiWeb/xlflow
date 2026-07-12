package testdiscover

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
)

type Options struct {
	RootDir string
	Config  config.Config
	Path    string
	Module  string
}

type Result struct {
	Root    string  `json:"root"`
	Summary Summary `json:"summary"`
	Items   []Test  `json:"items"`
}

type Summary struct {
	Files int `json:"files"`
	Tests int `json:"tests"`
}

type DuplicateTestError struct {
	Module     string
	Name       string
	FirstPath  string
	FirstLine  int
	SecondPath string
	SecondLine int
}

func (e DuplicateTestError) Error() string {
	return fmt.Sprintf("duplicate VBA test procedure %s in module %s at %s:%d and %s:%d", e.Name, e.Module, e.FirstPath, e.FirstLine, e.SecondPath, e.SecondLine)
}

type InvalidMetadataError struct {
	Path    string
	Line    int
	Module  string
	Message string
}

func (e InvalidMetadataError) Error() string {
	location := e.Path
	if e.Line > 0 {
		location = fmt.Sprintf("%s:%d", location, e.Line)
	}
	if e.Module != "" {
		return fmt.Sprintf("%s in module %s: %s", location, e.Module, e.Message)
	}
	return fmt.Sprintf("%s: %s", location, e.Message)
}

type Test struct {
	ID            string         `json:"id"`
	Module        string         `json:"module"`
	Name          string         `json:"name"`
	QualifiedName string         `json:"qualified_name"`
	SourcePath    string         `json:"source_path"`
	Line          int            `json:"line"`
	Tags          []string       `json:"tags"`
	StatusHint    string         `json:"status_hint,omitempty"`
	Skip          *StatusReason  `json:"skip,omitempty"`
	Todo          *StatusReason  `json:"todo,omitempty"`
	ExpectedError *ExpectedError `json:"expected_error,omitempty"`
}

type StatusReason struct {
	Reason *string `json:"reason,omitempty"`
}

type ExpectedError struct {
	Number      int     `json:"number"`
	Description *string `json:"description,omitempty"`
	Source      *string `json:"source,omitempty"`
}

var tagLineRE = regexp.MustCompile(`(?i)^'\s*@Tag\s*\("([^"]+)"\)`)
var expectedErrorLineRE = regexp.MustCompile(`(?i)^'\s*@ExpectedError\s*\((.*)\)\s*$`)
var expectedErrorPrefixRE = regexp.MustCompile(`(?i)^'\s*@ExpectedError\b`)
var skipTodoLineRE = regexp.MustCompile(`(?i)^'\s*@(Skip|Todo)(?:\s*\((.*)\))?\s*$`)
var skipPrefixRE = regexp.MustCompile(`(?i)^'\s*@Skip\b`)
var todoPrefixRE = regexp.MustCompile(`(?i)^'\s*@Todo\b`)

func Discover(opts Options) (*Result, error) {
	symbolResult, err := symbols.Inspect(symbols.Options{
		RootDir: opts.RootDir,
		Config:  opts.Config,
		Path:    opts.Path,
		Module:  opts.Module,
	})
	if err != nil {
		return nil, err
	}

	result := &Result{
		Root:  symbolResult.Root,
		Items: []Test{},
	}
	seen := map[string]Test{}
	for _, file := range symbolResult.Files {
		if !strings.EqualFold(file.ModuleKind, "standard") {
			continue
		}
		result.Summary.Files++
		lines, err := readSourceLines(opts.RootDir, file.Path)
		if err != nil {
			return nil, err
		}
		for _, sym := range file.Symbols {
			metadata, err := metadataAbove(lines, sym.StartLine, file.Path, file.ModuleName)
			if err != nil {
				return nil, err
			}
			if !isTestProcedure(sym) {
				if metadata.ExpectedError != nil {
					return nil, InvalidMetadataError{
						Path:    file.Path,
						Line:    metadata.ExpectedErrorLine,
						Module:  file.ModuleName,
						Message: "@ExpectedError annotation is only supported on test procedures",
					}
				}
				continue
			}
			qualifiedName := file.ModuleName + "." + sym.Name
			test := Test{
				ID:            qualifiedName,
				Module:        file.ModuleName,
				Name:          sym.Name,
				QualifiedName: qualifiedName,
				SourcePath:    file.Path,
				Line:          sym.StartLine,
				Tags:          metadata.Tags,
				StatusHint:    metadata.StatusHint(),
				Skip:          metadata.Skip,
				Todo:          metadata.Todo,
				ExpectedError: metadata.ExpectedError,
			}
			key := strings.ToLower(file.ModuleName) + "\x00" + strings.ToLower(sym.Name)
			if previous, ok := seen[key]; ok {
				return nil, DuplicateTestError{
					Module:     file.ModuleName,
					Name:       sym.Name,
					FirstPath:  previous.SourcePath,
					FirstLine:  previous.Line,
					SecondPath: test.SourcePath,
					SecondLine: test.Line,
				}
			}
			seen[key] = test
			result.Items = append(result.Items, test)
		}
	}
	sort.Slice(result.Items, func(i, j int) bool {
		if result.Items[i].SourcePath != result.Items[j].SourcePath {
			return result.Items[i].SourcePath < result.Items[j].SourcePath
		}
		if result.Items[i].Line != result.Items[j].Line {
			return result.Items[i].Line < result.Items[j].Line
		}
		return result.Items[i].QualifiedName < result.Items[j].QualifiedName
	})
	result.Summary.Tests = len(result.Items)
	return result, nil
}

func isTestProcedure(sym symbols.Symbol) bool {
	if sym.Kind != "sub" || len(sym.Parameters) != 0 {
		return false
	}
	name := strings.ToLower(sym.Name)
	return strings.HasPrefix(name, "test") || strings.HasSuffix(name, "_test")
}

func readSourceLines(rootDir, sourcePath string) ([]string, error) {
	path := filepath.FromSlash(sourcePath)
	if !filepath.IsAbs(path) {
		path = filepath.Join(rootDir, path)
	}
	body, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	return strings.Split(strings.ReplaceAll(string(body), "\r\n", "\n"), "\n"), nil
}

type testMetadata struct {
	Tags              []string
	Skip              *StatusReason
	SkipLine          int
	Todo              *StatusReason
	TodoLine          int
	ExpectedError     *ExpectedError
	ExpectedErrorLine int
}

func (m testMetadata) StatusHint() string {
	if m.Skip != nil {
		return "skipped"
	}
	if m.Todo != nil {
		return "todo"
	}
	return ""
}

func metadataAbove(lines []string, startLine int, sourcePath string, module string) (testMetadata, error) {
	metadata := testMetadata{Tags: []string{}}
	if len(lines) == 0 || startLine <= 1 {
		return metadata, nil
	}
	for i := startLine - 2; i >= 0 && i < len(lines); i-- {
		prev := strings.TrimSpace(lines[i])
		if prev == "" {
			continue
		}
		if match := tagLineRE.FindStringSubmatch(prev); match != nil {
			metadata.Tags = append(metadata.Tags, match[1])
			continue
		}
		if expectedErrorPrefixRE.MatchString(prev) {
			if metadata.ExpectedError != nil {
				return metadata, InvalidMetadataError{
					Path:    sourcePath,
					Line:    i + 1,
					Module:  module,
					Message: "multiple @ExpectedError annotations on one test procedure",
				}
			}
			expectedError, err := parseExpectedErrorAnnotation(prev)
			if err != nil {
				return metadata, InvalidMetadataError{
					Path:    sourcePath,
					Line:    i + 1,
					Module:  module,
					Message: err.Error(),
				}
			}
			metadata.ExpectedError = expectedError
			metadata.ExpectedErrorLine = i + 1
			continue
		}
		if skipPrefixRE.MatchString(prev) {
			if metadata.Skip != nil {
				return metadata, InvalidMetadataError{
					Path:    sourcePath,
					Line:    i + 1,
					Module:  module,
					Message: "multiple @Skip annotations on one test procedure",
				}
			}
			if metadata.Todo != nil {
				return metadata, InvalidMetadataError{
					Path:    sourcePath,
					Line:    i + 1,
					Module:  module,
					Message: "test cannot be both skipped and todo",
				}
			}
			status, err := parseSkipTodoAnnotation(prev)
			if err != nil {
				return metadata, InvalidMetadataError{
					Path:    sourcePath,
					Line:    i + 1,
					Module:  module,
					Message: err.Error(),
				}
			}
			metadata.Skip = status
			metadata.SkipLine = i + 1
			continue
		}
		if todoPrefixRE.MatchString(prev) {
			if metadata.Todo != nil {
				return metadata, InvalidMetadataError{
					Path:    sourcePath,
					Line:    i + 1,
					Module:  module,
					Message: "multiple @Todo annotations on one test procedure",
				}
			}
			if metadata.Skip != nil {
				return metadata, InvalidMetadataError{
					Path:    sourcePath,
					Line:    i + 1,
					Module:  module,
					Message: "test cannot be both skipped and todo",
				}
			}
			status, err := parseSkipTodoAnnotation(prev)
			if err != nil {
				return metadata, InvalidMetadataError{
					Path:    sourcePath,
					Line:    i + 1,
					Module:  module,
					Message: err.Error(),
				}
			}
			metadata.Todo = status
			metadata.TodoLine = i + 1
			continue
		}
		if strings.HasPrefix(prev, "''") {
			continue
		}
		break
	}
	return metadata, nil
}

func parseSkipTodoAnnotation(line string) (*StatusReason, error) {
	match := skipTodoLineRE.FindStringSubmatch(line)
	if match == nil {
		return nil, fmt.Errorf("malformed @Skip/@Todo annotation")
	}
	reasonExpr := ""
	if len(match) >= 3 {
		reasonExpr = strings.TrimSpace(match[2])
	}
	if reasonExpr == "" && strings.Contains(line, "(") {
		return nil, fmt.Errorf("malformed @%s reason: expected a quoted string literal", match[1])
	}
	if reasonExpr == "" {
		return &StatusReason{}, nil
	}
	reason, err := parseExpectedErrorStringArg(reasonExpr)
	if err != nil {
		return nil, fmt.Errorf("malformed @%s reason: %w", match[1], err)
	}
	return &StatusReason{Reason: &reason}, nil
}

func parseExpectedErrorAnnotation(line string) (*ExpectedError, error) {
	match := expectedErrorLineRE.FindStringSubmatch(line)
	if match == nil {
		return nil, fmt.Errorf("malformed @ExpectedError annotation")
	}
	args, err := splitExpectedErrorArgs(match[1])
	if err != nil {
		return nil, err
	}
	if len(args) < 1 || len(args) > 3 {
		return nil, fmt.Errorf("@ExpectedError supports 1 to 3 arguments")
	}
	number, err := strconv.Atoi(strings.TrimSpace(args[0]))
	if err != nil {
		return nil, fmt.Errorf("@ExpectedError error number must be numeric")
	}
	expected := &ExpectedError{Number: number}
	if len(args) >= 2 {
		description, err := parseExpectedErrorStringArg(args[1])
		if err != nil {
			return nil, fmt.Errorf("malformed @ExpectedError description: %w", err)
		}
		expected.Description = &description
	}
	if len(args) >= 3 {
		source, err := parseExpectedErrorStringArg(args[2])
		if err != nil {
			return nil, fmt.Errorf("malformed @ExpectedError source: %w", err)
		}
		expected.Source = &source
	}
	return expected, nil
}

func splitExpectedErrorArgs(input string) ([]string, error) {
	args := []string{}
	var current strings.Builder
	inString := false
	for i := 0; i < len(input); i++ {
		ch := input[i]
		if ch == '"' {
			current.WriteByte(ch)
			if inString && i+1 < len(input) && input[i+1] == '"' {
				i++
				current.WriteByte(input[i])
				continue
			}
			inString = !inString
			continue
		}
		if ch == ',' && !inString {
			args = append(args, strings.TrimSpace(current.String()))
			current.Reset()
			continue
		}
		current.WriteByte(ch)
	}
	if inString {
		return nil, fmt.Errorf("malformed string literal")
	}
	args = append(args, strings.TrimSpace(current.String()))
	for _, arg := range args {
		if arg == "" {
			return nil, fmt.Errorf("@ExpectedError arguments must not be empty")
		}
	}
	return args, nil
}

func parseExpectedErrorStringArg(input string) (string, error) {
	input = strings.TrimSpace(input)
	if len(input) < 2 || input[0] != '"' || input[len(input)-1] != '"' {
		return "", fmt.Errorf("expected a quoted string literal")
	}
	body := input[1 : len(input)-1]
	var out strings.Builder
	for i := 0; i < len(body); i++ {
		if body[i] == '"' {
			if i+1 < len(body) && body[i+1] == '"' {
				out.WriteByte('"')
				i++
				continue
			}
			return "", fmt.Errorf("unexpected quote")
		}
		out.WriteByte(body[i])
	}
	return out.String(), nil
}
