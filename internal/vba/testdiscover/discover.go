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

type InvalidTestCaseError struct {
	Path    string
	Line    int
	Module  string
	Message string
}

func (e InvalidTestCaseError) Error() string {
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
	ID             string         `json:"id"`
	Module         string         `json:"module"`
	Name           string         `json:"name"`
	QualifiedName  string         `json:"qualified_name"`
	CaseID         string         `json:"case_id,omitempty"`
	QualifiedProc  string         `json:"qualified_procedure,omitempty"`
	Arguments      []TestArgument `json:"arguments,omitempty"`
	SourcePath     string         `json:"source_path"`
	Line           int            `json:"line"`
	AnnotationLine int            `json:"annotation_line,omitempty"`
	ProcedureLine  int            `json:"procedure_line,omitempty"`
	Tags           []string       `json:"tags"`
	StatusHint     string         `json:"status_hint,omitempty"`
	Skip           *StatusReason  `json:"skip,omitempty"`
	Todo           *StatusReason  `json:"todo,omitempty"`
	ExpectedError  *ExpectedError `json:"expected_error,omitempty"`
}

type TestArgument struct {
	Type  string `json:"type"`
	Value any    `json:"value"`
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
var testCaseLineRE = regexp.MustCompile(`(?i)^'\s*@TestCase\s*\((.*)\)\s*$`)
var testCasePrefixRE = regexp.MustCompile(`(?i)^'\s*@TestCase\b`)

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
			if !isNamedTestProcedure(sym) {
				if metadata.ExpectedError != nil {
					return nil, InvalidMetadataError{
						Path:    file.Path,
						Line:    metadata.ExpectedErrorLine,
						Module:  file.ModuleName,
						Message: "@ExpectedError annotation is only supported on test procedures",
					}
				}
				if len(metadata.TestCases) > 0 {
					return nil, InvalidTestCaseError{
						Path:    file.Path,
						Line:    metadata.TestCases[0].Line,
						Module:  file.ModuleName,
						Message: "@TestCase annotation is only supported on test procedures",
					}
				}
				continue
			}
			tests, err := expandTestCases(file.Path, file.ModuleName, sym, metadata)
			if err != nil {
				return nil, err
			}
			for _, test := range tests {
				key := strings.ToLower(test.ID)
				if previous, ok := seen[key]; ok {
					if strings.EqualFold(previous.Module, test.Module) && strings.EqualFold(previous.Name, test.Name) && previous.ID == previous.QualifiedProc && test.ID == test.QualifiedProc {
						return nil, DuplicateTestError{
							Module:     file.ModuleName,
							Name:       sym.Name,
							FirstPath:  previous.SourcePath,
							FirstLine:  previous.Line,
							SecondPath: test.SourcePath,
							SecondLine: test.Line,
						}
					}
					return nil, InvalidTestCaseError{
						Path:    test.SourcePath,
						Line:    firstPositive(test.AnnotationLine, test.Line),
						Module:  test.Module,
						Message: fmt.Sprintf("duplicate generated test case id %s", test.ID),
					}
				}
				seen[key] = test
				result.Items = append(result.Items, test)
			}
		}
	}
	sort.Slice(result.Items, func(i, j int) bool {
		if result.Items[i].SourcePath != result.Items[j].SourcePath {
			return result.Items[i].SourcePath < result.Items[j].SourcePath
		}
		iLine := result.Items[i].Line
		if result.Items[i].AnnotationLine > 0 {
			iLine = result.Items[i].AnnotationLine
		}
		jLine := result.Items[j].Line
		if result.Items[j].AnnotationLine > 0 {
			jLine = result.Items[j].AnnotationLine
		}
		if iLine != jLine {
			return iLine < jLine
		}
		return result.Items[i].QualifiedName < result.Items[j].QualifiedName
	})
	result.Summary.Tests = len(result.Items)
	return result, nil
}

func isNamedTestProcedure(sym symbols.Symbol) bool {
	if sym.Kind != "sub" {
		return false
	}
	name := strings.ToLower(sym.Name)
	return strings.HasPrefix(name, "test") || strings.HasSuffix(name, "_test")
}

func firstPositive(values ...int) int {
	for _, value := range values {
		if value > 0 {
			return value
		}
	}
	return 0
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
	TestCases         []parsedTestCase
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
		if testCasePrefixRE.MatchString(prev) {
			testCase, err := parseTestCaseAnnotation(prev)
			if err != nil {
				return metadata, InvalidTestCaseError{
					Path:    sourcePath,
					Line:    i + 1,
					Module:  module,
					Message: err.Error(),
				}
			}
			testCase.Line = i + 1
			metadata.TestCases = append(metadata.TestCases, testCase)
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
	for i, j := 0, len(metadata.TestCases)-1; i < j; i, j = i+1, j-1 {
		metadata.TestCases[i], metadata.TestCases[j] = metadata.TestCases[j], metadata.TestCases[i]
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

type parsedTestCase struct {
	Name      string
	HasName   bool
	Arguments []testLiteral
	Line      int
}

type testLiteral struct {
	Kind      string
	Canonical string
	Value     any
}

func expandTestCases(sourcePath, module string, sym symbols.Symbol, metadata testMetadata) ([]Test, error) {
	qualifiedProc := module + "." + sym.Name
	if len(sym.Parameters) == 0 {
		if len(metadata.TestCases) > 0 {
			first := metadata.TestCases[0]
			if first.HasName || len(first.Arguments) > 0 {
				return nil, InvalidTestCaseError{Path: sourcePath, Line: first.Line, Module: module, Message: "parameterless test must not declare @TestCase arguments"}
			}
		}
		return []Test{{
			ID:            qualifiedProc,
			Module:        module,
			Name:          sym.Name,
			QualifiedName: qualifiedProc,
			QualifiedProc: qualifiedProc,
			SourcePath:    sourcePath,
			Line:          sym.StartLine,
			ProcedureLine: sym.StartLine,
			Tags:          metadata.Tags,
			StatusHint:    metadata.StatusHint(),
			Skip:          metadata.Skip,
			Todo:          metadata.Todo,
			ExpectedError: metadata.ExpectedError,
		}}, nil
	}
	if len(metadata.TestCases) == 0 {
		return nil, InvalidTestCaseError{Path: sourcePath, Line: sym.StartLine, Module: module, Message: fmt.Sprintf("parameterized test %s requires at least one @TestCase", sym.Name)}
	}
	params, err := validateTestParameters(sym.Parameters)
	if err != nil {
		return nil, InvalidTestCaseError{Path: sourcePath, Line: sym.StartLine, Module: module, Message: err.Error()}
	}
	tests := make([]Test, 0, len(metadata.TestCases))
	for _, tc := range metadata.TestCases {
		if len(tc.Arguments) != len(params) {
			return nil, InvalidTestCaseError{Path: sourcePath, Line: tc.Line, Module: module, Message: fmt.Sprintf("@TestCase provides %d arguments, but %s requires %d", len(tc.Arguments), sym.Name, len(params))}
		}
		caseID := ""
		if tc.HasName {
			caseID = tc.Name
		} else {
			parts := make([]string, 0, len(tc.Arguments))
			for _, arg := range tc.Arguments {
				parts = append(parts, arg.Canonical)
			}
			caseID = strings.Join(parts, ",")
		}
		args := make([]TestArgument, 0, len(tc.Arguments))
		for i, lit := range tc.Arguments {
			if err := validateLiteralForType(lit, params[i]); err != nil {
				return nil, InvalidTestCaseError{Path: sourcePath, Line: tc.Line, Module: module, Message: fmt.Sprintf("argument %d for %s: %s", i+1, sym.Name, err)}
			}
			args = append(args, TestArgument{Type: params[i].Type, Value: lit.Value})
		}
		id := qualifiedProc + "[" + caseID + "]"
		tests = append(tests, Test{
			ID:             id,
			Module:         module,
			Name:           sym.Name,
			QualifiedName:  id,
			CaseID:         caseID,
			QualifiedProc:  qualifiedProc,
			Arguments:      args,
			SourcePath:     sourcePath,
			Line:           sym.StartLine,
			AnnotationLine: tc.Line,
			ProcedureLine:  sym.StartLine,
			Tags:           metadata.Tags,
			StatusHint:     metadata.StatusHint(),
			Skip:           metadata.Skip,
			Todo:           metadata.Todo,
			ExpectedError:  metadata.ExpectedError,
		})
	}
	return tests, nil
}

func validateTestParameters(params []symbols.Parameter) ([]symbols.Parameter, error) {
	out := make([]symbols.Parameter, 0, len(params))
	for _, param := range params {
		if param.Optional {
			return nil, fmt.Errorf("optional parameters are not supported in parameterized tests")
		}
		if param.ParamArray {
			return nil, fmt.Errorf("ParamArray parameters are not supported in parameterized tests")
		}
		if !strings.EqualFold(param.Passing, "ByVal") {
			return nil, fmt.Errorf("parameter %s must be ByVal", param.Name)
		}
		param.Type = normalizeParameterType(param.Type)
		if param.Type == "" {
			param.Type = "Variant"
		}
		if !isSupportedParameterType(param.Type) {
			return nil, fmt.Errorf("unsupported parameter type %s for %s", param.Type, param.Name)
		}
		out = append(out, param)
	}
	return out, nil
}

func normalizeParameterType(input string) string {
	input = strings.TrimSpace(input)
	if input == "" {
		return ""
	}
	for _, typ := range []string{"Boolean", "Byte", "Integer", "Long", "LongLong", "LongPtr", "Single", "Double", "Currency", "Date", "String", "Variant"} {
		if strings.EqualFold(input, typ) {
			return typ
		}
	}
	return input
}

func isSupportedParameterType(typ string) bool {
	switch strings.ToLower(typ) {
	case "boolean", "byte", "integer", "long", "longlong", "longptr", "single", "double", "currency", "date", "string", "variant":
		return true
	default:
		return false
	}
}

func validateLiteralForType(lit testLiteral, param symbols.Parameter) error {
	typ := strings.ToLower(param.Type)
	if typ == "variant" {
		return nil
	}
	if lit.Kind == "empty" || lit.Kind == "null" {
		return nil
	}
	switch typ {
	case "boolean":
		if lit.Kind != "boolean" {
			return fmt.Errorf("literal %s cannot be passed to Boolean", lit.Canonical)
		}
	case "byte", "integer", "long", "longlong", "longptr":
		if lit.Kind != "integer" {
			return fmt.Errorf("literal %s cannot be passed to %s", lit.Canonical, param.Type)
		}
	case "single", "double", "currency":
		if lit.Kind != "integer" && lit.Kind != "float" {
			return fmt.Errorf("literal %s cannot be passed to %s", lit.Canonical, param.Type)
		}
	case "date":
		if lit.Kind != "date" {
			return fmt.Errorf("literal %s cannot be passed to Date", lit.Canonical)
		}
	case "string":
		if lit.Kind != "string" {
			return fmt.Errorf("literal %s cannot be passed to String", lit.Canonical)
		}
	}
	return nil
}

func parseTestCaseAnnotation(line string) (parsedTestCase, error) {
	match := testCaseLineRE.FindStringSubmatch(line)
	if match == nil {
		return parsedTestCase{}, fmt.Errorf("malformed @TestCase annotation")
	}
	body := strings.TrimSpace(match[1])
	nameExpr, argsExpr, hasName, err := splitTestCaseName(body)
	if err != nil {
		return parsedTestCase{}, err
	}
	tc := parsedTestCase{HasName: hasName}
	if hasName {
		name, err := parseExpectedErrorStringArg(nameExpr)
		if err != nil {
			return parsedTestCase{}, fmt.Errorf("malformed @TestCase name: %w", err)
		}
		if strings.TrimSpace(name) == "" {
			return parsedTestCase{}, fmt.Errorf("@TestCase name must not be empty")
		}
		tc.Name = canonicalCaseName(name)
	}
	if strings.TrimSpace(argsExpr) == "" {
		return tc, nil
	}
	parts, err := splitAnnotationArgs(argsExpr, "@TestCase")
	if err != nil {
		return parsedTestCase{}, err
	}
	for _, part := range parts {
		lit, err := parseTestLiteral(part)
		if err != nil {
			return parsedTestCase{}, err
		}
		tc.Arguments = append(tc.Arguments, lit)
	}
	return tc, nil
}

func splitTestCaseName(input string) (nameExpr, argsExpr string, hasName bool, err error) {
	inString := false
	inDate := false
	for i := 0; i < len(input); i++ {
		ch := input[i]
		if ch == '"' && !inDate {
			if inString && i+1 < len(input) && input[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
			continue
		}
		if ch == '#' && !inString {
			inDate = !inDate
			continue
		}
		if ch == ';' && !inString && !inDate {
			return strings.TrimSpace(input[:i]), strings.TrimSpace(input[i+1:]), true, nil
		}
	}
	if inString {
		return "", "", false, fmt.Errorf("malformed string literal")
	}
	if inDate {
		return "", "", false, fmt.Errorf("malformed date literal")
	}
	return "", input, false, nil
}

func splitAnnotationArgs(input string, annotationName string) ([]string, error) {
	args := []string{}
	var current strings.Builder
	inString := false
	inDate := false
	for i := 0; i < len(input); i++ {
		ch := input[i]
		if ch == '"' && !inDate {
			current.WriteByte(ch)
			if inString && i+1 < len(input) && input[i+1] == '"' {
				i++
				current.WriteByte(input[i])
				continue
			}
			inString = !inString
			continue
		}
		if ch == '#' && !inString {
			inDate = !inDate
			current.WriteByte(ch)
			continue
		}
		if ch == ',' && !inString && !inDate {
			args = append(args, strings.TrimSpace(current.String()))
			current.Reset()
			continue
		}
		current.WriteByte(ch)
	}
	if inString {
		return nil, fmt.Errorf("malformed string literal")
	}
	if inDate {
		return nil, fmt.Errorf("malformed date literal")
	}
	args = append(args, strings.TrimSpace(current.String()))
	for _, arg := range args {
		if arg == "" {
			return nil, fmt.Errorf("%s arguments must not be empty", annotationName)
		}
	}
	return args, nil
}

func parseTestLiteral(input string) (testLiteral, error) {
	raw := strings.TrimSpace(input)
	if raw == "" {
		return testLiteral{}, fmt.Errorf("@TestCase arguments must not be empty")
	}
	if strings.HasPrefix(raw, "\"") {
		value, err := parseExpectedErrorStringArg(raw)
		if err != nil {
			return testLiteral{}, fmt.Errorf("malformed string literal: %w", err)
		}
		return testLiteral{Kind: "string", Canonical: canonicalStringLiteral(value), Value: value}, nil
	}
	if strings.HasPrefix(raw, "#") {
		if len(raw) < 2 || raw[len(raw)-1] != '#' {
			return testLiteral{}, fmt.Errorf("malformed date literal")
		}
		value := strings.TrimSpace(raw[1 : len(raw)-1])
		if value == "" {
			return testLiteral{}, fmt.Errorf("malformed date literal")
		}
		return testLiteral{Kind: "date", Canonical: "#" + value + "#", Value: value}, nil
	}
	switch {
	case strings.EqualFold(raw, "True"):
		return testLiteral{Kind: "boolean", Canonical: "True", Value: true}, nil
	case strings.EqualFold(raw, "False"):
		return testLiteral{Kind: "boolean", Canonical: "False", Value: false}, nil
	case strings.EqualFold(raw, "Empty"):
		return testLiteral{Kind: "empty", Canonical: "Empty", Value: "Empty"}, nil
	case strings.EqualFold(raw, "Null"):
		return testLiteral{Kind: "null", Canonical: "Null", Value: nil}, nil
	}
	numeric := raw
	suffix := byte(0)
	if len(raw) > 0 && strings.ContainsRune("#!@&^%", rune(raw[len(raw)-1])) {
		suffix = raw[len(raw)-1]
		numeric = raw[:len(raw)-1]
	}
	if integerLiteralRE.MatchString(numeric) {
		value, err := strconv.ParseInt(numeric, 10, 64)
		if err != nil {
			return testLiteral{}, fmt.Errorf("integer literal %s is out of range", raw)
		}
		if suffix == '#' || suffix == '!' || suffix == '@' {
			return testLiteral{Kind: "float", Canonical: raw, Value: float64(value)}, nil
		}
		return testLiteral{Kind: "integer", Canonical: raw, Value: value}, nil
	}
	if floatLiteralRE.MatchString(numeric) {
		value, err := strconv.ParseFloat(numeric, 64)
		if err != nil {
			return testLiteral{}, fmt.Errorf("floating-point literal %s is invalid", raw)
		}
		return testLiteral{Kind: "float", Canonical: raw, Value: value}, nil
	}
	return testLiteral{}, fmt.Errorf("unsupported @TestCase literal %s", raw)
}

var integerLiteralRE = regexp.MustCompile(`^[+-]?\d+$`)
var floatLiteralRE = regexp.MustCompile(`^[+-]?(?:\d+\.\d*|\.\d+|\d+[eE][+-]?\d+|\d+\.\d*[eE][+-]?\d+|\.\d+[eE][+-]?\d+)$`)

func canonicalStringLiteral(value string) string {
	return `"` + strings.ReplaceAll(value, `"`, `""`) + `"`
}

func canonicalCaseName(value string) string {
	value = strings.TrimSpace(value)
	value = strings.ReplaceAll(value, "[", "_")
	value = strings.ReplaceAll(value, "]", "_")
	return value
}
