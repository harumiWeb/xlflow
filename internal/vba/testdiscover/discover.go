package testdiscover

import (
	"os"
	"path/filepath"
	"regexp"
	"sort"
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

type Test struct {
	Module        string   `json:"module"`
	Name          string   `json:"name"`
	QualifiedName string   `json:"qualified_name"`
	SourcePath    string   `json:"source_path"`
	Line          int      `json:"line"`
	Tags          []string `json:"tags"`
}

var tagLineRE = regexp.MustCompile(`(?i)^'\s*@Tag\s*\("([^"]+)"\)`)

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
			if !isTestProcedure(sym) {
				continue
			}
			test := Test{
				Module:        file.ModuleName,
				Name:          sym.Name,
				QualifiedName: file.ModuleName + "." + sym.Name,
				SourcePath:    file.Path,
				Line:          sym.StartLine,
				Tags:          tagsAbove(lines, sym.StartLine),
			}
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

func tagsAbove(lines []string, startLine int) []string {
	if len(lines) == 0 || startLine <= 1 {
		return []string{}
	}
	tags := make([]string, 0)
	for i := startLine - 2; i >= 0 && i < len(lines); i-- {
		prev := strings.TrimSpace(lines[i])
		if prev == "" {
			continue
		}
		if match := tagLineRE.FindStringSubmatch(prev); match != nil {
			tags = append(tags, match[1])
			continue
		}
		if strings.HasPrefix(prev, "''") {
			continue
		}
		break
	}
	return tags
}
