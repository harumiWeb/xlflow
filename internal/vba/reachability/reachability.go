// Package reachability builds the conservative root set used by project-wide
// private-procedure analysis.
package reachability

import (
	"os"
	"path/filepath"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/callgraph"
	"github.com/harumiWeb/xlflow/internal/vba/calls"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
	"github.com/harumiWeb/xlflow/internal/vba/testdiscover"
	"github.com/harumiWeb/xlflow/internal/vba/userforms"
)

type Options struct {
	RootDir string
	Config  config.Config
	Symbols *symbols.Result
	Calls   *calls.Result
}

type Result struct {
	Roots []callgraph.Root
	callgraph.ReachabilityResult
}

func Analyze(opts Options) (Result, error) {
	if opts.Calls == nil {
		return Result{}, nil
	}
	roots, err := buildRoots(opts)
	if err != nil {
		return Result{}, err
	}
	result := callgraph.AnalyzeReachability(callgraph.SnapshotFromResult(opts.Calls), callgraph.ReachabilityRequest{Roots: roots})
	return Result{Roots: roots, ReachabilityResult: result}, nil
}

func buildRoots(opts Options) ([]callgraph.Root, error) {
	roots := []callgraph.Root{}
	if entry := strings.TrimSpace(opts.Config.Project.Entry); entry != "" {
		roots = append(roots, callgraph.Root{Target: entry, Confidence: callgraph.RootConfirmed, Reason: "project.entry"})
	}
	if opts.Symbols == nil {
		return roots, nil
	}

	withevents := map[string]map[string]bool{}
	for _, file := range opts.Symbols.Files {
		fields := map[string]bool{}
		for _, sym := range file.Symbols {
			if sym.Kind == "withevents_field" && sym.Parent == "" && strings.TrimSpace(sym.Name) != "" {
				fields[strings.ToLower(sym.Name)] = true
			}
		}
		if len(fields) > 0 {
			withevents[strings.ToLower(file.ModuleName)] = fields
		}
	}

	for _, file := range opts.Symbols.Files {
		for _, sym := range file.Symbols {
			if !procedureSymbolKind(sym.Kind) || strings.TrimSpace(sym.Name) == "" {
				continue
			}
			target := sym.Module + "." + sym.Name
			lowerName := strings.ToLower(sym.Name)
			public := !strings.EqualFold(sym.Visibility, "Private")
			standard := strings.EqualFold(file.ModuleKind, "standard")
			publicMacro := standard && public && sym.Kind == "sub" && len(sym.Parameters) == 0
			testProcedure := standard && public && testdiscover.IsTestProcedure(sym)

			if publicMacro {
				roots = append(roots, callgraph.Root{Target: target, Confidence: callgraph.RootConfirmed, Reason: "public macro"})
			}
			if testProcedure {
				roots = append(roots, callgraph.Root{Target: target, Confidence: callgraph.RootConfirmed, Reason: "test procedure"})
			}
			if standard && public && !publicMacro && !testProcedure {
				roots = append(roots, callgraph.Root{Target: target, Confidence: callgraph.RootPossible, Reason: "public standard-module API"})
			}
			if event, kind := procedureir.ClassifyEvent(file.ModuleKind, sym.Name); event {
				roots = append(roots, callgraph.Root{Target: target, Confidence: callgraph.RootConfirmed, Reason: kind + " event"})
			}
			for field := range withevents[strings.ToLower(file.ModuleName)] {
				if strings.HasPrefix(lowerName, field+"_") {
					roots = append(roots, callgraph.Root{Target: target, Confidence: callgraph.RootConfirmed, Reason: "WithEvents callback"})
				}
			}
		}

		controls, err := controlNames(opts.RootDir, file)
		if err != nil {
			return nil, err
		}
		if len(controls) == 0 {
			continue
		}
		for _, sym := range file.Symbols {
			if !procedureSymbolKind(sym.Kind) {
				continue
			}
			lowerName := strings.ToLower(sym.Name)
			for _, control := range controls {
				if strings.HasPrefix(lowerName, strings.ToLower(control)+"_") {
					roots = append(roots, callgraph.Root{Target: sym.Module + "." + sym.Name, Confidence: callgraph.RootConfirmed, Reason: "UserForm control event"})
					break
				}
			}
		}
	}
	return roots, nil
}

func controlNames(rootDir string, file symbols.FileResult) ([]string, error) {
	if !strings.EqualFold(file.ModuleKind, "form") {
		return nil, nil
	}
	path := file.Path
	if !filepath.IsAbs(path) {
		path = filepath.Join(rootDir, filepath.FromSlash(path))
	}
	if strings.EqualFold(filepath.Ext(path), ".bas") && strings.EqualFold(filepath.Base(filepath.Dir(path)), "code") {
		path = filepath.Join(filepath.Dir(filepath.Dir(path)), strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))+".frm")
	}
	if !strings.EqualFold(filepath.Ext(path), ".frm") {
		return nil, nil
	}
	body, err := os.ReadFile(path)
	if err != nil {
		if os.IsNotExist(err) {
			return nil, nil
		}
		return nil, err
	}
	form := userforms.Parse(string(body))
	controls := make([]string, 0, len(form.Controls))
	for _, control := range form.Controls {
		if strings.TrimSpace(control.Name) != "" {
			controls = append(controls, control.Name)
		}
	}
	return controls, nil
}

func procedureSymbolKind(kind string) bool {
	switch kind {
	case "sub", "function", "property", "property_get", "property_let", "property_set":
		return true
	default:
		return false
	}
}
