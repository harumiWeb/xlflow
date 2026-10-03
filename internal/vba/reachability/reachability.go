// Package reachability builds the conservative root set used by project-wide
// private-procedure analysis.
package reachability

import (
	"context"
	"os"
	"path/filepath"
	"regexp"
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
	// Sources contains source bytes keyed by their project-relative or absolute
	// paths. Project collectors pass the same bytes used to build symbols and
	// procedure IR so root discovery never reopens those files.
	Sources     map[string][]byte
	sourceIndex map[string][]byte
}

type Result struct {
	Roots       []callgraph.Root
	EntryPoints []callgraph.RootResolution
	callgraph.ReachabilityResult
	// Reportable marks call-graph nodes that VB021 may surface. Private and
	// Friend procedures are always candidates; Public procedures qualify only
	// when Option Private Module hides them from the host and the module is
	// not VB_Exposed.
	Reportable map[string]bool
}

func Analyze(opts Options) (Result, error) {
	return analyzeContext(context.Background(), opts, false)
}

// AnalyzeContext is the cancellable form of Analyze. Callers may supply the
// bytes already captured for the project; when absent, source files are read
// from disk as before. It also retains entry-point resolution evidence.
func AnalyzeContext(ctx context.Context, opts Options) (Result, error) {
	return analyzeContext(ctx, opts, true)
}

func analyzeContext(ctx context.Context, opts Options, retainRoots bool) (Result, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return Result{}, err
	}
	opts.sourceIndex = indexCapturedSources(opts)
	if opts.Calls == nil {
		return Result{}, nil
	}
	roots, err := buildRootsContext(ctx, opts)
	if err != nil {
		return Result{}, err
	}
	reportable, err := buildReportableContext(ctx, opts)
	if err != nil {
		return Result{}, err
	}
	if err := ctx.Err(); err != nil {
		return Result{}, err
	}
	snapshot := callgraph.SnapshotFromResult(opts.Calls)
	request := callgraph.ReachabilityRequest{Roots: roots, Reportable: reportable}
	var result callgraph.ReachabilityResult
	var entryPoints []callgraph.RootResolution
	if retainRoots {
		result, entryPoints = callgraph.AnalyzeReachabilityWithRoots(snapshot, request)
	} else {
		result = callgraph.AnalyzeReachability(snapshot, request)
	}
	if err := ctx.Err(); err != nil {
		return Result{}, err
	}
	return Result{Roots: roots, EntryPoints: entryPoints, ReachabilityResult: result, Reportable: reportable}, nil
}

// buildReportable computes which procedures are internal enough to report
// when unreachable. Friend members are project-internal by visibility, and
// Public procedures in an Option Private Module lose the host-facing surface
// unless the module is VB_Exposed.
func buildReportableContext(ctx context.Context, opts Options) (map[string]bool, error) {
	reportable := make(map[string]bool)
	if opts.Symbols == nil {
		return reportable, nil
	}
	for _, file := range opts.Symbols.Files {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		body, err := sourceBytes(ctx, opts, file.Path)
		if err != nil {
			if ctx.Err() != nil {
				return nil, ctx.Err()
			}
			// Missing or unreadable Option Private Module metadata has always
			// failed open, leaving the procedure visible as a possible root.
			body = nil
		}
		privacy := modulePrivacyFacts(file, body)
		for _, sym := range file.Symbols {
			if !procedureSymbolKind(sym.Kind) || strings.TrimSpace(sym.Name) == "" {
				continue
			}
			key := callgraph.ID{
				Module: sym.Module, QualifiedName: sym.Module + "." + sym.Name,
				Kind: sym.Kind, File: sym.File, Line: sym.StartLine, Column: sym.StartColumn,
			}.String()
			switch strings.ToLower(strings.TrimSpace(sym.Visibility)) {
			case "private", "friend":
				reportable[key] = true
			default:
				if privacy.hostHidden {
					reportable[key] = true
				}
			}
		}
	}
	return reportable, ctx.Err()
}

// modulePrivacy describes how a module's public surface interacts with host
// visibility. hostHidden is true when `Option Private Module` applies and the
// module is not re-exposed through the VB_Exposed attribute.
type modulePrivacy struct {
	optionPrivate bool
	exposed       bool
	hostHidden    bool
}

func modulePrivacyFacts(file symbols.FileResult, source []byte) modulePrivacy {
	var privacy modulePrivacy
	// Option Private Module only narrows standard and class modules; document
	// and form modules keep their host-driven surface.
	if !strings.EqualFold(file.ModuleKind, "standard") && !strings.EqualFold(file.ModuleKind, "class") {
		return privacy
	}
	for _, sym := range file.Symbols {
		if sym.Kind != "module" {
			continue
		}
		for _, attribute := range sym.Attributes {
			if strings.EqualFold(attribute.Name, "VB_Exposed") && strings.EqualFold(strings.TrimSpace(attribute.Value), "True") {
				privacy.exposed = true
			}
		}
	}
	privacy.optionPrivate = moduleDeclaresOptionPrivate(source)
	privacy.hostHidden = privacy.optionPrivate && !privacy.exposed
	return privacy
}

var optionPrivateModuleRE = regexp.MustCompile(`(?i)^\s*option\s+private\s+module\b`)

func moduleDeclaresOptionPrivate(source []byte) bool {
	for _, line := range strings.Split(string(source), "\n") {
		if optionPrivateModuleRE.MatchString(line) {
			return true
		}
	}
	return false
}

func buildRootsContext(ctx context.Context, opts Options) ([]callgraph.Root, error) {
	roots := []callgraph.Root{}
	if entry := strings.TrimSpace(opts.Config.Project.Entry); entry != "" {
		roots = append(roots, callgraph.Root{Target: entry, Confidence: callgraph.RootConfirmed, Reason: "project.entry"})
	}
	if opts.Symbols == nil {
		return roots, nil
	}

	withevents := map[string]map[string]bool{}
	for _, file := range opts.Symbols.Files {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
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
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		body, err := sourceBytes(ctx, opts, file.Path)
		if err != nil {
			if ctx.Err() != nil {
				return nil, ctx.Err()
			}
			// Preserve moduleDeclaresOptionPrivate's previous fail-open behavior
			// when captured or on-disk source is unavailable.
			body = nil
		}
		privacy := modulePrivacyFacts(file, body)
		for _, sym := range file.Symbols {
			if !procedureSymbolKind(sym.Kind) || strings.TrimSpace(sym.Name) == "" {
				continue
			}
			target := sym.Module + "." + sym.Name
			lowerName := strings.ToLower(sym.Name)
			visibility := strings.ToLower(strings.TrimSpace(sym.Visibility))
			public := visibility != "private" && visibility != "friend"
			// Friend members and host-hidden public procedures have no
			// host-facing surface; they are reachable only from inside the
			// project or through statically discoverable dynamic references.
			external := public && !privacy.hostHidden
			standard := strings.EqualFold(file.ModuleKind, "standard")
			publicMacro := standard && external && sym.Kind == "sub" && len(sym.Parameters) == 0
			// The generated xlflow test runner is a standard module injected
			// into the same VBA project and invokes Module.TestName and the
			// BeforeAll/AfterAll/BeforeEach/AfterEach hooks with qualified
			// calls. Option Private Module only hides members from other
			// projects and the host macro UI, so a public test or hook stays
			// reachable even in a host-hidden module.
			testProcedure := standard && public && testdiscover.IsTestProcedure(sym)
			testHook := standard && public && sym.Kind == "sub" && testHookName(lowerName)

			if publicMacro {
				roots = append(roots, callgraph.Root{Target: target, Confidence: callgraph.RootConfirmed, Reason: "public macro"})
			}
			if testProcedure {
				roots = append(roots, callgraph.Root{Target: target, Confidence: callgraph.RootConfirmed, Reason: "test procedure"})
			}
			if testHook {
				roots = append(roots, callgraph.Root{Target: target, Confidence: callgraph.RootPossible, Reason: "test runner hook"})
			}
			if standard && external && !publicMacro && !testProcedure && !testHook {
				roots = append(roots, callgraph.Root{Target: target, Confidence: callgraph.RootPossible, Reason: "public standard-module API"})
			}
			// A VB_Exposed class publishes its public members to external
			// clients, so they are possible roots even without a project
			// caller. hostHidden is already false for an exposed module.
			if strings.EqualFold(file.ModuleKind, "class") && public && privacy.exposed {
				roots = append(roots, callgraph.Root{Target: target, Confidence: callgraph.RootPossible, Reason: "exposed class member"})
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

		controls, err := controlNamesContext(ctx, opts, file)
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
	return roots, ctx.Err()
}

func controlNamesContext(ctx context.Context, opts Options, file symbols.FileResult) ([]string, error) {
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	if !strings.EqualFold(file.ModuleKind, "form") {
		return nil, nil
	}
	rootDir := opts.RootDir
	if rootDir == "" {
		rootDir = "."
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
	body, ok := capturedSource(opts, path)
	var err error
	if !ok {
		body, err = os.ReadFile(path)
	}
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

func sourceBytes(ctx context.Context, opts Options, path string) ([]byte, error) {
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	if source, ok := capturedSource(opts, path); ok {
		return source, nil
	}
	if !filepath.IsAbs(path) {
		path = filepath.Join(opts.RootDir, filepath.FromSlash(path))
	}
	return os.ReadFile(path)
}

func capturedSource(opts Options, path string) ([]byte, bool) {
	if source, ok := opts.Sources[path]; ok {
		return source, true
	}
	root := opts.RootDir
	if root == "" {
		root = "."
	}
	target := path
	if !filepath.IsAbs(target) {
		target = filepath.Join(root, filepath.FromSlash(target))
	}
	target, err := filepath.Abs(target)
	if err != nil {
		return nil, false
	}
	source, ok := opts.sourceIndex[capturedSourceKey(target)]
	return source, ok
}

func indexCapturedSources(opts Options) map[string][]byte {
	root := opts.RootDir
	if root == "" {
		root = "."
	}
	index := make(map[string][]byte, len(opts.Sources))
	for candidate, source := range opts.Sources {
		if !filepath.IsAbs(candidate) {
			candidate = filepath.Join(root, filepath.FromSlash(candidate))
		}
		absolute, err := filepath.Abs(candidate)
		if err != nil {
			continue
		}
		index[capturedSourceKey(absolute)] = source
	}
	return index
}

func capturedSourceKey(path string) string {
	return strings.ToLower(filepath.Clean(path))
}

// testHookName reports whether name is one of the fixed per-module hook
// procedures the generated test runner calls (BeforeAll, AfterAll,
// BeforeEach, AfterEach). Hook procedures must be public Subs for the runner
// to bind them.
func testHookName(lowerName string) bool {
	switch lowerName {
	case "beforeall", "afterall", "beforeeach", "aftereach":
		return true
	default:
		return false
	}
}

func procedureSymbolKind(kind string) bool {
	switch kind {
	case "sub", "function", "property", "property_get", "property_let", "property_set":
		return true
	default:
		return false
	}
}
