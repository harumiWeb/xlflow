package lspserver

import (
	"context"
	"errors"
	"fmt"
	"io"
	"log"
	"net/url"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"strings"
	"sync"
	"time"
	"unicode/utf16"

	"github.com/sourcegraph/jsonrpc2"
	"github.com/tliron/glsp"
	protocol "github.com/tliron/glsp/protocol_3_16"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vbadb"
	"github.com/harumiWeb/xlflow/internal/vbafmt"
)

const serverName = "xlflow-vba-lsp"

const diagnosticsDebounce = 300 * time.Millisecond

type BuildInfo struct {
	Version string
	Commit  string
	Date    string
}

type Options struct {
	RootDir        string
	Config         config.Config
	Build          BuildInfo
	LogFile        string
	Stderr         io.Writer
	TypeDBDir      string
	PerformanceLog bool
}

type Server struct {
	opts                     Options
	db                       *vbadb.DB
	analyzer                 intel.Analyzer
	handler                  protocol.Handler
	docs                     *documents
	logger                   *log.Logger
	symbols                  *workspaceSymbolCache
	semanticTokens           *semanticTokenCache
	semanticTokenGenerator   func(intel.Document, []intel.Document) ([]intel.SemanticToken, error)
	codeLensConfig           intel.CodeLensConfig
	diagnostics              func(context.Context, intel.Document) []intel.Diagnostic
	diagnosticsDebounce      time.Duration
	diagnosticsAfterFunc     func(time.Duration, func()) diagnosticTimer
	beforeDiagnosticsPublish func()

	diagMu      sync.Mutex
	diagStates  map[string]*diagnosticState
	diagWorkers sync.WaitGroup
	diagStopped bool

	docLifecycleMu sync.Mutex
	docLifecycles  map[string]*sync.Mutex
}

type diagnosticTimer interface {
	Stop() bool
}

type diagnosticState struct {
	mu         sync.Mutex
	generation uint64
	latest     intel.Document
	notify     *glsp.Context
	timer      diagnosticTimer
	running    bool
	ready      bool
	open       bool
	cancel     context.CancelFunc
}

func Check(opts Options) error {
	result, err := typedb.LoadForRuntime(opts.TypeDBDir)
	if err != nil {
		return err
	}
	w := opts.Stderr
	if w == nil {
		w = os.Stderr
	}
	for _, warning := range result.Warnings {
		_, _ = fmt.Fprintf(w, "type database warning: %s\n", warning)
	}
	return intel.Analyzer{RootDir: opts.RootDir, Config: opts.Config, DB: result.DB}.Check()
}

func RunStdio(opts Options) error {
	s, cleanup, err := New(opts)
	if err != nil {
		return err
	}
	defer cleanup()
	stream := jsonrpc2.NewBufferedStream(stdioReadWriteCloser{}, jsonrpc2.VSCodeObjectCodec{})
	conn := jsonrpc2.NewConn(context.Background(), stream, rpcHandler{handler: &s.handler})
	<-conn.DisconnectNotify()
	return conn.Close()
}

func New(opts Options) (*Server, func(), error) {
	typeDB, err := typedb.LoadForRuntime(opts.TypeDBDir)
	if err != nil {
		return nil, nil, err
	}
	logger, cleanup, err := newLogger(opts)
	if err != nil {
		return nil, nil, err
	}
	for _, warning := range typeDB.Warnings {
		logger.Printf("type database warning: %s", warning)
	}
	s := &Server{
		opts: opts,
		db:   typeDB.DB,
		analyzer: intel.Analyzer{
			RootDir: opts.RootDir,
			Config:  opts.Config,
			DB:      typeDB.DB,
		},
		docs:           newDocuments(opts.RootDir),
		logger:         logger,
		symbols:        newWorkspaceSymbolCache(),
		semanticTokens: newSemanticTokenCache(),
		codeLensConfig: intel.DefaultCodeLensConfig(),
		diagStates:     make(map[string]*diagnosticState),
		docLifecycles:  make(map[string]*sync.Mutex),
	}
	s.analyzer.DocumentSymbolsFunc = s.cachedDocumentSourceSymbols
	s.analyzer.WorkspaceSymbolsFunc = s.cachedWorkspaceSymbols
	s.semanticTokenGenerator = s.analyzer.SemanticTokens
	s.diagnostics = s.analyzer.DiagnosticsContext
	s.diagnosticsDebounce = diagnosticsDebounce
	s.diagnosticsAfterFunc = func(delay time.Duration, callback func()) diagnosticTimer {
		return time.AfterFunc(delay, callback)
	}
	s.handler = protocol.Handler{
		Initialize:                     s.initialize,
		Initialized:                    s.initialized,
		Shutdown:                       s.shutdown,
		Exit:                           s.exit,
		TextDocumentDidOpen:            s.didOpen,
		TextDocumentDidChange:          s.didChange,
		TextDocumentDidClose:           s.didClose,
		TextDocumentDocumentSymbol:     s.documentSymbol,
		WorkspaceSymbol:                s.workspaceSymbol,
		TextDocumentDefinition:         s.definition,
		TextDocumentReferences:         s.references,
		TextDocumentPrepareRename:      s.prepareRename,
		TextDocumentRename:             s.rename,
		TextDocumentHover:              s.hover,
		TextDocumentCompletion:         s.completion,
		TextDocumentCodeAction:         s.codeAction,
		TextDocumentSignatureHelp:      s.signatureHelp,
		TextDocumentFormatting:         s.formatting,
		TextDocumentSemanticTokensFull: s.semanticTokensFull,
		TextDocumentCodeLens:           s.codeLens,
	}
	return s, func() {
		s.stopDiagnostics()
		s.docs.closeAll()
		cleanup()
	}, nil
}

func newLogger(opts Options) (*log.Logger, func(), error) {
	if strings.TrimSpace(opts.LogFile) != "" {
		path := opts.LogFile
		if !filepath.IsAbs(path) {
			path = filepath.Join(opts.RootDir, path)
		}
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			return nil, nil, err
		}
		file, err := os.OpenFile(path, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0o644)
		if err != nil {
			return nil, nil, err
		}
		return log.New(file, "xlflow-lsp: ", log.LstdFlags), func() { _ = file.Close() }, nil
	}
	w := opts.Stderr
	if w == nil {
		w = os.Stderr
	}
	return log.New(w, "xlflow-lsp: ", log.LstdFlags), func() {}, nil
}

func (s *Server) initialize(_ *glsp.Context, params *protocol.InitializeParams) (any, error) {
	s.codeLensConfig = codeLensConfigFromInitialize(params)
	capabilities := s.handler.CreateServerCapabilities()
	if capabilities.CodeLensProvider != nil {
		resolveProvider := false
		capabilities.CodeLensProvider.ResolveProvider = &resolveProvider
	}
	if syncOptions, ok := capabilities.TextDocumentSync.(*protocol.TextDocumentSyncOptions); ok {
		kind := protocol.TextDocumentSyncKindFull
		syncOptions.Change = &kind
	}
	if capabilities.CompletionProvider != nil {
		capabilities.CompletionProvider.TriggerCharacters = completionTriggerCharacters()
	}
	if capabilities.SignatureHelpProvider != nil {
		capabilities.SignatureHelpProvider.TriggerCharacters = []string{"(", ",", " "}
		capabilities.SignatureHelpProvider.RetriggerCharacters = []string{","}
	}
	if capabilities.RenameProvider != nil {
		prepareProvider := true
		capabilities.RenameProvider = protocol.RenameOptions{PrepareProvider: &prepareProvider}
	}
	if semantic, ok := capabilities.SemanticTokensProvider.(*protocol.SemanticTokensOptions); ok {
		semantic.Legend = protocol.SemanticTokensLegend{
			TokenTypes:     intel.SemanticTokenTypes,
			TokenModifiers: intel.SemanticTokenModifiers,
		}
		semantic.Full = true
		semantic.Range = nil
	}
	version := s.opts.Build.Version
	if version == "" {
		version = "dev"
	}
	return protocol.InitializeResult{
		Capabilities: capabilities,
		ServerInfo: &protocol.InitializeResultServerInfo{
			Name:    serverName,
			Version: &version,
		},
	}, nil
}

func (s *Server) initialized(_ *glsp.Context, _ *protocol.InitializedParams) error {
	s.logger.Printf("initialized")
	return nil
}

func codeLensConfigFromInitialize(params *protocol.InitializeParams) intel.CodeLensConfig {
	cfg := intel.DefaultCodeLensConfig()
	if params == nil {
		return cfg
	}
	options, ok := params.InitializationOptions.(map[string]any)
	if !ok {
		return cfg
	}
	codeLens, ok := options["codeLens"].(map[string]any)
	if !ok {
		return cfg
	}
	if value, ok := codeLens["enabled"].(bool); ok {
		cfg.Enabled = value
	}
	if value, ok := codeLens["runProcedure"].(bool); ok {
		cfg.RunProcedure = value
	}
	if value, ok := codeLens["runTests"].(bool); ok {
		cfg.RunTests = value
	}
	if value, ok := codeLens["userFormEvents"].(bool); ok {
		cfg.UserFormEvents = value
	}
	return cfg
}

func (s *Server) shutdown(_ *glsp.Context) error {
	s.stopDiagnostics()
	s.docs.closeAll()
	s.logger.Printf("shutdown")
	return nil
}

func (s *Server) exit(_ *glsp.Context) error {
	s.logger.Printf("exit")
	return nil
}

func (s *Server) didOpen(ctx *glsp.Context, params *protocol.DidOpenTextDocumentParams) error {
	uri := string(params.TextDocument.URI)
	unlock := s.lockDocumentLifecycle(uri)
	doc, err := s.docs.open(uri, params.TextDocument.Text, int32(params.TextDocument.Version))
	if err != nil {
		unlock()
		return err
	}
	s.semanticTokens.invalidateAll()
	done := s.openDiagnostics(ctx, doc)
	unlock()
	if done != nil {
		<-done
	}
	return nil
}

func (s *Server) didChange(ctx *glsp.Context, params *protocol.DidChangeTextDocumentParams) error {
	if len(params.ContentChanges) == 0 {
		return nil
	}
	text, ok := fullChangeText(params.ContentChanges[len(params.ContentChanges)-1])
	if !ok {
		return fmt.Errorf("textDocument/didChange expected full document synchronization")
	}
	uri := string(params.TextDocument.URI)
	unlock := s.lockDocumentLifecycle(uri)
	defer unlock()
	doc, err := s.docs.change(uri, text, int32(params.TextDocument.Version))
	if err != nil {
		return err
	}
	s.semanticTokens.invalidateAll()
	s.scheduleDiagnostics(ctx, doc)
	return nil
}

func fullChangeText(change any) (string, bool) {
	switch typed := change.(type) {
	case protocol.TextDocumentContentChangeEventWhole:
		return typed.Text, true
	case protocol.TextDocumentContentChangeEvent:
		if typed.Range == nil {
			return typed.Text, true
		}
		return "", false
	default:
		return "", false
	}
}

func (s *Server) didClose(ctx *glsp.Context, params *protocol.DidCloseTextDocumentParams) error {
	uri := string(params.TextDocument.URI)
	unlock := s.lockDocumentLifecycle(uri)
	defer unlock()
	s.closeDiagnostics(ctx, uri)
	s.docs.close(uri)
	s.semanticTokens.invalidateAll()
	s.symbols.invalidateBase()
	return nil
}

func (s *Server) documentSymbol(_ *glsp.Context, params *protocol.DocumentSymbolParams) (any, error) {
	measurement := s.startPerformanceURI("textDocument/documentSymbol", string(params.TextDocument.URI))
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	measurement.setDocument(doc)
	syms, err := s.analyzer.DocumentSymbols(doc)
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	out := make([]protocol.DocumentSymbol, 0, len(syms))
	for _, sym := range syms {
		detail := sym.Detail
		out = append(out, protocol.DocumentSymbol{
			Name:           sym.Name,
			Detail:         &detail,
			Kind:           symbolKind(sym.Kind),
			Range:          toProtocolRange(sym.Range),
			SelectionRange: toProtocolRange(sym.Selection),
		})
	}
	measurement.finish(len(out), nil)
	return out, nil
}

func (s *Server) workspaceSymbol(_ *glsp.Context, params *protocol.WorkspaceSymbolParams) ([]protocol.SymbolInformation, error) {
	measurement := s.startPerformance("workspace/symbol", intel.Document{})
	syms, err := s.analyzer.WorkspaceSymbols(s.docs.openDocuments(), params.Query)
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	out := make([]protocol.SymbolInformation, 0, len(syms))
	for _, sym := range syms {
		uri := sym.File
		if !strings.HasPrefix(uri, "file:") {
			uri = s.docs.uriForDisplayPath(sym.File)
		}
		out = append(out, protocol.SymbolInformation{
			Name: sym.Name,
			Kind: symbolKind(sym.Kind),
			Location: protocol.Location{
				URI:   protocol.DocumentUri(uri),
				Range: toProtocolRange(sym.Selection),
			},
			ContainerName: &sym.Module,
		})
	}
	measurement.finish(len(out), nil)
	return out, nil
}

func (s *Server) definition(_ *glsp.Context, params *protocol.DefinitionParams) (any, error) {
	measurement := s.startPerformanceURI("textDocument/definition", string(params.TextDocument.URI))
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	measurement.setDocument(doc)
	locs, err := s.analyzer.Definition(doc, fromProtocolPosition(params.Position), s.docs.openDocuments(), s.docs.uriForDisplayPath)
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	out := make([]protocol.Location, 0, len(locs))
	for _, loc := range locs {
		uri := loc.URI
		if !strings.HasPrefix(uri, "file:") {
			uri = s.docs.uriForDisplayPath(loc.Path)
		}
		out = append(out, protocol.Location{URI: protocol.DocumentUri(uri), Range: toProtocolRange(loc.Range)})
	}
	measurement.finish(len(out), nil)
	return out, nil
}

func (s *Server) references(_ *glsp.Context, params *protocol.ReferenceParams) ([]protocol.Location, error) {
	measurement := s.startPerformanceURI("textDocument/references", string(params.TextDocument.URI))
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	measurement.setDocument(doc)
	locs, err := s.analyzer.References(doc, fromProtocolPosition(params.Position), s.docs.openDocuments(), params.Context.IncludeDeclaration, s.docs.uriForDisplayPath)
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	out := make([]protocol.Location, 0, len(locs))
	for _, loc := range locs {
		uri := loc.URI
		if !strings.HasPrefix(uri, "file:") {
			uri = s.docs.uriForDisplayPath(loc.Path)
		}
		out = append(out, protocol.Location{URI: protocol.DocumentUri(uri), Range: toProtocolRange(loc.Range)})
	}
	measurement.finish(len(out), nil)
	return out, nil
}

func (s *Server) prepareRename(_ *glsp.Context, params *protocol.PrepareRenameParams) (any, error) {
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		return nil, err
	}
	target, err := s.analyzer.PrepareRename(doc, fromProtocolPosition(params.Position), s.docs.openDocuments())
	if err != nil {
		return nil, err
	}
	return protocol.RangeWithPlaceholder{
		Range:       toProtocolRange(target.Range),
		Placeholder: target.Name,
	}, nil
}

func (s *Server) rename(_ *glsp.Context, params *protocol.RenameParams) (*protocol.WorkspaceEdit, error) {
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		return nil, err
	}
	edits, err := s.analyzer.Rename(doc, fromProtocolPosition(params.Position), params.NewName, s.docs.openDocuments(), s.docs.uriForDisplayPath)
	if err != nil {
		return nil, err
	}
	changes := map[protocol.DocumentUri][]protocol.TextEdit{}
	for _, edit := range edits {
		uri := edit.URI
		if uri == "" {
			uri = s.docs.uriForDisplayPath(edit.Path)
		}
		docURI := protocol.DocumentUri(uri)
		changes[docURI] = append(changes[docURI], protocol.TextEdit{
			Range:   toProtocolRange(edit.Range),
			NewText: edit.NewText,
		})
	}
	return &protocol.WorkspaceEdit{Changes: changes}, nil
}

func (s *Server) hover(_ *glsp.Context, params *protocol.HoverParams) (*protocol.Hover, error) {
	measurement := s.startPerformanceURI("textDocument/hover", string(params.TextDocument.URI))
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	measurement.setDocument(doc)
	hover, err := s.analyzer.Hover(doc, fromProtocolPosition(params.Position), s.docs.openDocuments())
	if err != nil || hover == nil {
		measurement.finish(0, err)
		return nil, err
	}
	measurement.finish(1, nil)
	return &protocol.Hover{
		Contents: protocol.MarkupContent{
			Kind:  protocol.MarkupKindMarkdown,
			Value: hover.Contents,
		},
		Range: toProtocolRangePtr(hover.Range),
	}, nil
}

func (s *Server) completion(_ *glsp.Context, params *protocol.CompletionParams) (any, error) {
	measurement := s.startPerformanceURI("textDocument/completion", string(params.TextDocument.URI))
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	measurement.setDocument(doc)
	completions, err := s.analyzer.Completions(doc, fromProtocolPosition(params.Position), s.docs.openDocuments())
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	items := make([]protocol.CompletionItem, 0, len(completions))
	for _, completion := range completions {
		kind := completionItemKind(completion.Kind)
		item := protocol.CompletionItem{
			Label: completion.Label,
			Kind:  &kind,
		}
		if completion.InsertText != "" {
			item.InsertText = &completion.InsertText
		}
		if completion.SortText != "" {
			item.SortText = &completion.SortText
		}
		if completion.ReplaceRange != nil {
			item.TextEdit = protocol.TextEdit{
				Range:   toProtocolRange(*completion.ReplaceRange),
				NewText: firstNonEmpty(completion.InsertText, completion.Label),
			}
		}
		if completion.Snippet {
			format := protocol.InsertTextFormatSnippet
			item.InsertTextFormat = &format
		}
		if completion.Detail != "" {
			item.Detail = &completion.Detail
		}
		if completion.Documentation != "" {
			item.Documentation = protocol.MarkupContent{
				Kind:  protocol.MarkupKindMarkdown,
				Value: completion.Documentation,
			}
		}
		items = append(items, item)
	}
	measurement.finish(len(items), nil)
	return protocol.CompletionList{IsIncomplete: false, Items: items}, nil
}

func (s *Server) codeAction(_ *glsp.Context, params *protocol.CodeActionParams) (any, error) {
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		return nil, err
	}
	actions, err := s.analyzer.DocumentationCodeActions(doc, fromProtocolRange(params.Range))
	if err != nil {
		return nil, err
	}
	kind := protocol.CodeActionKindRefactorRewrite
	out := make([]protocol.CodeAction, 0, len(actions))
	requestURI := protocol.DocumentUri(params.TextDocument.URI)
	for _, action := range actions {
		out = append(out, protocol.CodeAction{
			Title: action.Title,
			Kind:  &kind,
			Edit: &protocol.WorkspaceEdit{Changes: map[protocol.DocumentUri][]protocol.TextEdit{
				requestURI: {{
					Range:   toProtocolRange(action.Range),
					NewText: action.NewText,
				}},
			}},
		})
	}
	return out, nil
}

func (s *Server) signatureHelp(_ *glsp.Context, params *protocol.SignatureHelpParams) (*protocol.SignatureHelp, error) {
	measurement := s.startPerformanceURI("textDocument/signatureHelp", string(params.TextDocument.URI))
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	measurement.setDocument(doc)
	help, err := s.analyzer.SignatureHelp(doc, fromProtocolPosition(params.Position), s.docs.openDocuments())
	if err != nil || help == nil || len(help.Signatures) == 0 {
		measurement.finish(0, err)
		return nil, err
	}
	activeSignature := protocol.UInteger(max(0, help.ActiveSignature))
	activeParameter := protocol.UInteger(max(0, help.ActiveParameter))
	signatures := make([]protocol.SignatureInformation, 0, len(help.Signatures))
	for _, sig := range help.Signatures {
		info := protocol.SignatureInformation{Label: sig.Label}
		if sig.Documentation != "" {
			info.Documentation = protocol.MarkupContent{Kind: protocol.MarkupKindMarkdown, Value: sig.Documentation}
		}
		for _, param := range sig.Parameters {
			paramInfo := protocol.ParameterInformation{Label: parameterLabel(param)}
			if param.Documentation != "" {
				paramInfo.Documentation = protocol.MarkupContent{Kind: protocol.MarkupKindMarkdown, Value: param.Documentation}
			}
			info.Parameters = append(info.Parameters, paramInfo)
		}
		signatures = append(signatures, info)
	}
	measurement.finish(len(signatures), nil)
	return &protocol.SignatureHelp{
		Signatures:      signatures,
		ActiveSignature: &activeSignature,
		ActiveParameter: &activeParameter,
	}, nil
}

func (s *Server) formatting(_ *glsp.Context, params *protocol.DocumentFormattingParams) ([]protocol.TextEdit, error) {
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		return nil, err
	}
	if !documentSupportsFormatting(doc) {
		return []protocol.TextEdit{}, nil
	}
	formatted, err := vbafmt.FormatTextWithOptions(doc.Source, documentIsClass(doc), vbafmt.FormatConfig{
		LineNumbers:           vbafmt.LineNumberModePreserve,
		OperatorSpacing:       s.opts.Config.Fmt.OperatorSpacing,
		OperatorSpacingSet:    true,
		DeclarationSpacing:    s.opts.Config.Fmt.DeclarationSpacing,
		DeclarationSpacingSet: true,
		KeywordCasing:         s.opts.Config.Fmt.KeywordCasing,
		KeywordCasingSet:      true,
		BuiltinCasing:         s.opts.Config.Fmt.BuiltinCasing,
		BuiltinCasingSet:      true,
	})
	if err != nil {
		if vbafmt.IsFormatParseError(err) {
			return []protocol.TextEdit{}, nil
		}
		return nil, err
	}
	if formatted == doc.Source {
		return []protocol.TextEdit{}, nil
	}
	return []protocol.TextEdit{{
		Range:   fullDocumentRange(doc.Source),
		NewText: formatted,
	}}, nil
}

func (s *Server) semanticTokensFull(_ *glsp.Context, params *protocol.SemanticTokensParams) (*protocol.SemanticTokens, error) {
	measurement := s.startPerformanceURI("textDocument/semanticTokens/full", string(params.TextDocument.URI))
	for {
		generation := s.semanticTokens.begin()
		doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
		if err != nil {
			measurement.finish(0, err)
			return nil, err
		}
		measurement.setDocument(doc)
		cacheStarted := time.Now()
		data, hit, err := s.semanticTokens.get(doc, generation, func() ([]protocol.UInteger, error) {
			tokens, err := s.semanticTokenGenerator(doc, s.docs.openDocuments())
			if err != nil {
				return nil, err
			}
			return encodeSemanticTokens(tokens), nil
		})
		if errors.Is(err, errSemanticTokensSuperseded) {
			continue
		}
		s.logDocumentCachePerformance("semanticTokens/cache", cacheStatus(hit), doc, len(data)/5, cacheStarted, err)
		if err != nil {
			measurement.finish(0, err)
			return nil, err
		}
		result := &protocol.SemanticTokens{Data: data}
		measurement.finish(len(data)/5, nil)
		return result, nil
	}
}

func (s *Server) codeLens(_ *glsp.Context, params *protocol.CodeLensParams) ([]protocol.CodeLens, error) {
	measurement := s.startPerformanceURI("textDocument/codeLens", string(params.TextDocument.URI))
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	measurement.setDocument(doc)
	procedures, err := s.analyzer.RunnableProcedures(doc, s.codeLensConfig)
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	out := make([]protocol.CodeLens, 0, len(procedures))
	for _, procedure := range procedures {
		title := "$(play) Run"
		command := "xlflow.runProcedure"
		if procedure.Kind == "test" {
			title = "$(beaker) Run Test"
			command = "xlflow.runTestProcedure"
		}
		pos := protocol.Position{Line: protocol.UInteger(max(0, procedure.Line)), Character: protocol.UInteger(max(0, procedure.Character))}
		out = append(out, protocol.CodeLens{
			Range: protocol.Range{Start: pos, End: pos},
			Command: &protocol.Command{
				Title:   title,
				Command: command,
				Arguments: []any{map[string]any{
					"uri":           procedure.URI,
					"name":          procedure.Name,
					"moduleName":    procedure.ModuleName,
					"qualifiedName": procedure.QualifiedName,
					"kind":          procedure.Kind,
					"moduleKind":    procedure.ModuleKind,
					"line":          procedure.Line,
					"character":     procedure.Character,
				}},
			},
		})
	}
	measurement.finish(len(out), nil)
	return out, nil
}

func (s *Server) openDiagnostics(ctx *glsp.Context, doc intel.Document) <-chan struct{} {
	s.diagMu.Lock()
	if s.diagStopped {
		s.diagMu.Unlock()
		return nil
	}
	state := s.diagStates[doc.URI]
	if state == nil {
		state = &diagnosticState{}
		s.diagStates[doc.URI] = state
	}
	state.mu.Lock()
	state.generation++
	state.latest = doc
	state.notify = ctx
	state.ready = true
	state.open = true
	if state.timer != nil {
		state.timer.Stop()
		state.timer = nil
	}
	if state.cancel != nil {
		state.cancel()
	}
	state.mu.Unlock()
	s.diagMu.Unlock()
	return s.launchDiagnostics(doc.URI, state)
}

func (s *Server) scheduleDiagnostics(ctx *glsp.Context, doc intel.Document) {
	s.diagMu.Lock()
	if s.diagStopped {
		s.diagMu.Unlock()
		return
	}
	state := s.diagStates[doc.URI]
	if state == nil {
		state = &diagnosticState{open: true}
		s.diagStates[doc.URI] = state
	}
	state.mu.Lock()
	state.generation++
	generation := state.generation
	state.latest = doc
	state.notify = ctx
	state.ready = false
	state.open = true
	if state.timer != nil {
		state.timer.Stop()
	}
	if state.cancel != nil {
		state.cancel()
	}
	state.timer = s.diagnosticsAfterFunc(s.diagnosticsDebounce, func() {
		s.diagnosticsReady(doc.URI, state, generation)
	})
	state.mu.Unlock()
	s.diagMu.Unlock()
}

func (s *Server) diagnosticsReady(uri string, state *diagnosticState, generation uint64) {
	state.mu.Lock()
	if !state.open || state.generation != generation {
		state.mu.Unlock()
		return
	}
	state.timer = nil
	state.ready = true
	state.mu.Unlock()
	s.launchDiagnostics(uri, state)
}

func (s *Server) launchDiagnostics(uri string, state *diagnosticState) <-chan struct{} {
	s.diagMu.Lock()
	if s.diagStopped || s.diagStates[uri] != state {
		s.diagMu.Unlock()
		return nil
	}
	state.mu.Lock()
	if !state.open || !state.ready || state.running {
		state.mu.Unlock()
		s.diagMu.Unlock()
		return nil
	}
	doc := state.latest
	notify := state.notify
	generation := state.generation
	runCtx, cancel := context.WithCancel(context.Background())
	state.latest = intel.Document{}
	state.ready = false
	state.running = true
	state.cancel = cancel
	s.diagWorkers.Add(1)
	done := make(chan struct{})
	state.mu.Unlock()
	s.diagMu.Unlock()

	go func() {
		defer close(done)
		s.runDiagnostics(runCtx, uri, state, generation, doc, notify)
	}()
	return done
}

func (s *Server) runDiagnostics(
	runCtx context.Context,
	uri string,
	state *diagnosticState,
	generation uint64,
	doc intel.Document,
	notify *glsp.Context,
) {
	defer s.diagWorkers.Done()
	measurement := s.startPerformance("diagnostics", doc)
	diagnostics := s.diagnostics(runCtx, doc)
	out := make([]protocol.Diagnostic, 0, len(diagnostics))
	for _, diag := range diagnostics {
		severity := diagnosticSeverity(diag.Severity)
		source := diag.Source
		code := protocol.IntegerOrString{Value: diag.Code}
		out = append(out, protocol.Diagnostic{
			Range:    toProtocolRange(diag.Range),
			Severity: &severity,
			Code:     &code,
			Source:   &source,
			Message:  diag.Message,
		})
	}
	if s.beforeDiagnosticsPublish != nil {
		s.beforeDiagnosticsPublish()
	}

	state.mu.Lock()
	discarded := !state.open || state.generation != generation || runCtx.Err() != nil
	if !discarded && notify != nil {
		notify.Notify(string(protocol.ServerTextDocumentPublishDiagnostics), protocol.PublishDiagnosticsParams{
			URI:         protocol.DocumentUri(doc.URI),
			Diagnostics: out,
		})
	}
	state.running = false
	state.cancel = nil
	ready := state.open && state.ready
	state.mu.Unlock()
	measurement.finishDiagnostics(len(out), generation, discarded)

	if ready {
		s.launchDiagnostics(uri, state)
	}
}

func (s *Server) closeDiagnostics(ctx *glsp.Context, uri string) {
	s.diagMu.Lock()
	state := s.diagStates[uri]
	s.diagMu.Unlock()
	if state != nil {
		state.mu.Lock()
		state.close()
		if ctx != nil {
			ctx.Notify(string(protocol.ServerTextDocumentPublishDiagnostics), protocol.PublishDiagnosticsParams{
				URI:         protocol.DocumentUri(uri),
				Diagnostics: []protocol.Diagnostic{},
			})
		}
		state.mu.Unlock()
	} else if ctx != nil {
		ctx.Notify(string(protocol.ServerTextDocumentPublishDiagnostics), protocol.PublishDiagnosticsParams{
			URI:         protocol.DocumentUri(uri),
			Diagnostics: []protocol.Diagnostic{},
		})
	}
}

func (state *diagnosticState) close() {
	state.generation++
	state.open = false
	state.ready = false
	state.latest = intel.Document{}
	state.notify = nil
	if state.timer != nil {
		state.timer.Stop()
		state.timer = nil
	}
	if state.cancel != nil {
		state.cancel()
	}
}

func (s *Server) stopDiagnostics() {
	s.diagMu.Lock()
	var states []*diagnosticState
	if !s.diagStopped {
		s.diagStopped = true
		for uri, state := range s.diagStates {
			states = append(states, state)
			delete(s.diagStates, uri)
		}
	}
	s.diagMu.Unlock()
	for _, state := range states {
		state.mu.Lock()
		state.close()
		state.mu.Unlock()
	}
	s.diagWorkers.Wait()
}

func (s *Server) lockDocumentLifecycle(uri string) func() {
	s.docLifecycleMu.Lock()
	lifecycle := s.docLifecycles[uri]
	if lifecycle == nil {
		lifecycle = &sync.Mutex{}
		s.docLifecycles[uri] = lifecycle
	}
	s.docLifecycleMu.Unlock()
	lifecycle.Lock()
	return lifecycle.Unlock
}

func (s *Server) baseAnalyzer() intel.Analyzer {
	return intel.Analyzer{RootDir: s.opts.RootDir, Config: s.opts.Config, DB: s.db}
}

func (s *Server) cachedDocumentSourceSymbols(doc intel.Document, load intel.DocumentSymbolLoader) ([]intel.Symbol, error) {
	started := time.Now()
	var syms []intel.Symbol
	var hit bool
	var err error
	if doc.Snapshot != nil && doc.Snapshot.Matches(doc) {
		syms, hit, err = doc.Snapshot.SourceSymbols(load)
	} else {
		syms, err = load()
	}
	s.logDocumentCachePerformance("documentSymbols/cache", cacheStatus(hit), doc, len(syms), started, err)
	return syms, err
}

func (s *Server) cachedWorkspaceSymbols(open []intel.Document, query string) ([]intel.Symbol, error) {
	analyzer := s.baseAnalyzer()
	started := time.Now()
	base, hit, err := s.symbols.base(analyzer)
	s.logCachePerformance("workspaceSymbols/cache/base", cacheStatus(hit), len(base), started, err)
	if err != nil {
		return nil, err
	}
	openKeys := make(map[string]bool, len(open))
	for _, doc := range open {
		for _, key := range s.symbolPathKeys(doc.Path) {
			openKeys[key] = true
		}
	}
	out := make([]intel.Symbol, 0, len(base))
	for _, sym := range base {
		if hasSymbolPathKey(openKeys, s.symbolPathKeys(sym.File)) {
			continue
		}
		out = append(out, sym)
	}
	for _, doc := range open {
		syms, err := s.analyzer.DocumentSymbols(doc)
		if err != nil {
			continue
		}
		out = append(out, syms...)
	}
	return filterWorkspaceSymbols(out, query), nil
}

type workspaceSymbolCache struct {
	mu          sync.RWMutex
	baseSymbols []intel.Symbol
	baseOK      bool
}

func newWorkspaceSymbolCache() *workspaceSymbolCache {
	return &workspaceSymbolCache{}
}

func (c *workspaceSymbolCache) base(analyzer intel.Analyzer) ([]intel.Symbol, bool, error) {
	c.mu.RLock()
	if c.baseOK {
		out := cloneSymbols(c.baseSymbols)
		c.mu.RUnlock()
		return out, true, nil
	}
	c.mu.RUnlock()

	syms, err := analyzer.WorkspaceSymbols(nil, "")
	if err != nil {
		return nil, false, err
	}
	c.mu.Lock()
	c.baseSymbols = cloneSymbols(syms)
	c.baseOK = true
	c.mu.Unlock()
	return cloneSymbols(syms), false, nil
}

func (c *workspaceSymbolCache) invalidateBase() {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.baseSymbols = nil
	c.baseOK = false
}

func documentSymbolKey(doc intel.Document) string {
	if doc.Path != "" {
		return symbolFileKey(doc.Path)
	}
	if doc.URI != "" {
		if path, err := fileURIToPath(doc.URI); err == nil {
			return symbolFileKey(path)
		}
		return strings.ToLower(doc.URI)
	}
	return ""
}

func symbolFileKey(path string) string {
	if path == "" {
		return ""
	}
	if strings.HasPrefix(path, "file:") {
		if decoded, err := fileURIToPath(path); err == nil {
			path = decoded
		}
	}
	return normalizePathKey(path)
}

func (s *Server) symbolPathKeys(path string) []string {
	var keys []string
	if key := symbolFileKey(path); key != "" {
		keys = append(keys, key)
	}
	if strings.TrimSpace(path) == "" || strings.TrimSpace(s.opts.RootDir) == "" {
		return keys
	}
	if filepath.IsAbs(path) {
		if rel, err := filepath.Rel(s.opts.RootDir, path); err == nil {
			if key := symbolFileKey(rel); key != "" {
				keys = append(keys, key)
			}
		}
		return keys
	}
	if key := symbolFileKey(filepath.Join(s.opts.RootDir, filepath.FromSlash(path))); key != "" {
		keys = append(keys, key)
	}
	return keys
}

func hasSymbolPathKey(set map[string]bool, keys []string) bool {
	for _, key := range keys {
		if set[key] {
			return true
		}
	}
	return false
}

func cloneSymbols(syms []intel.Symbol) []intel.Symbol {
	out := make([]intel.Symbol, len(syms))
	for i, sym := range syms {
		out[i] = sym
		out[i].Parameters = append([]intel.Parameter(nil), sym.Parameters...)
		out[i].Documentation.ParameterEntries = append(out[i].Documentation.ParameterEntries[:0:0], sym.Documentation.ParameterEntries...)
		if sym.Documentation.Parameters != nil {
			out[i].Documentation.Parameters = make(map[string]string, len(sym.Documentation.Parameters))
			for key, value := range sym.Documentation.Parameters {
				out[i].Documentation.Parameters[key] = value
			}
		}
		if sym.Documentation.UnknownSections != nil {
			out[i].Documentation.UnknownSections = make(map[string]string, len(sym.Documentation.UnknownSections))
			for key, value := range sym.Documentation.UnknownSections {
				out[i].Documentation.UnknownSections[key] = value
			}
		}
	}
	return out
}

func filterWorkspaceSymbols(syms []intel.Symbol, query string) []intel.Symbol {
	query = strings.ToLower(strings.TrimSpace(query))
	if query != "" {
		filtered := syms[:0]
		for _, sym := range syms {
			if strings.Contains(strings.ToLower(sym.Name), query) || strings.Contains(strings.ToLower(sym.Module+"."+sym.Name), query) {
				filtered = append(filtered, sym)
			}
		}
		syms = filtered
	}
	sort.SliceStable(syms, func(i, j int) bool {
		if syms[i].File != syms[j].File {
			return syms[i].File < syms[j].File
		}
		if syms[i].Range.Start.Line != syms[j].Range.Start.Line {
			return syms[i].Range.Start.Line < syms[j].Range.Start.Line
		}
		if syms[i].Range.Start.Character != syms[j].Range.Start.Character {
			return syms[i].Range.Start.Character < syms[j].Range.Start.Character
		}
		return syms[i].Name < syms[j].Name
	})
	return syms
}

type documents struct {
	root          string
	readFile      func(string) ([]byte, error)
	beforePublish func()
	mu            sync.RWMutex
	docs          map[string]documentEntry
	keys          map[string]string
	closed        bool
}

type documentEntry struct {
	snapshot *intel.AnalysisSnapshot
	open     bool
}

func newDocuments(root string) *documents {
	return &documents{root: root, readFile: os.ReadFile, docs: map[string]documentEntry{}, keys: map[string]string{}}
}

func (d *documents) open(uri, text string, versions ...int32) (intel.Document, error) {
	doc, err := d.docFromURI(uri, text)
	if err != nil {
		return intel.Document{}, err
	}
	if len(versions) > 0 {
		doc.Version = versions[0]
	}
	snapshot := intel.NewAnalysisSnapshot(doc)
	if d.beforePublish != nil {
		d.beforePublish()
	}
	d.mu.Lock()
	if d.closed {
		d.mu.Unlock()
		snapshot.Retire()
		return intel.Document{}, errDocumentsClosed
	}
	key := normalizePathKey(doc.Path)
	previous := d.docs[key].snapshot
	d.docs[key] = documentEntry{snapshot: snapshot, open: true}
	d.keys[uri] = key
	d.mu.Unlock()
	previous.Retire()
	return snapshot.Document(), nil
}

func (d *documents) change(uri, text string, versions ...int32) (intel.Document, error) {
	d.mu.RLock()
	if d.closed {
		d.mu.RUnlock()
		return intel.Document{}, errDocumentsClosed
	}
	key := d.keys[uri]
	if key == "" {
		if path, err := fileURIToPath(uri); err == nil {
			key = normalizePathKey(path)
		}
	}
	entry, ok := d.docs[key]
	d.mu.RUnlock()
	if !ok || !entry.open || entry.snapshot == nil {
		return d.open(uri, text, versions...)
	}
	current := entry.snapshot.Document()
	current.Source = text
	current.Snapshot = nil
	if len(versions) > 0 {
		current.Version = versions[0]
	}
	snapshot := intel.NewAnalysisSnapshot(current)
	if d.beforePublish != nil {
		d.beforePublish()
	}
	d.mu.Lock()
	if d.closed {
		d.mu.Unlock()
		snapshot.Retire()
		return intel.Document{}, errDocumentsClosed
	}
	latest := d.docs[key]
	if latest.snapshot != entry.snapshot || !latest.open {
		if latest.open && latest.snapshot != nil {
			d.keys[uri] = key
			if snapshot.Version() > latest.snapshot.Version() {
				d.docs[key] = documentEntry{snapshot: snapshot, open: true}
				d.mu.Unlock()
				latest.snapshot.Retire()
				return snapshot.Document(), nil
			}
		}
		d.mu.Unlock()
		snapshot.Retire()
		if latest.open && latest.snapshot != nil {
			return latest.snapshot.Document(), nil
		}
		return intel.Document{}, errDocumentChangedConcurrently
	}
	d.docs[key] = documentEntry{snapshot: snapshot, open: true}
	d.keys[uri] = key
	d.mu.Unlock()
	entry.snapshot.Retire()
	return snapshot.Document(), nil
}

func (d *documents) close(uri string) {
	d.mu.Lock()
	var snapshot *intel.AnalysisSnapshot
	if key := d.keys[uri]; key != "" {
		if entry := d.docs[key]; entry.open {
			snapshot = entry.snapshot
			delete(d.docs, key)
		}
		delete(d.keys, uri)
	}
	d.mu.Unlock()
	snapshot.Retire()
}

func (d *documents) getOrRead(uri string) (intel.Document, error) {
	path, err := fileURIToPath(uri)
	if err != nil {
		return intel.Document{}, err
	}
	key := normalizePathKey(path)
	d.mu.RLock()
	if d.closed {
		d.mu.RUnlock()
		return intel.Document{}, errDocumentsClosed
	}
	if entry, ok := d.docs[key]; ok && entry.open && entry.snapshot != nil {
		d.mu.RUnlock()
		return entry.snapshot.Document(), nil
	}
	d.mu.RUnlock()
	body, err := d.readFile(path)
	if err != nil {
		return intel.Document{}, err
	}
	candidate := intel.NewAnalysisSnapshot(intel.Document{URI: uri, Path: path, Source: string(body), ModuleKind: moduleKindForPath(path)})
	if d.beforePublish != nil {
		d.beforePublish()
	}
	d.mu.Lock()
	if d.closed {
		d.mu.Unlock()
		candidate.Retire()
		return intel.Document{}, errDocumentsClosed
	}
	current := d.docs[key]
	if current.open && current.snapshot != nil {
		d.mu.Unlock()
		candidate.Retire()
		return current.snapshot.Document(), nil
	}
	if current.snapshot != nil && current.snapshot.SourceHash() == candidate.SourceHash() &&
		current.snapshot.URI() == candidate.URI() && current.snapshot.ModuleKind() == candidate.ModuleKind() {
		d.mu.Unlock()
		candidate.Retire()
		return current.snapshot.Document(), nil
	}
	d.docs[key] = documentEntry{snapshot: candidate}
	d.mu.Unlock()
	current.snapshot.Retire()
	return candidate.Document(), nil
}

func (d *documents) openDocuments() []intel.Document {
	d.mu.RLock()
	defer d.mu.RUnlock()
	out := make([]intel.Document, 0, len(d.docs))
	for _, entry := range d.docs {
		if entry.open && entry.snapshot != nil {
			out = append(out, entry.snapshot.Document())
		}
	}
	return out
}

func (d *documents) closeAll() {
	d.mu.Lock()
	snapshots := make([]*intel.AnalysisSnapshot, 0, len(d.docs))
	for _, entry := range d.docs {
		if entry.snapshot != nil {
			snapshots = append(snapshots, entry.snapshot)
		}
	}
	d.docs = make(map[string]documentEntry)
	d.keys = make(map[string]string)
	d.closed = true
	d.mu.Unlock()
	for _, snapshot := range snapshots {
		snapshot.Retire()
	}
}

var (
	errDocumentsClosed             = errors.New("LSP document snapshot store is closed")
	errDocumentChangedConcurrently = errors.New("LSP document changed concurrently")
)

func (d *documents) docFromURI(uri, text string) (intel.Document, error) {
	path, err := fileURIToPath(uri)
	if err != nil {
		return intel.Document{}, err
	}
	return intel.Document{URI: uri, Path: path, Source: text, ModuleKind: moduleKindForPath(path)}, nil
}

func (d *documents) uriForDisplayPath(path string) string {
	if strings.HasPrefix(path, "file:") {
		return path
	}
	if !filepath.IsAbs(path) {
		path = filepath.Join(d.root, filepath.FromSlash(path))
	}
	return pathToFileURI(path)
}

func fileURIToPath(raw string) (string, error) {
	u, err := url.Parse(raw)
	if err != nil {
		return "", err
	}
	if u.Scheme != "file" {
		return "", fmt.Errorf("unsupported URI scheme %q", u.Scheme)
	}
	path, err := url.PathUnescape(u.EscapedPath())
	if err != nil {
		return "", err
	}
	if runtime.GOOS == "windows" {
		if u.Host != "" {
			path = `\\` + u.Host + filepath.FromSlash(path)
		} else {
			path = strings.TrimPrefix(path, "/")
			path = filepath.FromSlash(path)
		}
	} else {
		path = filepath.FromSlash(path)
	}
	return filepath.Clean(path), nil
}

func pathToFileURI(path string) string {
	path = filepath.Clean(path)
	host := ""
	if runtime.GOOS == "windows" {
		vol := filepath.VolumeName(path)
		if strings.HasPrefix(vol, `\\`) {
			rest := strings.TrimPrefix(path, vol)
			hostShare := strings.TrimPrefix(vol, `\\`)
			parts := strings.SplitN(hostShare, `\`, 2)
			if len(parts) == 2 {
				host = parts[0]
				path = "/" + parts[1] + filepath.ToSlash(rest)
			} else {
				path = "/" + filepath.ToSlash(path)
			}
		} else {
			path = "/" + filepath.ToSlash(path)
		}
	} else {
		path = filepath.ToSlash(path)
	}
	return (&url.URL{Scheme: "file", Host: host, Path: path}).String()
}

func normalizePathKey(path string) string {
	clean := filepath.Clean(path)
	if runtime.GOOS == "windows" {
		clean = strings.ToLower(clean)
	}
	return clean
}

func moduleKindForPath(path string) string {
	switch strings.ToLower(filepath.Ext(path)) {
	case ".cls":
		return "class"
	case ".frm":
		return "form"
	default:
		return "standard"
	}
}

func documentSupportsFormatting(doc intel.Document) bool {
	ext := strings.ToLower(filepath.Ext(doc.Path))
	return ext == ".bas" || ext == ".cls"
}

func documentIsClass(doc intel.Document) bool {
	return strings.EqualFold(doc.ModuleKind, "class") || strings.EqualFold(filepath.Ext(doc.Path), ".cls")
}

func fullDocumentRange(source string) protocol.Range {
	source = strings.ReplaceAll(source, "\r\n", "\n")
	source = strings.ReplaceAll(source, "\r", "\n")
	lines := strings.Split(source, "\n")
	lastLine := len(lines) - 1
	lastChar := 0
	if lastLine >= 0 {
		lastChar = utf16Len(lines[lastLine])
	}
	return protocol.Range{
		Start: protocol.Position{Line: 0, Character: 0},
		End:   protocol.Position{Line: protocol.UInteger(max(0, lastLine)), Character: protocol.UInteger(max(0, lastChar))},
	}
}

func toProtocolRange(r intel.Range) protocol.Range {
	return protocol.Range{Start: toProtocolPosition(r.Start), End: toProtocolPosition(r.End)}
}

func toProtocolRangePtr(r intel.Range) *protocol.Range {
	out := toProtocolRange(r)
	return &out
}

func toProtocolPosition(pos intel.Position) protocol.Position {
	return protocol.Position{Line: protocol.UInteger(max(0, pos.Line)), Character: protocol.UInteger(max(0, pos.Character))}
}

func fromProtocolPosition(pos protocol.Position) intel.Position {
	return intel.Position{Line: int(pos.Line), Character: int(pos.Character)}
}

func fromProtocolRange(r protocol.Range) intel.Range {
	return intel.Range{Start: fromProtocolPosition(r.Start), End: fromProtocolPosition(r.End)}
}

func utf16Len(s string) int {
	return len(utf16.Encode([]rune(s)))
}

func diagnosticSeverity(severity string) protocol.DiagnosticSeverity {
	switch strings.ToLower(severity) {
	case "error":
		return protocol.DiagnosticSeverityError
	case "info":
		return protocol.DiagnosticSeverityInformation
	case "hint":
		return protocol.DiagnosticSeverityHint
	default:
		return protocol.DiagnosticSeverityWarning
	}
}

func encodeSemanticTokens(tokens []intel.SemanticToken) []protocol.UInteger {
	out := make([]protocol.UInteger, 0, len(tokens)*5)
	prevLine, prevStart := 0, 0
	for _, token := range tokens {
		line := max(0, token.Range.Start.Line)
		start := max(0, token.Range.Start.Character)
		length := max(0, token.Range.End.Character-token.Range.Start.Character)
		if length == 0 || token.Range.End.Line != token.Range.Start.Line {
			continue
		}
		deltaLine := line - prevLine
		deltaStart := start
		if deltaLine == 0 {
			deltaStart = start - prevStart
		}
		if deltaStart < 0 {
			continue
		}
		typeIndex := semanticTokenTypeIndex(token.Type)
		if typeIndex < 0 {
			continue
		}
		out = append(out,
			protocol.UInteger(deltaLine),
			protocol.UInteger(deltaStart),
			protocol.UInteger(length),
			protocol.UInteger(typeIndex),
			protocol.UInteger(semanticTokenModifierMask(token.Modifiers)),
		)
		prevLine = line
		prevStart = start
	}
	return out
}

func semanticTokenTypeIndex(tokenType string) int {
	for i, candidate := range intel.SemanticTokenTypes {
		if candidate == tokenType {
			return i
		}
	}
	return -1
}

func semanticTokenModifierMask(modifiers []string) int {
	mask := 0
	for _, modifier := range modifiers {
		for i, candidate := range intel.SemanticTokenModifiers {
			if modifier == candidate {
				mask |= 1 << i
			}
		}
	}
	return mask
}

func symbolKind(kind string) protocol.SymbolKind {
	switch strings.ToLower(kind) {
	case "module":
		return protocol.SymbolKindModule
	case "class":
		return protocol.SymbolKindClass
	case "sub", "function", "property", "property_get", "property_let", "property_set":
		return protocol.SymbolKindFunction
	case "const":
		return protocol.SymbolKindConstant
	case "field", "module_variable":
		return protocol.SymbolKindField
	case "local_variable":
		return protocol.SymbolKindVariable
	case "enum":
		return protocol.SymbolKindEnum
	case "event":
		return protocol.SymbolKindEvent
	default:
		return protocol.SymbolKindObject
	}
}

func completionItemKind(kind string) protocol.CompletionItemKind {
	switch strings.ToLower(kind) {
	case "method":
		return protocol.CompletionItemKindMethod
	case "function":
		return protocol.CompletionItemKindFunction
	case "property":
		return protocol.CompletionItemKindProperty
	case "variable":
		return protocol.CompletionItemKindVariable
	case "type":
		return protocol.CompletionItemKindClass
	case "constant":
		return protocol.CompletionItemKindConstant
	case "keyword":
		return protocol.CompletionItemKindKeyword
	case "snippet":
		return protocol.CompletionItemKindSnippet
	default:
		return protocol.CompletionItemKindText
	}
}

func parameterLabel(param intel.Parameter) string {
	var b strings.Builder
	if param.Optional {
		b.WriteString("Optional ")
	}
	b.WriteString(param.Name)
	if param.Type != "" {
		b.WriteString(" As ")
		b.WriteString(param.Type)
	}
	return b.String()
}

func completionTriggerCharacters() []string {
	return []string{".", "\"", "'", "@"}
}

func max(a, b int) int {
	if a > b {
		return a
	}
	return b
}

func firstNonEmpty(values ...string) string {
	for _, value := range values {
		if strings.TrimSpace(value) != "" {
			return value
		}
	}
	return ""
}

type stdioReadWriteCloser struct{}

func (stdioReadWriteCloser) Read(p []byte) (int, error)  { return os.Stdin.Read(p) }
func (stdioReadWriteCloser) Write(p []byte) (int, error) { return os.Stdout.Write(p) }
func (stdioReadWriteCloser) Close() error                { return nil }

type rpcHandler struct {
	handler glsp.Handler
}

func (h rpcHandler) Handle(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) {
	params := []byte("{}")
	if req.Params != nil {
		params = *req.Params
	}
	glspCtx := &glsp.Context{
		Method: req.Method,
		Params: params,
		Notify: func(method string, params any) {
			_ = conn.Notify(ctx, method, params)
		},
		Call: func(method string, params any, result any) {
			_ = conn.Call(ctx, method, params, result)
		},
	}
	result, validMethod, validParams, err := h.handler.Handle(glspCtx)
	if !validMethod {
		if !req.Notif {
			_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: jsonrpc2.CodeMethodNotFound, Message: "method not found"})
		}
		return
	}
	if !validParams {
		if !req.Notif {
			_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: jsonrpc2.CodeInvalidParams, Message: "invalid params"})
		}
		return
	}
	if err != nil {
		if !req.Notif {
			_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: jsonrpc2.CodeInternalError, Message: err.Error()})
		}
		return
	}
	if !req.Notif {
		_ = conn.Reply(ctx, req.ID, result)
	}
	if req.Method == "exit" {
		_ = conn.Close()
	}
}
