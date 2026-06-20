package lspserver

import (
	"context"
	"fmt"
	"io"
	"log"
	"net/url"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"time"

	"github.com/sourcegraph/jsonrpc2"
	"github.com/tliron/glsp"
	protocol "github.com/tliron/glsp/protocol_3_16"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

const serverName = "xlflow-vba-lsp"

const diagnosticsDebounce = 300 * time.Millisecond

type BuildInfo struct {
	Version string
	Commit  string
	Date    string
}

type Options struct {
	RootDir string
	Config  config.Config
	Build   BuildInfo
	LogFile string
	Stderr  io.Writer
}

type Server struct {
	opts     Options
	db       *vbadb.DB
	analyzer intel.Analyzer
	handler  protocol.Handler
	docs     *documents
	logger   *log.Logger

	diagMu     sync.Mutex
	diagTimers map[string]*time.Timer
}

func Check(opts Options) error {
	db, err := vbadb.LoadBuiltin()
	if err != nil {
		return err
	}
	return intel.Analyzer{RootDir: opts.RootDir, Config: opts.Config, DB: db}.Check()
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
	db, err := vbadb.LoadBuiltin()
	if err != nil {
		return nil, nil, err
	}
	logger, cleanup, err := newLogger(opts)
	if err != nil {
		return nil, nil, err
	}
	s := &Server{
		opts: opts,
		db:   db,
		analyzer: intel.Analyzer{
			RootDir: opts.RootDir,
			Config:  opts.Config,
			DB:      db,
		},
		docs:       newDocuments(opts.RootDir),
		logger:     logger,
		diagTimers: make(map[string]*time.Timer),
	}
	s.handler = protocol.Handler{
		Initialize:                 s.initialize,
		Initialized:                s.initialized,
		Shutdown:                   s.shutdown,
		Exit:                       s.exit,
		TextDocumentDidOpen:        s.didOpen,
		TextDocumentDidChange:      s.didChange,
		TextDocumentDidClose:       s.didClose,
		TextDocumentDocumentSymbol: s.documentSymbol,
		WorkspaceSymbol:            s.workspaceSymbol,
		TextDocumentDefinition:     s.definition,
		TextDocumentReferences:     s.references,
		TextDocumentHover:          s.hover,
		TextDocumentCompletion:     s.completion,
	}
	return s, func() {
		s.stopDiagnosticTimers()
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

func (s *Server) initialize(_ *glsp.Context, _ *protocol.InitializeParams) (any, error) {
	capabilities := s.handler.CreateServerCapabilities()
	if syncOptions, ok := capabilities.TextDocumentSync.(*protocol.TextDocumentSyncOptions); ok {
		kind := protocol.TextDocumentSyncKindFull
		syncOptions.Change = &kind
	}
	if capabilities.CompletionProvider != nil {
		capabilities.CompletionProvider.TriggerCharacters = completionTriggerCharacters()
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

func (s *Server) shutdown(_ *glsp.Context) error {
	s.logger.Printf("shutdown")
	return nil
}

func (s *Server) exit(_ *glsp.Context) error {
	s.logger.Printf("exit")
	return nil
}

func (s *Server) didOpen(ctx *glsp.Context, params *protocol.DidOpenTextDocumentParams) error {
	doc, err := s.docs.open(string(params.TextDocument.URI), params.TextDocument.Text)
	if err != nil {
		return err
	}
	s.publishDiagnostics(ctx, doc)
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
	doc, err := s.docs.change(string(params.TextDocument.URI), text)
	if err != nil {
		return err
	}
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
	s.cancelDiagnostics(uri)
	s.docs.close(uri)
	ctx.Notify(string(protocol.ServerTextDocumentPublishDiagnostics), protocol.PublishDiagnosticsParams{
		URI:         protocol.DocumentUri(uri),
		Diagnostics: []protocol.Diagnostic{},
	})
	return nil
}

func (s *Server) documentSymbol(_ *glsp.Context, params *protocol.DocumentSymbolParams) (any, error) {
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		return nil, err
	}
	syms, err := s.analyzer.DocumentSymbols(doc)
	if err != nil {
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
	return out, nil
}

func (s *Server) workspaceSymbol(_ *glsp.Context, params *protocol.WorkspaceSymbolParams) ([]protocol.SymbolInformation, error) {
	syms, err := s.analyzer.WorkspaceSymbols(s.docs.openDocuments(), params.Query)
	if err != nil {
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
	return out, nil
}

func (s *Server) definition(_ *glsp.Context, params *protocol.DefinitionParams) (any, error) {
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		return nil, err
	}
	locs, err := s.analyzer.Definition(doc, fromProtocolPosition(params.Position), s.docs.openDocuments(), s.docs.uriForDisplayPath)
	if err != nil {
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
	return out, nil
}

func (s *Server) references(_ *glsp.Context, params *protocol.ReferenceParams) ([]protocol.Location, error) {
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		return nil, err
	}
	locs, err := s.analyzer.References(doc, fromProtocolPosition(params.Position), s.docs.openDocuments(), params.Context.IncludeDeclaration, s.docs.uriForDisplayPath)
	if err != nil {
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
	return out, nil
}

func (s *Server) hover(_ *glsp.Context, params *protocol.HoverParams) (*protocol.Hover, error) {
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		return nil, err
	}
	hover, err := s.analyzer.Hover(doc, fromProtocolPosition(params.Position), s.docs.openDocuments())
	if err != nil || hover == nil {
		return nil, err
	}
	return &protocol.Hover{
		Contents: protocol.MarkupContent{
			Kind:  protocol.MarkupKindMarkdown,
			Value: hover.Contents,
		},
		Range: toProtocolRangePtr(hover.Range),
	}, nil
}

func (s *Server) completion(_ *glsp.Context, params *protocol.CompletionParams) (any, error) {
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		return nil, err
	}
	completions, err := s.analyzer.Completions(doc, fromProtocolPosition(params.Position), s.docs.openDocuments())
	if err != nil {
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
	return protocol.CompletionList{IsIncomplete: false, Items: items}, nil
}

func (s *Server) scheduleDiagnostics(ctx *glsp.Context, doc intel.Document) {
	if doc.URI == "" {
		s.publishDiagnostics(ctx, doc)
		return
	}
	s.diagMu.Lock()
	if timer := s.diagTimers[doc.URI]; timer != nil {
		timer.Stop()
	}
	s.diagTimers[doc.URI] = time.AfterFunc(diagnosticsDebounce, func() {
		s.diagMu.Lock()
		delete(s.diagTimers, doc.URI)
		s.diagMu.Unlock()
		s.publishDiagnostics(ctx, doc)
	})
	s.diagMu.Unlock()
}

func (s *Server) cancelDiagnostics(uri string) {
	s.diagMu.Lock()
	defer s.diagMu.Unlock()
	if timer := s.diagTimers[uri]; timer != nil {
		timer.Stop()
		delete(s.diagTimers, uri)
	}
}

func (s *Server) stopDiagnosticTimers() {
	s.diagMu.Lock()
	defer s.diagMu.Unlock()
	for uri, timer := range s.diagTimers {
		timer.Stop()
		delete(s.diagTimers, uri)
	}
}

func (s *Server) publishDiagnostics(ctx *glsp.Context, doc intel.Document) {
	diagnostics := s.analyzer.Diagnostics(doc)
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
	ctx.Notify(string(protocol.ServerTextDocumentPublishDiagnostics), protocol.PublishDiagnosticsParams{
		URI:         protocol.DocumentUri(doc.URI),
		Diagnostics: out,
	})
}

type documents struct {
	root string
	mu   sync.RWMutex
	docs map[string]intel.Document
	keys map[string]string
}

func newDocuments(root string) *documents {
	return &documents{root: root, docs: map[string]intel.Document{}, keys: map[string]string{}}
}

func (d *documents) open(uri, text string) (intel.Document, error) {
	doc, err := d.docFromURI(uri, text)
	if err != nil {
		return intel.Document{}, err
	}
	d.mu.Lock()
	defer d.mu.Unlock()
	key := normalizePathKey(doc.Path)
	d.docs[key] = doc
	d.keys[uri] = key
	return doc, nil
}

func (d *documents) change(uri, text string) (intel.Document, error) {
	d.mu.RLock()
	key := d.keys[uri]
	current, ok := d.docs[key]
	d.mu.RUnlock()
	if !ok {
		return d.open(uri, text)
	}
	current.Source = text
	d.mu.Lock()
	d.docs[key] = current
	d.mu.Unlock()
	return current, nil
}

func (d *documents) close(uri string) {
	d.mu.Lock()
	defer d.mu.Unlock()
	if key := d.keys[uri]; key != "" {
		delete(d.docs, key)
		delete(d.keys, uri)
	}
}

func (d *documents) getOrRead(uri string) (intel.Document, error) {
	path, err := fileURIToPath(uri)
	if err != nil {
		return intel.Document{}, err
	}
	key := normalizePathKey(path)
	d.mu.RLock()
	if doc, ok := d.docs[key]; ok {
		d.mu.RUnlock()
		return doc, nil
	}
	d.mu.RUnlock()
	body, err := os.ReadFile(path)
	if err != nil {
		return intel.Document{}, err
	}
	return intel.Document{URI: uri, Path: path, Source: string(body), ModuleKind: moduleKindForPath(path)}, nil
}

func (d *documents) openDocuments() []intel.Document {
	d.mu.RLock()
	defer d.mu.RUnlock()
	out := make([]intel.Document, 0, len(d.docs))
	for _, doc := range d.docs {
		out = append(out, doc)
	}
	return out
}

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

func completionTriggerCharacters() []string {
	return []string{"."}
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
