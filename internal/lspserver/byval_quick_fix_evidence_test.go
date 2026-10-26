package lspserver

import (
	"context"
	"encoding/json"
	"fmt"
	"net"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/sourcegraph/jsonrpc2"
	"github.com/tliron/glsp"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

func TestCurrentQuickFixDiagnosticsRequiresCurrentFullEvidence(t *testing.T) {
	doc := intel.Document{
		URI:        "file:///workspace/Main.bas",
		Path:       filepath.Join("workspace", "Main.bas"),
		Source:     "Sub Work()\nEnd Sub\n",
		ModuleKind: "standard",
		Version:    7,
	}
	tests := []struct {
		name      string
		prepare   func(*diagnosticState, *intel.Document)
		wantCount int
	}{
		{
			name:      "current full evidence",
			wantCount: 1,
		},
		{
			name: "missing evidence",
			prepare: func(state *diagnosticState, _ *intel.Document) {
				state.quickFixEvidence = nil
			},
		},
		{
			name: "fast only publication",
			prepare: func(state *diagnosticState, _ *intel.Document) {
				state.publishedMode = intel.DiagnosticModeFast
			},
		},
		{
			name: "stale generation",
			prepare: func(state *diagnosticState, _ *intel.Document) {
				state.generation++
			},
		},
		{
			name: "stale dependency generation",
			prepare: func(state *diagnosticState, _ *intel.Document) {
				state.dependencyGeneration++
			},
		},
		{
			name: "dependent diagnostics pending",
			prepare: func(state *diagnosticState, _ *intel.Document) {
				state.dependentPending = true
			},
		},
		{
			name: "requested document mismatch",
			prepare: func(_ *diagnosticState, requested *intel.Document) {
				requested.Source += "' changed\n"
			},
		},
		{
			name: "latest document differs from accepted evidence",
			prepare: func(state *diagnosticState, requested *intel.Document) {
				requested.Source += "' changed\n"
				state.latest = *requested
			},
		},
		{
			name: "closed",
			prepare: func(state *diagnosticState, _ *intel.Document) {
				state.close()
			},
		},
		{
			name: "reopened after close",
			prepare: func(state *diagnosticState, requested *intel.Document) {
				state.close()
				state.open = true
				state.latest = *requested
			},
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			requested := doc
			state := &diagnosticState{
				generation:           3,
				latest:               doc,
				open:                 true,
				publishedMode:        intel.DiagnosticModeFull,
				hasPublished:         true,
				dependencyGeneration: 11,
				quickFixEvidence: newQuickFixDiagnosticEvidence(doc, 3, 11, true, []intel.Diagnostic{
					{Code: byValQuickFixDiagnosticCode, Message: "parameter can be ByVal"},
				}),
			}
			server := &Server{diagStates: map[string]*diagnosticState{doc.URI: state}}
			if test.prepare != nil {
				test.prepare(state, &requested)
			}

			got := server.currentQuickFixDiagnostics(requested)
			if len(got) != test.wantCount {
				t.Fatalf("currentQuickFixDiagnostics() returned %d diagnostics, want %d: %+v", len(got), test.wantCount, got)
			}
		})
	}
}

func TestCurrentQuickFixDiagnosticsReturnsIndependentVBA275Copies(t *testing.T) {
	doc := intel.Document{URI: "file:///workspace/Main.bas", Path: "Main.bas", Source: "Sub Work()\nEnd Sub\n", Version: 1}
	input := []intel.Diagnostic{
		{Code: "VBA274", Message: "other parameter diagnostic"},
		{Code: byValQuickFixDiagnosticCode, Message: "parameter can be ByVal", Range: intel.Range{Start: intel.Position{Line: 2}}},
	}
	evidence := newQuickFixDiagnosticEvidence(doc, 4, 9, true, input)
	input[1].Message = "mutated source slice"
	if evidence == nil || len(evidence.diagnostics) != 1 || evidence.diagnostics[0].Code != byValQuickFixDiagnosticCode {
		t.Fatalf("stored evidence = %+v, want only VBA275", evidence)
	}

	state := &diagnosticState{
		generation:           4,
		latest:               doc,
		open:                 true,
		publishedMode:        intel.DiagnosticModeFull,
		hasPublished:         true,
		dependencyGeneration: 9,
		quickFixEvidence:     evidence,
	}
	server := &Server{diagStates: map[string]*diagnosticState{doc.URI: state}}
	first := server.currentQuickFixDiagnostics(doc)
	first[0].Message = "mutated returned copy"
	first[0].Range.Start.Line = 99
	second := server.currentQuickFixDiagnostics(doc)
	if len(second) != 1 || second[0].Message != "parameter can be ByVal" || second[0].Range.Start.Line != 2 {
		t.Fatalf("subsequent evidence = %+v, want an unchanged independent copy", second)
	}
}

func TestCodeActionWithholdsByValFixWhenDependencyChangesDuringDiscovery(t *testing.T) {
	const source = "Private Function Echo(ByRef value As Long) As Long\n Echo = value\nEnd Function\n"
	server, doc, cleanup := newByValProjectTestServer(t, source)
	defer cleanup()
	byRefStart := strings.Index(source, "ByRef")
	diagnostic := intel.Diagnostic{
		Code:     byValQuickFixDiagnosticCode,
		Severity: "information",
		Source:   "xlflow",
		Message:  "ByRef parameter value is never written and can be passed ByVal.",
		Range: intel.Range{
			Start: intel.Position{Character: byRefStart},
			End:   intel.Position{Character: byRefStart + len("ByRef")},
		},
	}
	state := &diagnosticState{
		generation:           1,
		latest:               doc,
		open:                 true,
		publishedMode:        intel.DiagnosticModeFull,
		hasPublished:         true,
		dependencyGeneration: 8,
		quickFixEvidence:     newQuickFixDiagnosticEvidence(doc, 1, 8, true, []intel.Diagnostic{diagnostic}),
	}
	server.diagMu.Lock()
	server.diagStates[doc.URI] = state
	server.diagMu.Unlock()

	currentEvidence := server.analyzer.QuickFixDiagnosticsFunc
	callbackCalls := 0
	server.analyzer.QuickFixDiagnosticsFunc = func(request intel.Document) []intel.Diagnostic {
		got := currentEvidence(request)
		if len(got) != 0 {
			callbackCalls++
			state.mu.Lock()
			state.dependencyGeneration++
			state.mu.Unlock()
		}
		return got
	}
	params := &protocol.CodeActionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(doc.URI)},
		Range:        toProtocolRange(diagnostic.Range),
		Context:      protocol.CodeActionContext{Only: []protocol.CodeActionKind{protocol.CodeActionKindQuickFix}},
	}
	result, err := server.codeActionForDocument(t.Context(), doc, params)
	if err != nil {
		t.Fatal(err)
	}
	if callbackCalls == 0 {
		t.Fatal("ByVal discovery did not read the initially current VBA275 evidence")
	}
	for _, action := range result.([]protocol.CodeAction) {
		if action.Title == "Pass parameter ByVal" {
			t.Fatalf("stale ByVal action survived response-boundary revalidation: %+v", action)
		}
	}
}

func TestByValActionRevalidationPreservesUnrelatedActions(t *testing.T) {
	actions := []intel.CodeAction{
		{Title: "Pass parameter ByVal", Kind: "quickfix", Diagnostics: []intel.Diagnostic{{Code: byValQuickFixDiagnosticCode}}},
		{Title: "Add Option Explicit", Kind: "quickfix", Diagnostics: []intel.Diagnostic{{Code: "VB001"}}},
	}
	projected := []protocol.CodeAction{{Title: "Pass parameter ByVal"}, {Title: "Add Option Explicit"}}
	got := retainCurrentByValQuickFixActions(actions, projected, nil)
	if len(got) != 1 || got[0].Title != "Add Option Explicit" {
		t.Fatalf("revalidated actions = %+v, want unrelated VB001 action retained", got)
	}
}

func TestCodeActionProjectionRevalidatesByValEvidence(t *testing.T) {
	for _, versioned := range []bool{false, true} {
		for _, stale := range []bool{false, true} {
			t.Run(fmt.Sprintf("versioned=%t/stale=%t", versioned, stale), func(t *testing.T) {
				const source = "Private Function Echo(ByRef value As Long) As Long\n Echo = value\nEnd Function\n"
				server, doc, cleanup := newByValProjectTestServer(t, source)
				defer cleanup()
				server.codeActionDocumentChanges.Store(versioned)
				start := strings.Index(source, "ByRef")
				diagnostic := intel.Diagnostic{
					Code: byValQuickFixDiagnosticCode, Severity: "information", Source: "xlflow", Message: "parameter can be ByVal",
					Range: intel.Range{Start: intel.Position{Character: start}, End: intel.Position{Character: start + len("ByRef")}},
				}
				state := &diagnosticState{
					generation: 1, latest: doc, open: true, publishedMode: intel.DiagnosticModeFull,
					hasPublished: true, dependencyGeneration: 8,
					quickFixEvidence: newQuickFixDiagnosticEvidence(doc, 1, 8, true, []intel.Diagnostic{diagnostic}),
				}
				server.diagMu.Lock()
				server.diagStates[doc.URI] = state
				server.diagMu.Unlock()
				candidates, err := server.analyzer.CodeActionsContext(t.Context(), doc, diagnostic.Range, []string{"quickfix"})
				if err != nil || len(candidates) != 1 {
					t.Fatalf("candidates=%+v err=%v", candidates, err)
				}
				// Model the old response gap: discovery and the former pre-projection
				// evidence check succeeded, then a dependency changed without an edit.
				if len(server.currentQuickFixDiagnostics(doc)) != 1 {
					t.Fatal("expected current evidence before projection")
				}
				if stale {
					state.mu.Lock()
					state.dependencyGeneration++
					state.mu.Unlock()
				}
				candidates = append(candidates,
					intel.CodeAction{Title: "Unrelated Fix All", Kind: "source.fixAll.xlflow.vb084"},
					intel.CodeAction{Title: "Add Option Explicit", Kind: "quickfix", Diagnostics: []intel.Diagnostic{{Code: "VB001"}}},
				)
				got := server.projectCodeActions(doc, protocol.DocumentUri(doc.URI), candidates)
				wantCount := 2
				if versioned {
					wantCount++
				}
				if stale {
					wantCount--
				}
				if len(got) != wantCount || got[len(got)-1].Title != "Add Option Explicit" {
					t.Fatalf("projected actions=%+v, want %d with unrelated action retained", got, wantCount)
				}
				for _, action := range got {
					if action.Title == "Pass parameter ByVal" && stale {
						t.Fatalf("stale ByVal survived projection: %+v", action)
					}
					if versioned {
						edit := action.Edit.DocumentChanges[0].(protocol.TextDocumentEdit)
						if edit.TextDocument.Version == nil || *edit.TextDocument.Version != protocol.Integer(doc.Version) {
							t.Fatalf("projection lost unchanged document version: %+v", edit)
						}
					}
				}
			})
		}
	}
}

func TestQuickFixEvidenceTracksAcceptedFullPublications(t *testing.T) {
	tests := []struct {
		name         string
		mode         intel.DiagnosticMode
		cancel       bool
		obsolete     bool
		wantEvidence bool
		wantPublish  int
	}{
		{name: "non-project-aware full fallback", mode: intel.DiagnosticModeFull, wantPublish: 1},
		{name: "fast only", mode: intel.DiagnosticModeFast, wantPublish: 1},
		{name: "canceled full", mode: intel.DiagnosticModeFull, cancel: true},
		{name: "obsolete full", mode: intel.DiagnosticModeFull, obsolete: true},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			server, _, cleanup := newDiagnosticsTestServer(t)
			defer cleanup()

			path := filepath.Join(server.opts.RootDir, "Main.bas")
			doc := intel.Document{
				URI:        pathToFileURI(path),
				Path:       path,
				Source:     "Sub Work()\nEnd Sub\n",
				ModuleKind: "standard",
				Version:    2,
			}
			state := &diagnosticState{
				generation:           5,
				latest:               doc,
				open:                 true,
				dependencyGeneration: 12,
			}
			server.diagMu.Lock()
			server.diagStates[doc.URI] = state
			server.diagMu.Unlock()
			server.diagnosticsRequest = func(context.Context, intel.DiagnosticRequest) intel.DiagnosticResult {
				return intel.DiagnosticResult{Diagnostics: []intel.Diagnostic{
					{Code: "VBA274", Message: "not evidence for this fix"},
					{Code: byValQuickFixDiagnosticCode, Message: "current ByVal evidence"},
				}}
			}

			var published []protocol.PublishDiagnosticsParams
			notify := &glsp.Context{Notify: func(_ string, params any) {
				if diagnostics, ok := params.(protocol.PublishDiagnosticsParams); ok {
					published = append(published, diagnostics)
				}
			}}
			if test.obsolete {
				server.beforeDiagnosticsPublish = func() {
					state.mu.Lock()
					state.generation++
					state.mu.Unlock()
				}
			}
			runCtx := t.Context()
			if test.cancel {
				var cancel context.CancelFunc
				runCtx, cancel = context.WithCancel(runCtx)
				cancel()
			}
			server.runDiagnosticsBody(runCtx, doc.URI, state, 5, doc, notify, test.mode, 0)

			got := server.currentQuickFixDiagnostics(doc)
			if (len(got) == 1) != test.wantEvidence {
				t.Fatalf("current evidence = %+v, want present=%t", got, test.wantEvidence)
			}
			if len(published) != test.wantPublish {
				t.Fatalf("publish count = %d, want %d", len(published), test.wantPublish)
			}
			if test.wantEvidence && (got[0].Code != byValQuickFixDiagnosticCode || got[0].Message != "current ByVal evidence") {
				t.Fatalf("evidence = %+v, want the published VBA275 only", got)
			}
		})
	}
}

func TestQuickFixEvidenceRequiresCompleteProjectFullPublication(t *testing.T) {
	const source = "Private Function Echo(ByRef value As Long) As Long\n Echo = value\nEnd Function\n"
	byRefStart := strings.Index(source, "ByRef")
	wantRange := intel.Range{
		Start: intel.Position{Line: 0, Character: byRefStart},
		End:   intel.Position{Line: 0, Character: byRefStart + len("ByRef")},
	}

	for _, test := range []struct {
		name       string
		incomplete bool
		wantProof  bool
	}{
		{name: "complete project Full result", wantProof: true},
		{name: "incomplete project Full fallback", incomplete: true},
	} {
		t.Run(test.name, func(t *testing.T) {
			server, doc, cleanup := newByValProjectTestServer(t, source)
			defer cleanup()
			if !server.projectDiagnosticsEnabled() {
				t.Fatal("project-aware diagnostics must be enabled for this test")
			}
			if test.incomplete {
				server.analysis.beginOverlay(doc, 1)
			}
			project := server.analysis.projectSnapshotClassContext(t.Context(), analysisWorkBackground.String())
			if project.Complete == test.incomplete {
				t.Fatalf("project completeness = %t, incomplete case=%t", project.Complete, test.incomplete)
			}

			state := &diagnosticState{generation: 1, latest: doc, open: true, dependencyGeneration: 1}
			server.diagMu.Lock()
			server.diagStates[doc.URI] = state
			server.diagMu.Unlock()
			var published []protocol.PublishDiagnosticsParams
			notify := &glsp.Context{Notify: func(_ string, params any) {
				if diagnostics, ok := params.(protocol.PublishDiagnosticsParams); ok {
					published = append(published, diagnostics)
				}
			}}

			server.runDiagnosticsBody(t.Context(), doc.URI, state, 1, doc, notify, intel.DiagnosticModeFull, 0)
			got := server.currentQuickFixDiagnostics(doc)
			if (len(got) == 1) != test.wantProof {
				t.Fatalf("current evidence = %+v, want present=%t", got, test.wantProof)
			}
			publishedByRef := false
			for _, publication := range published {
				for _, diagnostic := range publication.Diagnostics {
					if diagnostic.Code != nil && diagnostic.Code.Value == byValQuickFixDiagnosticCode {
						publishedByRef = true
					}
				}
			}
			if !publishedByRef {
				t.Fatalf("Full publication did not contain analyzer VBA275: %+v", published)
			}
			if test.wantProof && got[0].Range != wantRange {
				t.Fatalf("published VBA275 range = %+v, want ByRef header range %+v", got[0].Range, wantRange)
			}
		})
	}
}

func TestQuickFixEvidenceIsReleasedOnCloseAndNotRestoredOnReopen(t *testing.T) {
	server, _, cleanup := newDiagnosticsTestServer(t)
	defer cleanup()
	doc := intel.Document{
		URI:        "file:///workspace/Main.bas",
		Path:       filepath.Join("workspace", "Main.bas"),
		Source:     "Sub Work()\nEnd Sub\n",
		ModuleKind: "standard",
		Version:    1,
	}
	state := &diagnosticState{
		generation:           2,
		latest:               doc,
		open:                 true,
		publishedMode:        intel.DiagnosticModeFull,
		hasPublished:         true,
		dependencyGeneration: 3,
		quickFixEvidence: newQuickFixDiagnosticEvidence(doc, 2, 3, true, []intel.Diagnostic{
			{Code: byValQuickFixDiagnosticCode, Message: "parameter can be ByVal"},
		}),
	}
	server.diagMu.Lock()
	server.diagStates[doc.URI] = state
	server.diagMu.Unlock()

	server.closeDiagnostics(nil, doc.URI)
	if state.quickFixEvidence != nil {
		t.Fatal("close retained quick-fix evidence")
	}
	if got := server.currentQuickFixDiagnostics(doc); len(got) != 0 {
		t.Fatalf("closed document evidence = %+v, want none", got)
	}

	server.openDiagnostics(nil, doc)
	if got := server.currentQuickFixDiagnostics(doc); len(got) != 0 {
		t.Fatalf("reopened document evidence = %+v, want none until a new Full publication", got)
	}
}

func TestByValQuickFixEvidenceAnalyzerHookIsWired(t *testing.T) {
	server, _, cleanup := newDiagnosticsTestServer(t)
	defer cleanup()
	if server.analyzer.QuickFixDiagnosticsFunc == nil {
		t.Fatal("analyzer QuickFixDiagnosticsFunc is not wired to server evidence")
	}
}

func TestJSONRPCByValQuickFixUsesCurrentPublishedProjectDiagnostic(t *testing.T) {
	const positiveSource = "Private Function Echo(ByRef value As Long) As Long\n Echo = value\nEnd Function\n"
	addressOfSource := positiveSource + "\nPublic Sub RegisterCallback()\n    Dim VTable(0) As LongPtr\n    VTable(0) = FncPtr(AddressOf Echo)\nEnd Sub\n"
	tests := []struct {
		name         string
		enabled      bool
		source       string
		diskSource   string
		wantVBA275   bool
		wantByValFix bool
	}{
		{name: "published diagnostic produces versioned edit from unsaved buffer", enabled: true, source: positiveSource, diskSource: strings.Replace(positiveSource, "ByRef", "ByVal", 1), wantVBA275: true, wantByValFix: true},
		{name: "opt-in disabled", source: positiveSource},
		{name: "suppressed diagnostic", enabled: true, source: "' xlflow:disable-next-line VBA275\n" + positiveSource},
		{name: "AddressOf constrained function", enabled: true, source: addressOfSource},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			ctx, cancel := context.WithTimeout(t.Context(), 30*time.Second)
			defer cancel()

			cfg := config.Default()
			cfg.Analyze.DetectByRefParametersCanBeByVal = test.enabled
			root := t.TempDir()
			path := filepath.Join(root, cfg.Src.Modules, "Main.bas")
			diskSource := test.diskSource
			if diskSource == "" {
				diskSource = test.source
			}
			if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
				t.Fatal(err)
			}
			if err := os.WriteFile(path, []byte(diskSource), 0o600); err != nil {
				t.Fatal(err)
			}
			server, cleanup, err := New(Options{RootDir: root, Config: cfg})
			if err != nil {
				t.Fatal(err)
			}
			defer cleanup()

			recorder := &rpcRecorder{}
			type completedProjectFull struct {
				publishedCount       int
				generation           uint64
				dependencyGeneration uint64
			}
			projectFullRuns := make(chan completedProjectFull, 8)
			projectDiagnostics := server.projectDiagnosticsRequest
			server.projectDiagnosticsRequest = func(runCtx context.Context, request intel.DiagnosticRequest, project intel.ProjectAnalysisSnapshot) intel.DiagnosticResult {
				result := projectDiagnostics(runCtx, request, project)
				if project.Complete && request.Document.URI == pathToFileURI(path) {
					server.diagMu.Lock()
					state := server.diagStates[request.Document.URI]
					server.diagMu.Unlock()
					var generation, dependencyGeneration uint64
					if state != nil {
						state.mu.Lock()
						generation = state.generation
						dependencyGeneration = state.dependencyGeneration
						state.mu.Unlock()
					}
					projectFullRuns <- completedProjectFull{
						publishedCount:       len(recorder.publishDiagnostics()),
						generation:           generation,
						dependencyGeneration: dependencyGeneration,
					}
				}
				return result
			}

			serverSide, clientSide := net.Pipe()
			serverConn := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(serverSide, jsonrpc2.VSCodeObjectCodec{}), rpcHandler{
				handler: &server.handler, server: server, custom: server.handleCustomNotification, dispatch: server.dispatchCodeAction,
			})
			defer func() { _ = serverConn.Close() }()
			clientConn := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(clientSide, jsonrpc2.VSCodeObjectCodec{}), recorder)
			defer func() { _ = clientConn.Close() }()

			var initializeResult protocol.InitializeResult
			if err := clientConn.Call(ctx, "initialize", fixAllInitializeParams(t, true), &initializeResult); err != nil {
				t.Fatal(err)
			}
			if err := clientConn.Notify(ctx, "initialized", map[string]any{}); err != nil {
				t.Fatal(err)
			}
			uri := pathToFileURI(path)
			if err := clientConn.Notify(ctx, string(protocol.MethodTextDocumentDidOpen), protocol.DidOpenTextDocumentParams{
				TextDocument: protocol.TextDocumentItem{URI: protocol.DocumentUri(uri), LanguageID: "vba", Version: 41, Text: test.source},
			}); err != nil {
				t.Fatal(err)
			}

			var acceptedPublication protocol.PublishDiagnosticsParams
			acceptedPublicationIndex := -1
			var currentDoc intel.Document
			for {
				var attempt completedProjectFull
				select {
				case attempt = <-projectFullRuns:
				case <-ctx.Done():
					t.Fatalf("project-aware Full diagnostics did not run: %v; publications=%+v", ctx.Err(), recorder.publishDiagnostics())
				}
				if !recorder.waitForDiagnostics(ctx, func(publications []protocol.PublishDiagnosticsParams) bool {
					return len(publications) > attempt.publishedCount
				}) {
					t.Fatalf("project-aware Full publication was not received: %v; publications=%+v", ctx.Err(), recorder.publishDiagnostics())
				}
				publications := recorder.publishDiagnostics()
				if len(publications) <= attempt.publishedCount {
					continue
				}
				server.diagMu.Lock()
				state := server.diagStates[uri]
				server.diagMu.Unlock()
				if state == nil {
					continue
				}
				state.mu.Lock()
				accepted := state.open && state.generation == attempt.generation &&
					state.dependencyGeneration == attempt.dependencyGeneration && state.hasPublished &&
					state.publishedMode == intel.DiagnosticModeFull && !state.dependentPending
				state.mu.Unlock()
				if !accepted {
					continue
				}
				acceptedPublication = publications[attempt.publishedCount]
				acceptedPublicationIndex = attempt.publishedCount
				currentDoc, err = server.docs.getOrRead(uri)
				if err != nil {
					t.Fatal(err)
				}
				break
			}

			recorder.mu.Lock()
			rawPublication := append(json.RawMessage(nil), recorder.paramsSeen[string(protocol.ServerTextDocumentPublishDiagnostics)][acceptedPublicationIndex]...)
			recorder.mu.Unlock()
			var wirePublication struct {
				Diagnostics []struct {
					Code    string         `json:"code"`
					Message string         `json:"message"`
					Range   protocol.Range `json:"range"`
				} `json:"diagnostics"`
			}
			if err := json.Unmarshal(rawPublication, &wirePublication); err != nil {
				t.Fatalf("decode accepted publishDiagnostics wire payload: %v", err)
			}
			publishedByValIndex := -1
			for i := range wirePublication.Diagnostics {
				diagnostic := &wirePublication.Diagnostics[i]
				if diagnostic.Code == byValQuickFixDiagnosticCode {
					publishedByValIndex = i
					break
				}
			}
			if (publishedByValIndex >= 0) != test.wantVBA275 {
				t.Fatalf("accepted project Full diagnostics contain VBA275=%t, want %t: %+v", publishedByValIndex >= 0, test.wantVBA275, acceptedPublication.Diagnostics)
			}
			evidence := server.currentQuickFixDiagnostics(currentDoc)
			if (len(evidence) == 1) != test.wantVBA275 {
				t.Fatalf("current VBA275 evidence = %+v, want present=%t", evidence, test.wantVBA275)
			}
			if test.wantVBA275 {
				publishedByVal := wirePublication.Diagnostics[publishedByValIndex]
				wantRange := toProtocolRange(evidence[0].Range)
				if publishedByVal.Range != wantRange || evidence[0].Range.Start.Line != 0 || publishedByVal.Message != evidence[0].Message {
					t.Fatalf("published VBA275=%+v evidence=%+v, want the current ByRef header diagnostic", publishedByVal, evidence[0])
				}
			}

			byRefStart := strings.Index(test.source, "ByRef")
			line := strings.Count(test.source[:byRefStart], "\n")
			lineStart := strings.LastIndex(test.source[:byRefStart], "\n") + 1
			selection := protocol.Range{
				Start: protocol.Position{Line: protocol.UInteger(line), Character: protocol.UInteger(byRefStart - lineStart)},
				End:   protocol.Position{Line: protocol.UInteger(line), Character: protocol.UInteger(byRefStart - lineStart + len("ByRef"))},
			}
			var actionsJSON json.RawMessage
			if err := clientConn.Call(ctx, "textDocument/codeAction", protocol.CodeActionParams{
				TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
				Range:        selection,
				Context:      protocol.CodeActionContext{Only: []protocol.CodeActionKind{protocol.CodeActionKindQuickFix}},
			}, &actionsJSON); err != nil {
				t.Fatal(err)
			}
			var actions []protocol.CodeAction
			if err := json.Unmarshal(actionsJSON, &actions); err != nil {
				t.Fatalf("decode codeAction response: %v", err)
			}
			var wireActions []struct {
				Title       string `json:"title"`
				Diagnostics []struct {
					Code string `json:"code"`
				} `json:"diagnostics"`
			}
			if err := json.Unmarshal(actionsJSON, &wireActions); err != nil {
				t.Fatalf("decode codeAction wire diagnostics: %v", err)
			}
			var byValAction *protocol.CodeAction
			wireByValActionHasVBA275 := false
			for i := range actions {
				if actions[i].Title == "Pass parameter ByVal" {
					byValAction = &actions[i]
				}
			}
			for _, action := range wireActions {
				if action.Title == "Pass parameter ByVal" {
					for _, diagnostic := range action.Diagnostics {
						if diagnostic.Code == byValQuickFixDiagnosticCode {
							wireByValActionHasVBA275 = true
						}
					}
				}
			}
			if (byValAction != nil) != test.wantByValFix {
				t.Fatalf("ByVal quick fix present=%t, want %t; actions=%+v", byValAction != nil, test.wantByValFix, actions)
			}
			if !test.wantByValFix {
				return
			}
			if byValAction.Kind == nil || *byValAction.Kind != protocol.CodeActionKindQuickFix || byValAction.Edit == nil ||
				len(byValAction.Diagnostics) != 1 || byValAction.Diagnostics[0].Message != evidence[0].Message ||
				byValAction.Diagnostics[0].Range != toProtocolRange(evidence[0].Range) || !wireByValActionHasVBA275 {
				t.Fatalf("ByVal action does not retain the published VBA275 diagnostic: %+v", byValAction)
			}
			if len(byValAction.Edit.Changes) != 0 || len(byValAction.Edit.DocumentChanges) != 1 {
				t.Fatalf("ByVal edit is not a single versioned document change: %+v", byValAction.Edit)
			}
			change, ok := byValAction.Edit.DocumentChanges[0].(protocol.TextDocumentEdit)
			if !ok || change.TextDocument.URI != protocol.DocumentUri(uri) || change.TextDocument.Version == nil || *change.TextDocument.Version != 41 || len(change.Edits) != 1 {
				t.Fatalf("versioned ByVal document edit = %#v", byValAction.Edit.DocumentChanges)
			}
			edit, ok := change.Edits[0].(protocol.TextEdit)
			if !ok || edit.Range != selection || edit.NewText != "ByVal" {
				t.Fatalf("ByVal edit = %#v, want replacement of the published ByRef header range", change.Edits[0])
			}
			gotDiskSource, err := os.ReadFile(path)
			if err != nil {
				t.Fatal(err)
			}
			if string(gotDiskSource) != diskSource {
				t.Fatalf("code action changed disk source to %q, want unchanged %q", string(gotDiskSource), diskSource)
			}
		})
	}
}

func newByValProjectTestServer(t *testing.T, source string) (*Server, intel.Document, func()) {
	t.Helper()
	cfg := config.Default()
	cfg.Analyze.DetectByRefParametersCanBeByVal = true
	root := t.TempDir()
	path := filepath.Join(root, cfg.Src.Modules, "Main.bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(path, []byte(source), 0o600); err != nil {
		t.Fatal(err)
	}
	server, cleanup, err := New(Options{RootDir: root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	if err := server.analysis.waitReady(); err != nil {
		cleanup()
		t.Fatal(err)
	}
	doc, err := server.docs.open(pathToFileURI(path), source, 1)
	if err != nil {
		cleanup()
		t.Fatal(err)
	}
	if server.documentKind(doc) != DocumentKindVBA || !strings.EqualFold(doc.ModuleKind, "standard") {
		cleanup()
		t.Fatalf("source-root document classification = (%v, %q), want standard VBA module", server.documentKind(doc), doc.ModuleKind)
	}
	return server, doc, cleanup
}
